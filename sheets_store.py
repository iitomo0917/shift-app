"""
Googleスプレッドシートを永続化バックエンドとするデータ保存層。

背景:
Streamlit Community Cloud はアプリの再起動・再デプロイのたびに、ローカル
ファイルシステム(dataフォルダ)への書き込み内容を全て消去する(=永続化されない)
仕様であることが確認された。そのため、希望休・有給申請・特別休業日設定・
シフト結果・スタッフ名簿のいずれも、ローカルCSV/JSONではなく、このモジュール
経由でGoogleスプレッドシートへ保存する方式に変更した。

セットアップ(利用者側で1回だけ必要な準備):
  1. Google Cloudプロジェクトを作成し、Google Sheets API・Google Drive APIを有効化
  2. サービスアカウントを作成し、JSON形式の鍵を発行
  3. 保存先にしたいGoogleスプレッドシートを作成し、サービスアカウントの
     メールアドレス(client_email)を編集者として共有
  4. StreamlitのSecrets(st.secrets)に以下の形式で設定する:

     [gcp_service_account]
     type = "service_account"
     project_id = "..."
     private_key_id = "..."
     private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
     client_email = "..."
     client_id = "..."
     auth_uri = "https://accounts.google.com/o/oauth2/auth"
     token_uri = "https://oauth2.googleapis.com/token"
     auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
     client_x509_cert_url = "..."

     spreadsheet_url = "https://docs.google.com/spreadsheets/d/xxxxxxxx/edit"

  (ダウンロードしたJSON鍵の各キーを、そのまま [gcp_service_account] の下に
  TOML形式で転記する。spreadsheet_url は保存先スプレッドシートのURL。)

このモジュールが未設定(Secrets未設定/gspread未インストール)の場合は
SheetsNotConfiguredError を送出する。呼び出し側(utils.py)でこれを捕捉し、
画面にセットアップ手順への案内を表示した上で、安全側(空のデータ)にフォール
バックすること。
"""

from __future__ import annotations

import datetime as dt
import json
from typing import Any

import pandas as pd
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials

    _IMPORT_ERROR: Exception | None = None
except Exception as exc:  # pragma: no cover - 環境にライブラリが無い場合のみ
    gspread = None  # type: ignore[assignment]
    Credentials = None  # type: ignore[assignment]
    _IMPORT_ERROR = exc


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# --- ワークシート(タブ)名・列定義 -------------------------------------------
KYUKA_SHEET = "kyuka_requests"
# utils.KYUKA_LOG_COLUMNS と列名を揃えている(4列目は "updated_at")。
KYUKA_COLUMNS = ["staff_name", "date", "request_type", "updated_at"]

SPECIAL_DAYS_SHEET = "special_days"
SPECIAL_DAYS_COLUMNS = ["year", "month", "date", "category", "reason"]
# category: "special_closure"(特別休業日) または "forced_open"(臨時営業日)

STAFF_SHEET = "staff_roster"
# リスト型・複雑な値を持つ列はJSON文字列として1セルに格納する。
STAFF_JSON_COLUMNS = [
    "allowed_stores",
    "custom_allowed_stores",
    "weekend_holiday_forbidden_stores",
    "saturday_sunday_forbidden_stores",
    "weekend_holiday_avoid_stores",
]
STAFF_DATE_COLUMNS = ["preferred_off_date"]
STAFF_COLUMNS = [
    "staff_id",
    "name",
    "emp_type",
    "home_store",
    "allowed_stores",
    "can_help",
    "custom_allowed_stores",
    "part_role",
    "has_skill",
    "min_workdays",
    "max_workdays",
    "preferred_off_date",
    "weekend_holiday_forbidden_stores",
    "preferred_store",
    "saturday_sunday_forbidden_stores",
    "weekend_holiday_avoid_stores",
]

SHIFT_RESULT_SHEET = "shift_result"
SHIFT_RESULT_COLUMNS = ["date", "store", "staff_id", "name", "emp_type"]

SHIFT_META_SHEET = "shift_result_meta"
SHIFT_META_COLUMNS = ["key", "value"]


class SheetsNotConfiguredError(RuntimeError):
    """Secrets未設定、またはgspread未インストールのため、Sheetsへ接続できない。"""


def _require_gspread() -> None:
    if gspread is None or Credentials is None:
        raise SheetsNotConfiguredError(
            "gspread / google-auth がインストールされていません。requirements.txt をご確認ください。"
            f"(詳細: {_IMPORT_ERROR})"
        )


@st.cache_resource(show_spinner=False)
def _get_client():
    _require_gspread()
    if "gcp_service_account" not in st.secrets:
        raise SheetsNotConfiguredError(
            "Secretsに [gcp_service_account] が設定されていません。セットアップ手順に従って設定してください。"
        )
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def _get_spreadsheet():
    if "spreadsheet_url" not in st.secrets:
        raise SheetsNotConfiguredError(
            "Secretsに spreadsheet_url が設定されていません。セットアップ手順に従って設定してください。"
        )
    client = _get_client()
    return client.open_by_url(st.secrets["spreadsheet_url"])


def is_configured() -> bool:
    """Secrets設定・ライブラリの両方が揃っているかどうかを、接続を試みずに判定する。"""
    if gspread is None or Credentials is None:
        return False
    try:
        return "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets
    except Exception:
        return False


def _get_or_create_worksheet(title: str, columns: list[str]):
    sheet = _get_spreadsheet()
    try:
        ws = sheet.worksheet(title)
    except gspread.WorksheetNotFound:
        try:
            ws = sheet.add_worksheet(title=title, rows=2000, cols=max(10, len(columns)))
            ws.append_row(columns)
            return ws
        except Exception:
            # 複数のブラウザ/端末からほぼ同時にアクセスした場合、双方が「まだ
            # シートが無い」と判断して同時に作成しようとすることがある。この
            # 場合Google側が名前の重複を避けて "xxx_conflictNNN" のような別名の
            # シートを作ってしまうことがあるため、まず既存のシートが実は
            # (直前の競合相手により)既に作られていないか再確認してから諦める。
            try:
                return sheet.worksheet(title)
            except gspread.WorksheetNotFound:
                raise
    values = ws.get_all_values()
    if not values:
        ws.append_row(columns)
    return ws


def _records_to_df(ws, columns: list[str]) -> pd.DataFrame:
    values = ws.get_all_values()
    if len(values) <= 1:
        return pd.DataFrame(columns=columns)
    header, rows = values[0], values[1:]
    df = pd.DataFrame(rows, columns=header)
    return df.reindex(columns=columns)


# ---------------------------------------------------------------------------
# 希望休・有給申請ログ(追記型・全期間共通の1シート)
# ---------------------------------------------------------------------------

def append_kyuka_request(staff_name: str, date: dt.date, request_type: str) -> None:
    ws = _get_or_create_worksheet(KYUKA_SHEET, KYUKA_COLUMNS)
    ws.append_row(
        [
            staff_name,
            date.isoformat() if isinstance(date, dt.date) else str(date),
            request_type,
            dt.datetime.now().isoformat(timespec="seconds"),
        ],
        value_input_option="RAW",
    )
    # 書き込み直後は必ず最新内容を読み直せるよう、読み込みキャッシュを破棄する。
    load_kyuka_log.clear()


@st.cache_data(ttl=8, show_spinner=False)
def load_kyuka_log() -> pd.DataFrame:
    """ログの全行を読み込む。結果は数秒間キャッシュする。

    画面操作(ボタン押下・ウィジェット変更)のたびにStreamlitのスクリプトが
    再実行され、その都度この関数が呼ばれるため、キャッシュ無しだと短時間に
    大量のGoogle Sheets API呼び出しが発生し、レート制限エラー
    (gspread.exceptions.APIError)の原因になっていた。ttl秒以内の再読み込みは
    キャッシュされた結果を返すことでAPI呼び出し回数を抑える。書き込み直後は
    append_kyuka_request/clear_kyuka_requests_in_range側でキャッシュを明示的に
    破棄しているため、保存した内容が反映されないことはない。
    """
    ws = _get_or_create_worksheet(KYUKA_SHEET, KYUKA_COLUMNS)
    return _records_to_df(ws, KYUKA_COLUMNS)


def clear_kyuka_requests_in_range(active_requests: list[tuple[str, dt.date]]) -> None:
    """指定された(スタッフ名, 日付)の組み合わせについて、「取消」ログを追記する。

    Sheetsの行を直接削除するのではなく、既存の追記型ログの仕組みに従い
    「取消」種別を追記することで論理削除する(監査履歴も保持される)。
    """
    for name, date in active_requests:
        append_kyuka_request(name, date, "取消")


# ---------------------------------------------------------------------------
# 特別休業日・臨時営業日の設定(全期間共通の1シート、(year, month, date, category)で一意)
# ---------------------------------------------------------------------------

@st.cache_data(ttl=8, show_spinner=False)
def load_special_days_all() -> pd.DataFrame:
    """全期間分の特別休業日・臨時営業日設定を読み込む。結果は数秒間キャッシュする
    (理由はload_kyuka_logのコメントを参照。app.py側でこの関数が画面操作のたびに
    毎回呼ばれるため、キャッシュしないとAPIレート制限に達しやすい)。
    """
    ws = _get_or_create_worksheet(SPECIAL_DAYS_SHEET, SPECIAL_DAYS_COLUMNS)
    return _records_to_df(ws, SPECIAL_DAYS_COLUMNS)


def save_special_days(
    year: int,
    month: int,
    special_closure_map: dict[dt.date, str],
    forced_open_map: dict[dt.date, str],
) -> None:
    """指定月度分の特別休業日・臨時営業日の設定を書き換える(その月度分のみ全体置換)。"""
    ws = _get_or_create_worksheet(SPECIAL_DAYS_SHEET, SPECIAL_DAYS_COLUMNS)
    # ここは保存(マージ)のための読み込みなので、キャッシュを経由せず常に最新を取得する。
    df = _records_to_df(ws, SPECIAL_DAYS_COLUMNS)
    if not df.empty:
        keep = df[~((df["year"] == str(year)) & (df["month"] == str(month)))]
    else:
        keep = df

    new_rows = []
    for d, reason in special_closure_map.items():
        new_rows.append([str(year), str(month), d.isoformat(), "special_closure", reason])
    for d, reason in forced_open_map.items():
        new_rows.append([str(year), str(month), d.isoformat(), "forced_open", reason])

    out_rows = keep.values.tolist() + new_rows
    ws.clear()
    ws.append_row(SPECIAL_DAYS_COLUMNS)
    if out_rows:
        ws.append_rows(out_rows, value_input_option="RAW")
    load_special_days_all.clear()


# ---------------------------------------------------------------------------
# スタッフ名簿(全体を毎回まるごと置換保存する、少量データのシート)
# ---------------------------------------------------------------------------

def _encode_staff_row(row: dict) -> list[str]:
    out = []
    for col in STAFF_COLUMNS:
        val = row.get(col)
        if col in STAFF_JSON_COLUMNS:
            out.append(json.dumps(val if val is not None else [], ensure_ascii=False))
        elif col in STAFF_DATE_COLUMNS:
            out.append(val.isoformat() if isinstance(val, dt.date) else "")
        elif isinstance(val, bool):
            out.append("TRUE" if val else "FALSE")
        elif val is None:
            out.append("")
        else:
            out.append(str(val))
    return out


def _decode_staff_row(row: dict) -> dict:
    out: dict[str, Any] = {}
    for col in STAFF_COLUMNS:
        raw = row.get(col, "")
        if col in STAFF_JSON_COLUMNS:
            try:
                out[col] = json.loads(raw) if raw else []
            except (TypeError, ValueError):
                out[col] = []
        elif col in STAFF_DATE_COLUMNS:
            out[col] = dt.date.fromisoformat(raw) if raw else None
        elif col == "has_skill" or col == "can_help":
            out[col] = str(raw).strip().upper() == "TRUE"
        elif col in ("min_workdays", "max_workdays"):
            out[col] = int(raw) if raw not in ("", None) else None
        else:
            out[col] = raw if raw not in ("", None) else None
    return out


@st.cache_data(ttl=8, show_spinner=False)
def load_staff_roster() -> pd.DataFrame | None:
    """保存済みのスタッフ名簿を読み込む。1件も無い場合はNoneを返す(=呼び出し側でデフォルト値を使う)。

    結果は数秒間キャッシュする(理由はload_kyuka_logのコメントを参照)。
    """
    ws = _get_or_create_worksheet(STAFF_SHEET, STAFF_COLUMNS)
    df = _records_to_df(ws, STAFF_COLUMNS)
    if df.empty:
        return None
    decoded_rows = [_decode_staff_row(r) for _, r in df.iterrows()]
    return pd.DataFrame(decoded_rows, columns=STAFF_COLUMNS)


def save_staff_roster(staff_df: pd.DataFrame) -> None:
    """スタッフ名簿全体を上書き保存する(件数が少ないため全置換方式)。"""
    ws = _get_or_create_worksheet(STAFF_SHEET, STAFF_COLUMNS)
    rows = [_encode_staff_row(r.to_dict()) for _, r in staff_df.iterrows()]
    ws.clear()
    ws.append_row(STAFF_COLUMNS)
    if rows:
        ws.append_rows(rows, value_input_option="RAW")
    load_staff_roster.clear()


# ---------------------------------------------------------------------------
# 最新シフト結果(最適化結果 or 手動編集後の状態、常に最新の1状態のみ保持)
# ---------------------------------------------------------------------------

def save_shift_result(shift_df: pd.DataFrame, meta: dict) -> None:
    ws = _get_or_create_worksheet(SHIFT_RESULT_SHEET, SHIFT_RESULT_COLUMNS)
    out = shift_df.copy()
    if not out.empty:
        out["date"] = out["date"].apply(lambda d: d.isoformat() if isinstance(d, dt.date) else d)
    out = out.reindex(columns=SHIFT_RESULT_COLUMNS).fillna("")
    ws.clear()
    ws.append_row(SHIFT_RESULT_COLUMNS)
    if not out.empty:
        ws.append_rows(out.values.tolist(), value_input_option="RAW")

    meta_ws = _get_or_create_worksheet(SHIFT_META_SHEET, SHIFT_META_COLUMNS)
    meta_ws.clear()
    meta_ws.append_row(SHIFT_META_COLUMNS)
    meta_rows = [[k, json.dumps(v, ensure_ascii=False)] for k, v in meta.items()]
    if meta_rows:
        meta_ws.append_rows(meta_rows, value_input_option="RAW")
    load_shift_result.clear()


@st.cache_data(ttl=8, show_spinner=False)
def load_shift_result() -> tuple[pd.DataFrame | None, dict]:
    """結果は数秒間キャッシュする(理由はload_kyuka_logのコメントを参照)。"""
    try:
        ws = _get_or_create_worksheet(SHIFT_RESULT_SHEET, SHIFT_RESULT_COLUMNS)
        df = _records_to_df(ws, SHIFT_RESULT_COLUMNS)
    except Exception:
        return None, {}
    if df.empty:
        return None, {}
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    df = df.dropna(subset=["date"])

    meta: dict = {}
    try:
        meta_ws = _get_or_create_worksheet(SHIFT_META_SHEET, SHIFT_META_COLUMNS)
        meta_df = _records_to_df(meta_ws, SHIFT_META_COLUMNS)
        for _, r in meta_df.iterrows():
            try:
                meta[r["key"]] = json.loads(r["value"])
            except (TypeError, ValueError, KeyError):
                continue
    except Exception:
        meta = {}
    return df, meta


def clear_shift_result() -> None:
    ws = _get_or_create_worksheet(SHIFT_RESULT_SHEET, SHIFT_RESULT_COLUMNS)
    ws.clear()
    ws.append_row(SHIFT_RESULT_COLUMNS)
    meta_ws = _get_or_create_worksheet(SHIFT_META_SHEET, SHIFT_META_COLUMNS)
    meta_ws.clear()
    meta_ws.append_row(SHIFT_META_COLUMNS)
    load_shift_result.clear()

