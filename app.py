"""
眼鏡店7店舗向け シフト作成・最適化Webアプリ
Streamlit + Google OR-Tools (CP-SAT)

起動: streamlit run app.py
"""

from __future__ import annotations

import datetime as dt

import pandas as pd
import streamlit as st

import utils
from optimizer import solve_shift

st.set_page_config(page_title="眼鏡店シフト最適化", layout="wide")


# ---------------------------------------------------------------------------
# セッション初期化
# ---------------------------------------------------------------------------
def init_state():
    if "staff_df" not in st.session_state:
        st.session_state.staff_df = utils.default_staff_df()
    if "requests_df" not in st.session_state:
        # ローカル保存データがあれば復元し、なければ空の雛形を表示する。
        st.session_state.requests_df = utils.load_requests_from_disk()
    if "solve_result" not in st.session_state:
        st.session_state.solve_result = None
    if "period" not in st.session_state:
        st.session_state.period = (2026, 9)
    if "special_closure_labels" not in st.session_state:
        st.session_state.special_closure_labels = []


init_state()


def recompute_allowed_stores(staff_df: pd.DataFrame) -> pd.DataFrame:
    """emp_type・部署固定ルール・他店舗ヘルプ許可フラグから allowed_stores を再計算する。

    店長は「他店舗ヘルプ許可」チェックボックスの値に関わらず、常に can_help=False・
    home_store 1店舗のみに強制する（ハード制約の二重保証）。パート、および
    custom_allowed_stores(若松・竹内など個別に勤務可能店舗が限定されたスタッフ)は、
    主所属店舗がそれぞれの勤務可能店舗と矛盾しないよう補正する。
    """
    df = staff_df.copy()

    df.loc[df["emp_type"] == "店長", "can_help"] = False

    def _fix_home_store(row):
        if row["emp_type"] == "パート" and row.get("part_role"):
            allowed = utils.PART_ROLE_ALLOWED_STORES.get(row["part_role"], [])
            if allowed and row["home_store"] not in allowed:
                return allowed[0]
        custom = row.get("custom_allowed_stores")
        if isinstance(custom, list) and len(custom) > 0 and row["home_store"] not in custom:
            return custom[0]
        return row["home_store"]

    df["home_store"] = df.apply(_fix_home_store, axis=1)
    df["allowed_stores"] = df.apply(utils.derive_allowed_stores, axis=1)
    return df


# ---------------------------------------------------------------------------
# サイドバー
# ---------------------------------------------------------------------------
st.sidebar.title("設定")

year = st.sidebar.number_input("対象年", min_value=2020, max_value=2100, value=st.session_state.period[0])
month = st.sidebar.selectbox(
    "対象月",
    options=list(range(1, 13)),
    index=st.session_state.period[1] - 1,
    format_func=lambda m: f"{m}月度",
)
st.session_state.period = (year, month)

period_dates = utils.get_period_dates(year, month)
period_label = f"{year}年{month}月度 ({period_dates[0].month}/{period_dates[0].day}〜{period_dates[-1].month}/{period_dates[-1].day})"

# 特別休業日(お盆・年末年始等)の選択は Tab1 のウィジェットで行うが、
# dates_df の計算(このすぐ下)より前に値を読む必要があるため、ここでは
# 前回までにセッションへ保存された選択値を読み出すだけにする(Streamlitの
# ウィジェットは key を通じて session_state と自動同期されるため、この
# パターンで安全に「後で描画するウィジェットの値を先に使う」ことができる)。
special_closure_labels_all = [
    f"{d.month}/{d.day}({wd})" for d, wd in zip(period_dates, [utils.WEEKDAY_JP[d.weekday()] for d in period_dates])
]
special_closure_label_to_date = dict(zip(special_closure_labels_all, period_dates))
st.session_state.special_closure_labels = [
    l for l in st.session_state.special_closure_labels if l in special_closure_label_to_date
]
special_closure_dates = [special_closure_label_to_date[l] for l in st.session_state.special_closure_labels]

dates_df = utils.classify_days(period_dates, special_closure_dates=special_closure_dates)
st.sidebar.markdown(f"**シフト期間:** {period_label}")

closed_days_count_sidebar = int((~dates_df["is_business_day"]).sum())
special_closure_count_sidebar = int(dates_df["is_special_closure"].sum())
base_holiday_quota = closed_days_count_sidebar + 1
base_workdays_sidebar = len(period_dates) - base_holiday_quota
st.sidebar.metric(
    "店長・正社員 基準公休日数(自動計算)",
    f"{base_holiday_quota}日",
    help=(
        f"総定休日数({closed_days_count_sidebar}日 ※特別休業日{special_closure_count_sidebar}日を含む)"
        f" ＋ 個別公休1日 = {base_holiday_quota}日。基準出勤日数は "
        f"{len(period_dates)}日 − {base_holiday_quota}日 = {base_workdays_sidebar}日です。"
    ),
)
time_limit_sec = st.sidebar.slider("最適化 計算時間上限(秒)", min_value=10, max_value=180, value=30, step=10)

st.sidebar.markdown("---")
run_clicked = st.sidebar.button("🚀 最適化を実行", type="primary", width="stretch")

st.sidebar.markdown("---")
st.sidebar.caption(
    "社長は現場出勤ゼロが絶対要件のため、スタッフマスタ・シフト変数の"
    "いずれにも一切登場しません（そもそもデータとして保持していません）。"
)

if run_clicked:
    staff_df = recompute_allowed_stores(st.session_state.staff_df)
    with st.spinner("シフトを最適化しています…"):
        result = solve_shift(
            dates_df=dates_df,
            staff_df=staff_df,
            requests_df=st.session_state.requests_df,
            base_holiday_quota=base_holiday_quota,
            time_limit_sec=time_limit_sec,
        )
    st.session_state.solve_result = result
    if result.is_feasible:
        st.sidebar.success(f"完了: {result.status_name} (不足アラート {len(result.shortages)}件)")
    else:
        st.sidebar.error(f"求解失敗: {result.status_name}")


# ---------------------------------------------------------------------------
# タブ
# ---------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    ["📅 カレンダー・有給可能日設定", "👥 スタッフ・希望休設定", "📋 シフト結果・警告", "📤 エクスポート"]
)

# --- Tab1: カレンダー・有給可能日設定 ---------------------------------------
with tab1:
    st.subheader("特別休業日（臨時定休日）の設定")
    st.caption(
        "お盆・年末年始など、全店一斉で臨時休業とする日を指定できます。指定した日は"
        "強制的に全スタッフ公休となり、店舗の営業枠(必要人数)も0名になります。"
        "定休日数にカウントされ、正社員・店長の基準公休日数・基準出勤日数に"
        "自動的に反映されます。"
    )
    st.multiselect(
        "特別休業日（全店一斉休業 / お盆・年末年始等）",
        options=special_closure_labels_all,
        key="special_closure_labels",
    )

    st.markdown("---")
    st.subheader(f"営業日一覧 — {period_label}")
    show_df = dates_df.copy()
    show_df["date"] = show_df["date"].apply(lambda d: f"{d.month}/{d.day}")
    st.dataframe(
        show_df[["date", "weekday_jp", "day_type", "hours", "note"]].rename(
            columns={"date": "日付", "weekday_jp": "曜日", "day_type": "区分", "hours": "営業時間", "note": "備考"}
        ),
        width="stretch",
        height=320,
        hide_index=True,
    )

    n_closed = int((~dates_df["is_business_day"]).sum())
    n_business = int(dates_df["is_business_day"].sum())
    n_short = int((dates_df["day_type"] == "特別営業(短縮)").sum())
    n_special = int(dates_df["is_special_closure"].sum())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("通常営業日数", n_business - n_short)
    c2.metric("特別営業(短縮)日数", n_short)
    c3.metric("定休日数(曜日定休)", n_closed - n_special)
    c4.metric("特別休業日数", n_special)

    st.markdown("---")
    st.subheader("有給休暇 取得可能日・人数枠の自動算出")
    st.caption(
        "各営業日について、全7店舗の最低運営人員(社員のみ換算)を確保した上での"
        "「社員・嘱託の余力人数」を自動算出します。既に絶対休/有休確定が入っている"
        "スタッフは差し引いて計算されます。"
    )

    staff_df_current = recompute_allowed_stores(st.session_state.staff_df)
    availability_df = utils.compute_paid_leave_availability(
        dates_df, staff_df_current, st.session_state.requests_df
    )

    disp = availability_df.copy()
    disp["date"] = disp["date"].apply(lambda d: f"{d.month}/{d.day}")
    disp = disp.rename(
        columns={
            "date": "日付",
            "weekday_jp": "曜日",
            "day_type": "区分",
            "employee_headcount": "社員+嘱託 総数",
            "already_confirmed_off": "既存の絶対休/有休確定",
            "min_required_employees": "最低必要人員",
            "paid_leave_slots": "有休取得可能枠",
        }
    )

    def _highlight_slots(row):
        color = "background-color: #d4edda" if row["有休取得可能枠"] > 0 else ""
        return [color] * len(row)

    st.dataframe(
        disp.style.apply(_highlight_slots, axis=1),
        width="stretch",
        height=320,
        hide_index=True,
    )

    if st.button("📝 スタッフ向け案内文を生成"):
        text = utils.generate_paid_leave_announcement(availability_df, period_label)
        st.session_state["announcement_text"] = text

    if "announcement_text" in st.session_state:
        st.text_area("案内文（コピーしてご利用ください）", st.session_state["announcement_text"], height=280)

    st.markdown("---")
    st.subheader("有休申請をシフトに反映（希望休一覧へ追加）")
    slot_days = availability_df[availability_df["paid_leave_slots"] > 0]["date"].tolist()
    if slot_days:
        colA, colB, colC = st.columns([2, 2, 1])
        with colA:
            target_name = st.selectbox(
                "対象スタッフ",
                options=staff_df_current[staff_df_current["emp_type"].isin(utils.EMPLOYEE_TYPES)]["name"].tolist(),
                key="leave_target_name",
            )
        with colB:
            target_date = st.selectbox(
                "有休取得可能日", options=slot_days, format_func=lambda d: f"{d.month}/{d.day}", key="leave_target_date"
            )
        with colC:
            st.write("")
            st.write("")
            if st.button("追加"):
                sid = staff_df_current.loc[staff_df_current["name"] == target_name, "staff_id"].iloc[0]
                new_row = {"staff_id": sid, "name": target_name, "date": target_date, "kind": "有休申請"}
                st.session_state.requests_df = pd.concat(
                    [st.session_state.requests_df, pd.DataFrame([new_row])], ignore_index=True
                )
                st.success(f"{target_name} / {target_date.month}/{target_date.day} を有休申請として追加しました。")
    else:
        st.info("現在、有休取得可能枠のある日はありません。")


# --- Tab2: スタッフ・希望休設定 ---------------------------------------------
with tab2:
    st.subheader("希望休・有休 CSV一括インポート")
    st.caption(
        "「スタッフ名, 日付, 種別」の縦持ち3列形式、または「スタッフ名×日付」の"
        "マトリクス形式のいずれのCSVも読み込めます。値は 希望休(または休)/有休/絶対休 "
        "のいずれかを記入してください（空欄・0・-は指定なし扱いです）。"
    )
    template_bytes = utils.generate_requests_csv_template(st.session_state.staff_df, dates_df)
    col_up1, col_up2 = st.columns([2, 1])
    with col_up1:
        uploaded_csv = st.file_uploader("希望休・有休CSVをアップロード", type=["csv"], key="requests_csv_uploader")
    with col_up2:
        st.write("")
        st.download_button(
            "⬇️ 入力用テンプレートCSVをダウンロード",
            data=template_bytes,
            file_name=f"kyuka_template_{year}{month:02d}.csv",
            mime="text/csv",
        )

    if uploaded_csv is not None:
        clear_first = st.checkbox(
            "取り込み前に既存の希望休・有休データをすべてクリアする"
            "（チェックなしの場合は同一スタッフ・同一日付のみ上書きしてマージします）",
            value=False,
            key="csv_clear_first",
        )
        if st.button("📥 この内容を取り込む", key="csv_import_button"):
            try:
                new_rows_df, csv_warnings = utils.parse_requests_csv(
                    uploaded_csv.getvalue(), st.session_state.staff_df, dates_df
                )
            except Exception as e:
                st.error(f"CSVの解析中にエラーが発生しました: {e}")
                new_rows_df, csv_warnings = pd.DataFrame(), []

            for w in csv_warnings:
                st.warning(w)

            if new_rows_df.empty:
                st.error("有効な希望休・有休データを1件も読み取れませんでした。CSVの内容をご確認ください。")
            else:
                existing = st.session_state.requests_df
                if clear_first:
                    existing = existing.iloc[0:0]
                elif not existing.empty:
                    new_keys = set(zip(new_rows_df["staff_id"], new_rows_df["date"]))
                    keep_mask = ~existing.apply(lambda r: (r["staff_id"], r["date"]) in new_keys, axis=1)
                    existing = existing[keep_mask]
                st.session_state.requests_df = pd.concat([existing, new_rows_df], ignore_index=True)
                st.success(f"{len(new_rows_df)} 件の希望休・有休データを取り込みました。")

    st.markdown("---")
    st.subheader("スタッフマスタ")
    st.caption(
        "社員区分・主所属店舗・測定/加工スキル保有・嘱託の最低勤務日数を編集できます。"
        "勤務可能店舗はパートの役割(A〜E)に応じて自動制限されます。"
    )

    st.caption(
        "「他店舗ヘルプ許可」は店長では常にOFF（自店舗完全固定・ハード制約）です。"
        "チェックを入れても店長行は保存時に自動的にOFFへ戻ります。"
    )
    st.caption(
        "若松・若林・中村は勤務可能店舗が個別に限定されており（若松: 新蟹江店/"
        "大治店/稲沢店/名古屋中川店、若林: 極楽店/徳重店/天白植田店、中村: 稲沢店以外の"
        "6店舗）、「他店舗ヘルプ許可」チェックのON/OFFに関わらずこの限定リストが"
        "ハード制約として優先されます（竹内は全7店舗ヘルプ可能）。"
    )
    st.caption(
        "「個別公休希望日」は店長・正社員11名のみに適用されます（定休日を除いた"
        "基準公休の残り1日分）。日付を指定するとその日は確定公休(出勤なし)として"
        "扱われ、他店舗ヘルプで店舗体制を維持します。「未指定」のままにすると、"
        "全社体制のバランスを考慮してAIが最も無理のない1日を自動で割り当てます。"
    )

    def _range_label(row):
        if row["emp_type"] == "嘱託":
            if pd.notna(row["min_workdays"]) and pd.notna(row.get("max_workdays")):
                return f"{int(row['min_workdays'])}〜{int(row['max_workdays'])}日"
            return f"{int(row['min_workdays'])}日以上" if pd.notna(row["min_workdays"]) else "-"
        if row["emp_type"] == "パート":
            lo, hi = utils.PART_ROLE_WORKDAY_RANGE.get(row["part_role"], (None, None))
            return f"{lo}〜{hi}日" if lo is not None else "-"
        return "-"

    # 個別公休希望日は「日付一覧 or なし」のドロップダウンで選択させる
    # (カレンダー入力ではなく、定休日を誤って選べない選択式にするため)。
    business_day_rows = dates_df[dates_df["is_business_day"]]
    off_date_labels = [
        f"{d.month}/{d.day}({wd})" for d, wd in zip(business_day_rows["date"], business_day_rows["weekday_jp"])
    ]
    label_to_date = dict(zip(off_date_labels, business_day_rows["date"]))
    date_to_label = {v: k for k, v in label_to_date.items()}
    UNSPECIFIED = "未指定（なし）"
    off_date_options = [UNSPECIFIED] + off_date_labels

    edit_df = st.session_state.staff_df.copy()
    edit_df["勤務可能店舗"] = edit_df["allowed_stores"].apply(lambda lst: "、".join(lst))
    edit_df["稼働日数ルール(自動)"] = edit_df.apply(_range_label, axis=1)
    edit_df["個別公休希望日"] = edit_df["preferred_off_date"].apply(lambda d: date_to_label.get(d, UNSPECIFIED))

    edited = st.data_editor(
        edit_df[
            [
                "staff_id",
                "name",
                "emp_type",
                "home_store",
                "part_role",
                "has_skill",
                "can_help",
                "min_workdays",
                "max_workdays",
                "個別公休希望日",
                "勤務可能店舗",
                "稼働日数ルール(自動)",
            ]
        ],
        column_config={
            "staff_id": st.column_config.TextColumn("ID", disabled=True),
            "name": st.column_config.TextColumn("氏名"),
            "emp_type": st.column_config.SelectboxColumn(
                "役職", options=["店長", "正社員", "嘱託", "パート"], disabled=True
            ),
            "home_store": st.column_config.SelectboxColumn("主所属店舗", options=utils.STORES),
            "part_role": st.column_config.TextColumn("パート役割", disabled=True),
            "has_skill": st.column_config.CheckboxColumn("測定・加工スキル"),
            "can_help": st.column_config.CheckboxColumn("他店舗ヘルプ許可"),
            "稼働日数ルール(自動)": st.column_config.TextColumn("稼働日数ルール(自動)", disabled=True),
            "min_workdays": st.column_config.NumberColumn("嘱託:最低勤務日数", min_value=0, max_value=31),
            "max_workdays": st.column_config.NumberColumn("嘱託:最高勤務日数", min_value=0, max_value=31),
            "個別公休希望日": st.column_config.SelectboxColumn(
                "個別公休希望日(店長/正社員)", options=off_date_options
            ),
            "勤務可能店舗": st.column_config.TextColumn("勤務可能店舗(自動)", disabled=True),
        },
        hide_index=True,
        width="stretch",
        num_rows="fixed",
        key="staff_editor",
    )

    merged = st.session_state.staff_df.copy()
    for col in ["name", "home_store", "has_skill", "can_help", "min_workdays", "max_workdays"]:
        merged[col] = edited[col].values
    merged["preferred_off_date"] = edited["個別公休希望日"].map(lambda lbl: label_to_date.get(lbl))
    st.session_state.staff_df = recompute_allowed_stores(merged)

    st.markdown("---")
    st.subheader("希望休・有休 入力テーブル")
    st.caption(
        "種別: 希望休(ソフト) / 絶対休(ハード=100%遵守) / 有休申請(ソフト・有休可能日で優先) / 有休確定(ハード)"
        "　※ 手入力・編集内容はローカル(data/saved_kyuka.csv)へ即時自動保存され、"
        "再起動や別ブラウザでのアクセス時にも復元されます。"
    )

    if st.button("🗑️ 休暇データをリセット", key="reset_requests_button"):
        utils.clear_saved_requests()
        st.session_state.requests_df = utils.default_requests_df()
        st.success("希望休・有休データをリセットしました。")
        st.rerun()

    req_df = st.session_state.requests_df.copy()
    if req_df.empty:
        req_df = pd.DataFrame(
            {"staff_id": [None], "name": [None], "date": [period_dates[0]], "kind": ["希望休"]}
        )

    name_options = st.session_state.staff_df["name"].tolist()
    id_by_name = dict(zip(st.session_state.staff_df["name"], st.session_state.staff_df["staff_id"]))

    req_edited = st.data_editor(
        req_df[["name", "date", "kind"]],
        column_config={
            "name": st.column_config.SelectboxColumn("スタッフ", options=name_options),
            "date": st.column_config.DateColumn(
                "日付", min_value=period_dates[0], max_value=period_dates[-1], format="YYYY-MM-DD"
            ),
            "kind": st.column_config.SelectboxColumn("種別", options=utils.REQUEST_KINDS),
        },
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key="requests_editor",
    )
    req_edited = req_edited.dropna(subset=["name", "date", "kind"])
    req_edited["staff_id"] = req_edited["name"].map(id_by_name)
    req_edited["date"] = req_edited["date"].apply(lambda d: d if isinstance(d, dt.date) else pd.to_datetime(d).date())
    st.session_state.requests_df = req_edited[["staff_id", "name", "date", "kind"]].reset_index(drop=True)
    # 手入力・編集(および直前のCSVインポート)の内容を即時にローカルへ自動保存する。
    utils.save_requests_to_disk(st.session_state.requests_df)


# --- Tab3: シフト結果・警告 ---------------------------------------------------
with tab3:
    result = st.session_state.solve_result
    if result is None:
        st.info("サイドバーの「🚀 最適化を実行」を押すとここに結果が表示されます。")
    elif not result.is_feasible:
        st.error(f"求解できませんでした（ステータス: {result.status_name}）。制約設定を見直してください。")
    else:
        st.subheader("不足アラート")
        if result.shortages:
            st.warning(f"{len(result.shortages)} 件の人員不足があります。")
            weekday_by_date = dict(zip(dates_df["date"], dates_df["weekday_jp"]))
            for s in result.shortages:
                d = s["date"]
                wd = weekday_by_date.get(d, "")
                label = "人員不足" if s["内容"] == "不足人数" else s["内容"]
                st.warning(
                    f"⚠️ {d.month}月{d.day}日({wd}) {s['store']}：{label} "
                    f"{s['不足数']}名（{s['原因候補']}）"
                )
            shortage_disp = pd.DataFrame(result.shortages).copy()
            shortage_disp["date"] = shortage_disp["date"].apply(lambda d: f"{d.month}/{d.day}")
            shortage_disp = shortage_disp.rename(
                columns={"date": "日付", "store": "店舗", "不足数": "不足数", "原因候補": "原因候補(休み希望)"}
            )
            with st.expander("不足アラートの一覧（表形式）を表示"):
                st.dataframe(shortage_disp, width="stretch", hide_index=True)
        else:
            st.success("人員不足はありません。全店舗・全日で必要体制を充足しています。")

        st.markdown("---")
        st.subheader("店舗別シフト表")
        shift_df = result.shift_df
        if shift_df.empty:
            st.info("出勤データがありません。")
        else:
            store_sel = st.selectbox("店舗を選択", options=utils.STORES, key="store_view")
            sub = shift_df[shift_df["store"] == store_sel].copy()
            sub["date"] = sub["date"].apply(lambda d: f"{d.month}/{d.day}")
            pivot = sub.groupby("date")["name"].apply(lambda s: "、".join(s)).reset_index()
            pivot.columns = ["日付", "出勤スタッフ"]
            st.dataframe(pivot, width="stretch", hide_index=True, height=400)

            st.markdown("---")
            st.subheader("スタッフ別カレンダー")
            staff_sel = st.selectbox("スタッフを選択", options=st.session_state.staff_df["name"].tolist(), key="staff_view")
            sid_sel = st.session_state.staff_df.loc[st.session_state.staff_df["name"] == staff_sel, "staff_id"].iloc[0]
            cal_rows = []
            work_map = {r["date"]: r["store"] for _, r in shift_df[shift_df["staff_id"] == sid_sel].iterrows()}
            for _, day in dates_df.iterrows():
                d = day["date"]
                if not day["is_business_day"]:
                    status = "定休日"
                elif d in work_map:
                    status = work_map[d]
                else:
                    status = "休み"
                cal_rows.append({"日付": f"{d.month}/{d.day}", "曜日": day["weekday_jp"], "状態": status})
            st.dataframe(pd.DataFrame(cal_rows), width="stretch", hide_index=True, height=400)

        st.markdown("---")
        st.subheader("スタッフ別サマリー")
        st.dataframe(result.staff_summary_df, width="stretch", hide_index=True)

        st.markdown("---")
        st.subheader("有休取得可能枠（最適化結果ベース）")
        st.caption(
            "確定したシフト配置をもとに、各店舗の成立条件(平日/土日祝別の必要人数・"
            "名古屋中川店/天白植田店/徳重店の土日祝は社員2名以上必須・"
            "測定加工スキル保有者1名以上)を崩さずに、追加で休暇へ回せる人数を"
            "営業日ごとに算出しています。"
        )
        post_availability_df = utils.compute_post_solve_leave_availability(
            result.shift_df, st.session_state.staff_df, dates_df, st.session_state.requests_df
        )
        st.session_state["post_availability_df"] = post_availability_df

        badge_cols = st.columns(4)
        for i, (_, row) in enumerate(post_availability_df.iterrows()):
            d = row["date"]
            n = int(row["extra_leave_slots"])
            label = f"{d.month}/{d.day}({row['weekday_jp']})"
            badge = utils.leave_slot_badge(n)
            with badge_cols[i % 4]:
                if n > 0:
                    st.success(f"{label}: {badge}")
                else:
                    st.caption(f"{label}: {badge}")

        with st.expander("詳細（間引き候補の内訳）を表示"):
            detail_disp = post_availability_df.copy()
            detail_disp["date"] = detail_disp["date"].apply(lambda d: f"{d.month}/{d.day}")
            detail_disp = detail_disp.rename(
                columns={
                    "date": "日付",
                    "weekday_jp": "曜日",
                    "day_type": "区分",
                    "extra_leave_slots": "追加取得可能枠",
                    "detail": "内訳(店舗:候補者)",
                }
            )
            st.dataframe(detail_disp, width="stretch", hide_index=True, height=320)

        if st.button("📝 社内案内文を生成（有休 追加取得可能日）"):
            post_text = utils.generate_post_solve_leave_announcement(post_availability_df, period_label)
            st.session_state["post_announcement_text"] = post_text

        if "post_announcement_text" in st.session_state:
            st.caption("下記をコピーしてLINE・社内掲示にご利用ください（右上のコピーアイコンからコピーできます）。")
            st.code(st.session_state["post_announcement_text"], language=None)


# --- Tab4: エクスポート -------------------------------------------------------
with tab4:
    result = st.session_state.solve_result
    if result is None or not result.is_feasible or result.shift_df.empty:
        st.info("先にサイドバーで最適化を実行してください。")
    else:
        st.subheader("Excel(.xlsx) エクスポート")
        st.caption("シート1: 店舗別日別シフト表 / シート2: スタッフ別出勤一覧表")
        workbook_bytes = utils.build_export_workbook(
            result.shift_df, dates_df, st.session_state.staff_df, st.session_state.requests_df
        )
        st.download_button(
            label="⬇️ シフト表をダウンロード (.xlsx)",
            data=workbook_bytes,
            file_name=f"shift_{year}{month:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
