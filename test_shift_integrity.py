"""
シフト最適化アプリの網羅的な整合性検証・バグチェックスクリプト。

optimizer.py / utils.py / app.py が実装している全社シフト要件・制約ルールを、
実際にソルバーを実行した結果に対して横断的に検証する。単体テストフレームワーク
(pytest等)を使わず、単一スクリプトとして `python test_shift_integrity.py` で
実行でき、NG/警告となった項目を一覧で報告する。

実行方法:
    python test_shift_integrity.py
"""

from __future__ import annotations

import datetime as dt
import os
import tempfile

import pandas as pd

import utils
from optimizer import solve_shift

RESULTS: list[tuple[str, bool, str]] = []  # (check name, passed, detail)


def record(name: str, issues: list[str]) -> None:
    passed = len(issues) == 0
    detail = "OK" if passed else "; ".join(issues[:20]) + (f" ...(他{len(issues)-20}件)" if len(issues) > 20 else "")
    RESULTS.append((name, passed, detail))


def setup_default_scenario():
    """デフォルトの19名スタッフ・2026年9月度・希望休なしで最適化を実行する。"""
    dates = utils.get_period_dates(2026, 9)
    dates_df = utils.classify_days(dates)
    staff_df = utils.default_staff_df()
    requests_df = utils.default_requests_df()
    closed_days = int((~dates_df["is_business_day"]).sum())
    base_holiday_quota = closed_days + 1
    result = solve_shift(
        dates_df=dates_df,
        staff_df=staff_df,
        requests_df=requests_df,
        base_holiday_quota=base_holiday_quota,
        time_limit_sec=90,
    )
    return dates_df, staff_df, requests_df, result, base_holiday_quota


# ---------------------------------------------------------------------------
# 1. 正社員・店長の出勤日数(基準出勤日数の厳守)
# ---------------------------------------------------------------------------
def check_seishain_workdays(result, staff_df, dates_df, base_holiday_quota) -> list[str]:
    """期待出勤日数は「期間の総日数 − 基準公休日数」で動的に算出する。

    基準公休日数(=総定休日数+個別公休1日)は暦月ごとの曜日配置(水曜・火曜の
    出現回数)によって変動するため、"21日"のような固定値ではなく、対象期間の
    実際のカレンダーから導出した期待値と比較する(2026年9月度はたまたま21日に
    なるが、これは一般則ではない)。
    """
    issues = []
    expected_workdays = len(dates_df) - base_holiday_quota
    salaried = staff_df[staff_df["emp_type"].isin(["店長", "正社員"])]
    for _, row in salaried.iterrows():
        cnt = len(result.shift_df[result.shift_df["staff_id"] == row["staff_id"]])
        if cnt != expected_workdays:
            issues.append(f"{row['name']}: 出勤{cnt}日(期待値{expected_workdays}日)")
    return issues


# ---------------------------------------------------------------------------
# 2. 店舗成立要件・スキル制約
# ---------------------------------------------------------------------------
def check_store_requirements(result, staff_df, dates_df):
    alerts = utils.check_manual_shift_alerts(result.shift_df, staff_df, dates_df)
    known_exempt = {(s["date"], s["store"]) for s in utils.MANUAL_STORE_ASSIGNMENTS if s.get("exclude_others")}

    real_shortages = [
        f"{s['date']} {s['store']}: {s['詳細']}(不足{s['不足数']})"
        for s in alerts["shortages"]
        if (s["date"], s["store"]) not in known_exempt
    ]
    skill_issues = [f"{s['date']} {s['store']}: スキル保有者不在" for s in alerts["skill_issues"]]
    duplicates = [
        f"{d['date']} {d['name']}: 重複配置({'、'.join(d['stores'])})" for d in alerts["duplicates"]
    ]
    return real_shortages, skill_issues, duplicates


def check_nakagawa_weekend(result, staff_df, dates_df) -> list[str]:
    issues = []
    emp_type_by_id = dict(zip(staff_df["staff_id"], staff_df["emp_type"]))
    wh_days = set(dates_df.loc[dates_df["is_business_day"] & dates_df["is_weekend_or_holiday"], "date"])
    naka = result.shift_df[result.shift_df["store"] == "名古屋中川店"]
    for d, grp in naka.groupby("date"):
        if d not in wh_days:
            continue
        e_count = sum(1 for sid in grp["staff_id"] if emp_type_by_id.get(sid) in utils.EMPLOYEE_TYPES)
        if e_count < 2:
            issues.append(f"{d}(土日祝): 社員{e_count}名のみ")
    return issues


# ---------------------------------------------------------------------------
# 3. 個別スタッフの勤務制限
# ---------------------------------------------------------------------------
def check_individual_restrictions(result, staff_df, dates_df) -> list[str]:
    issues = []
    name_to_id = dict(zip(staff_df["name"], staff_df["staff_id"]))
    wh_days = set(dates_df.loc[dates_df["is_business_day"] & dates_df["is_weekend_or_holiday"], "date"])

    wm_id = name_to_id.get("若松")
    if wm_id:
        wm_rows = result.shift_df[result.shift_df["staff_id"] == wm_id]
        bad = wm_rows[wm_rows["date"].isin(wh_days) & wm_rows["store"].isin(["新蟹江店", "名古屋中川店"])]
        for _, r in bad.iterrows():
            issues.append(f"若松: {r['date']} {r['store']}(土日祝配置禁止店舗)")

    wl_id = name_to_id.get("若林")
    if wl_id:
        wl_rows = result.shift_df[result.shift_df["staff_id"] == wl_id]
        bad2 = wl_rows[~wl_rows["store"].isin(["徳重店", "極楽店", "天白植田店"])]
        for _, r in bad2.iterrows():
            issues.append(f"若林: {r['date']} {r['store']}(許可店舗外)")

    nk_id = name_to_id.get("中村")
    if nk_id:
        nk_rows = result.shift_df[result.shift_df["staff_id"] == nk_id]
        bad3 = nk_rows[nk_rows["store"] == "稲沢店"]
        for _, r in bad3.iterrows():
            issues.append(f"中村: {r['date']} 稲沢店(禁止店舗)")

    return issues


def check_wakamatsu_takeuchi_inazawa_priority(result, staff_df) -> list[str]:
    """稲沢店における若松優先配置ルールの検証。

    同一営業日に若松・竹内がともに出勤している場合、「竹内が稲沢店に配置され、
    かつ若松が稲沢店以外に配置される」組み合わせが発生していないことを確認する。
    """
    issues = []
    name_to_id = dict(zip(staff_df["name"], staff_df["staff_id"]))
    wm_id = name_to_id.get("若松")
    tk_id = name_to_id.get("竹内")
    if not wm_id or not tk_id:
        return issues

    wm_rows = result.shift_df[result.shift_df["staff_id"] == wm_id]
    tk_rows = result.shift_df[result.shift_df["staff_id"] == tk_id]
    wm_store_by_date = dict(zip(wm_rows["date"], wm_rows["store"]))
    tk_store_by_date = dict(zip(tk_rows["date"], tk_rows["store"]))

    for d, tk_store in tk_store_by_date.items():
        if tk_store != "稲沢店":
            continue
        wm_store = wm_store_by_date.get(d)
        if wm_store is not None and wm_store != "稲沢店":
            issues.append(f"{d}: 竹内=稲沢店・若松={wm_store}(禁止されるべき組み合わせ)")

    return issues


def check_day_ranges(result, staff_df) -> list[str]:
    """嘱託・パートの契約勤務日数レンジは utils.py の設定値(単一の真実の情報源)と
    照合する(要件文中の数値を再度ハードコードしない。値そのものの妥当性が別途
    問題になる場合は utils.py 側の定数を確認・修正すること)。"""
    issues = []
    for _, row in staff_df.iterrows():
        sid = row["staff_id"]
        cnt = len(result.shift_df[result.shift_df["staff_id"] == sid])
        if row["emp_type"] == "嘱託":
            lo, hi = row["min_workdays"], row["max_workdays"]
            if pd.notna(lo) and pd.notna(hi) and not (lo <= cnt <= hi):
                issues.append(f"{row['name']}(嘱託): 出勤{cnt}日(規定{int(lo)}〜{int(hi)}日)")
        elif row["emp_type"] == "パート":
            lo, hi = utils.PART_ROLE_WORKDAY_RANGE.get(row["part_role"], (None, None))
            if lo is not None and not (lo <= cnt <= hi):
                issues.append(f"{row['name']}(パート): 出勤{cnt}日(規定{lo}〜{hi}日)")
    return issues


# ---------------------------------------------------------------------------
# 4. 有休取得可能枠(余力判定)のロジック整合性
# ---------------------------------------------------------------------------
def check_leave_availability_integrity(result, staff_df, dates_df, requests_df) -> list[str]:
    issues = []
    avail = utils.compute_post_solve_leave_availability(result.shift_df, staff_df, dates_df, requests_df)
    salaried_names = set(staff_df[staff_df["emp_type"].isin(["店長", "正社員"])]["name"])
    total_workdays = result.shift_df.groupby("staff_id").size().to_dict() if not result.shift_df.empty else {}

    for _, row in avail.iterrows():
        detail = row["detail"]
        if not detail:
            continue
        for part in detail.split(" / "):
            if "→" not in part:
                continue
            sub_name = part.split("→")[1].split("代替可")[0]
            if sub_name in salaried_names:
                issues.append(f"{row['date']}: 代替候補に正社員/店長({sub_name})が含まれる")
                continue
            match = staff_df[staff_df["name"] == sub_name]
            if match.empty:
                issues.append(f"{row['date']}: 代替候補「{sub_name}」がスタッフ名簿に存在しない")
                continue
            sub_row = match.iloc[0]
            cnt = total_workdays.get(sub_row["staff_id"], 0)
            if sub_row["emp_type"] == "嘱託":
                max_d = sub_row["max_workdays"]
                if pd.notna(max_d) and cnt >= int(max_d):
                    issues.append(f"{row['date']}: 代替候補{sub_name}は既に上限{int(max_d)}日に到達済み")
            elif sub_row["emp_type"] == "パート":
                _, max_d = utils.PART_ROLE_WORKDAY_RANGE.get(sub_row["part_role"], (0, 999))
                if cnt >= max_d:
                    issues.append(f"{row['date']}: 代替候補{sub_name}は既に上限{max_d}日に到達済み")
    return issues


def check_leave_candidate_excludes_already_off(staff_df, dates_df) -> list[str]:
    """既に希望休/絶対休が確定しているスタッフが、有休候補(除去候補・代替候補の
    いずれ)としても一切出現しないことを、意図的に休み希望を仕込んだ別シナリオで
    検証する(デフォルトシナリオには休み希望が無いため、このチェック専用に再計算する)。
    """
    issues = []
    target = dt.date(2026, 9, 24)
    sid_a = staff_df.loc[staff_df["name"] == "若林", "staff_id"].iloc[0]
    sid_b = staff_df.loc[staff_df["name"] == "若松", "staff_id"].iloc[0]
    requests_df = pd.DataFrame(
        [
            {"staff_id": sid_a, "name": "若林", "date": target, "kind": "希望休"},
            {"staff_id": sid_b, "name": "若松", "date": target, "kind": "絶対休"},
        ]
    )
    dates = utils.get_period_dates(2026, 9)
    dates_df2 = utils.classify_days(dates)
    closed_days = int((~dates_df2["is_business_day"]).sum())
    result = solve_shift(dates_df2, staff_df, requests_df, base_holiday_quota=closed_days + 1, time_limit_sec=90)
    if not result.is_feasible:
        return ["休み希望シナリオの求解に失敗(INFEASIBLE)"]

    works_a = target in result.shift_df[result.shift_df["staff_id"] == sid_a]["date"].tolist()
    works_b = target in result.shift_df[result.shift_df["staff_id"] == sid_b]["date"].tolist()
    if works_a or works_b:
        issues.append(f"休み希望が尊重されていない(若林出勤={works_a}, 若松出勤={works_b})")

    avail = utils.compute_post_solve_leave_availability(result.shift_df, staff_df, dates_df2, requests_df)
    for _, row in avail.iterrows():
        detail = row["detail"]
        if not detail or row["date"] != target:
            continue
        if "若林" in detail or "若松" in detail:
            issues.append(f"{row['date']}: 休み確定済みの若林/若松が候補として出現({detail})")
    return issues


# ---------------------------------------------------------------------------
# 5. データ永続化と手動編集の連動
# ---------------------------------------------------------------------------
def check_requests_persistence() -> list[str]:
    """本番の data/saved_kyuka.csv には一切触れず、一時ディレクトリ上で検証する
    (実運用データを誤って破壊しないため)。"""
    issues = []
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = os.path.join(tmp_dir, "saved_kyuka.csv")
        df = pd.DataFrame(
            [
                {"staff_id": "T01", "name": "生駒", "date": dt.date(2026, 9, 20), "kind": "希望休"},
                {"staff_id": "P01", "name": "尾澤", "date": dt.date(2026, 10, 3), "kind": "有休確定"},
            ]
        )
        utils.save_requests_to_disk(df, path=tmp_path)
        loaded = utils.load_requests_from_disk(path=tmp_path)
        if len(loaded) != 2:
            issues.append(f"保存件数(2件)と復元件数({len(loaded)}件)が一致しない")
        elif not all(isinstance(d, dt.date) for d in loaded["date"]):
            issues.append("復元後のdate列がdatetime.date型になっていない")
        utils.clear_saved_requests(path=tmp_path)
        if len(utils.load_requests_from_disk(path=tmp_path)) != 0:
            issues.append("クリア後も希望休データが残っている")
    return issues


def check_shift_result_persistence() -> list[str]:
    """本番の data/latest_shift_result.csv には一切触れず、一時ディレクトリ上で
    検証する(実運用データを誤って破壊しないため)。"""
    issues = []
    with tempfile.TemporaryDirectory() as tmp_dir:
        shift_path = os.path.join(tmp_dir, "latest_shift_result.csv")
        meta_path = os.path.join(tmp_dir, "latest_shift_meta.json")
        df = pd.DataFrame(
            [{"date": dt.date(2026, 9, 17), "store": "稲沢店", "staff_id": "T01", "name": "生駒", "emp_type": "店長"}]
        )
        meta = {"year": 2026, "month": 9, "status_name": "OPTIMAL", "objective_value": 1.0, "solver_wall_time": 0.1}
        utils.save_shift_result_to_disk(df, meta, shift_path=shift_path, meta_path=meta_path)
        loaded_df, loaded_meta = utils.load_shift_result_from_disk(shift_path=shift_path, meta_path=meta_path)
        if loaded_df is None or len(loaded_df) != 1:
            issues.append("シフト結果の保存/復元件数が一致しない")
        if loaded_meta.get("year") != 2026 or loaded_meta.get("month") != 9:
            issues.append("メタ情報(年月)が正しく復元されない")
        utils.clear_saved_shift_result(shift_path=shift_path, meta_path=meta_path)
        loaded_df2, _ = utils.load_shift_result_from_disk(shift_path=shift_path, meta_path=meta_path)
        if loaded_df2 is not None:
            issues.append("クリア後もシフト結果ファイルが残っている")
    return issues


def check_manual_edit_alert_sync(result, staff_df, dates_df) -> list[str]:
    """手動編集(空欄化→再充足)に対してアラートが正しく連動するかを検証する。"""
    issues = []
    wide = utils.build_manual_shift_wide(result.shift_df, dates_df)
    labels = utils.manual_shift_date_labels(dates_df)
    target_label = labels[0]

    row_idx = wide[(wide["店舗"] == "稲沢店") & (wide["枠"] == 1)].index[0]
    original_name = wide.loc[row_idx, target_label]

    # 空欄化 -> 不足アラートが出るはず
    blanked = wide.copy()
    blanked.loc[row_idx, target_label] = ""
    long_blanked = utils.manual_shift_wide_to_long(blanked, dates_df, staff_df)
    alerts_blanked = utils.check_manual_shift_alerts(long_blanked, staff_df, dates_df)
    target_date = dict(zip(labels, dates_df.loc[dates_df["is_business_day"], "date"]))[target_label]
    if not any(s["store"] == "稲沢店" and s["date"] == target_date for s in alerts_blanked["shortages"]):
        issues.append("空欄化しても稲沢店の不足アラートが検出されない")

    # 再充足 -> 消えるはず
    refilled = blanked.copy()
    refilled.loc[row_idx, target_label] = original_name
    long_refilled = utils.manual_shift_wide_to_long(refilled, dates_df, staff_df)
    alerts_refilled = utils.check_manual_shift_alerts(long_refilled, staff_df, dates_df)
    if any(s["store"] == "稲沢店" and s["date"] == target_date for s in alerts_refilled["shortages"]):
        issues.append("再充足しても稲沢店の不足アラートが消えない")

    # 重複配置検知
    dup = wide.copy()
    oji_idx = dup[(dup["店舗"] == "大治店") & (dup["枠"] == 1)].index[0]
    dup.loc[oji_idx, target_label] = original_name
    long_dup = utils.manual_shift_wide_to_long(dup, dates_df, staff_df)
    alerts_dup = utils.check_manual_shift_alerts(long_dup, staff_df, dates_df)
    if not any(d["name"] == original_name and d["date"] == target_date for d in alerts_dup["duplicates"]):
        issues.append("同一スタッフの複数店舗配置が重複アラートとして検出されない")

    return issues


# ---------------------------------------------------------------------------
# 6. 境界条件・追加ストレステスト(潜在バグの洗い出し用)
# ---------------------------------------------------------------------------
def check_different_period_no_crash() -> list[str]:
    """MANUAL_STORE_ASSIGNMENTS(10/10固定)が含まれない期間でもクラッシュせず、
    正社員21日ルール等の基本要件が保たれることを確認する。"""
    issues = []
    dates = utils.get_period_dates(2026, 6)  # 6/16-7/15、10/10を含まない
    dates_df = utils.classify_days(dates)
    staff_df = utils.default_staff_df()
    requests_df = utils.default_requests_df()
    closed_days = int((~dates_df["is_business_day"]).sum())
    result = solve_shift(dates_df, staff_df, requests_df, base_holiday_quota=closed_days + 1, time_limit_sec=90)
    if not result.is_feasible:
        return [f"6月度期間の求解に失敗: {result.status_name}"]
    issues.extend(
        [f"(6月度) {m}" for m in check_seishain_workdays(result, staff_df, dates_df, closed_days + 1)]
    )
    real_shortages, skill_issues, duplicates = check_store_requirements(result, staff_df, dates_df)
    issues.extend([f"(6月度) {m}" for m in real_shortages])
    issues.extend([f"(6月度) {m}" for m in duplicates])
    return issues


def check_special_closure_days_integration() -> list[str]:
    """特別休業日を追加した期間でも、基準公休日数の自動計算・21日ルールからの
    差分・不足許容日以外の店舗成立要件が正しく保たれることを確認する。"""
    issues = []
    dates = utils.get_period_dates(2026, 9)
    special = [dt.date(2026, 9, 17), dt.date(2026, 9, 18), dt.date(2026, 10, 1)]
    dates_df = utils.classify_days(dates, special_closure_dates=special)
    staff_df = utils.default_staff_df()
    requests_df = utils.default_requests_df()
    closed_days = int((~dates_df["is_business_day"]).sum())
    special_count = int(dates_df["is_special_closure"].sum())
    if closed_days != 8 + special_count:
        issues.append(f"総定休日数の計算不整合: closed={closed_days}, special={special_count}")
    base_quota = closed_days + 1
    expected_workdays = len(dates) - base_quota
    result = solve_shift(dates_df, staff_df, requests_df, base_holiday_quota=base_quota, time_limit_sec=90)
    if not result.is_feasible:
        return issues + [f"特別休業日ありの求解に失敗: {result.status_name}"]
    salaried = staff_df[staff_df["emp_type"].isin(["店長", "正社員"])]
    for _, row in salaried.iterrows():
        cnt = len(result.shift_df[result.shift_df["staff_id"] == row["staff_id"]])
        if cnt != expected_workdays:
            issues.append(f"{row['name']}: 出勤{cnt}日(特別休業考慮後の期待値{expected_workdays}日)")
    real_shortages, skill_issues, duplicates = check_store_requirements(result, staff_df, dates_df)
    issues.extend(real_shortages)
    issues.extend(duplicates)
    return issues


def check_heavy_conflict_no_crash() -> list[str]:
    """大量の絶対休が重複しても、ソルバーがクラッシュ(INFEASIBLE)せず、
    不足として警告付きで解を返すことを確認する(既存の設計方針の回帰確認)。"""
    dates = utils.get_period_dates(2026, 9)
    dates_df = utils.classify_days(dates)
    staff_df = utils.default_staff_df()
    biz_days = list(dates_df.loc[dates_df["is_business_day"], "date"])
    rows = []
    for name in ["生駒", "辻本", "内田", "小林", "加藤", "田中", "長瀬"]:
        sid = staff_df.loc[staff_df["name"] == name, "staff_id"].iloc[0]
        for d in biz_days[:15]:
            rows.append({"staff_id": sid, "name": name, "date": d, "kind": "絶対休"})
    requests_df = pd.DataFrame(rows)
    closed_days = int((~dates_df["is_business_day"]).sum())
    result = solve_shift(dates_df, staff_df, requests_df, base_holiday_quota=closed_days + 1, time_limit_sec=90)
    if not result.is_feasible:
        return [f"大量休み希望衝突時にソルバーがクラッシュ: {result.status_name}"]
    return []


def check_manual_shift_alerts_robustness() -> list[str]:
    """未計算月度への切り替え等で shift_df が None/空/列欠損になった場合でも
    check_manual_shift_alerts がクラッシュしないことを確認する(回帰確認)。"""
    issues = []
    dates = utils.get_period_dates(2026, 9)
    dates_df = utils.classify_days(dates)
    staff_df = utils.default_staff_df()
    cases = {
        "None": None,
        "empty_no_columns": pd.DataFrame(),
        "properly_empty": pd.DataFrame(columns=["date", "store", "staff_id", "name", "emp_type"]),
        "not_a_dataframe": "garbage",
    }
    for label, shift_df in cases.items():
        try:
            alerts = utils.check_manual_shift_alerts(shift_df, staff_df, dates_df)
            if alerts != {"duplicates": [], "shortages": [], "skill_issues": []}:
                issues.append(f"{label}: 想定外のアラートが返された({alerts})")
        except Exception as e:  # noqa: BLE001 - このチェック自体が「例外が出ないこと」を検証する
            issues.append(f"{label}: 例外が発生しクラッシュした({type(e).__name__}: {e})")
    return issues


def check_month_switch_no_crash_and_kyuka_isolation() -> list[str]:
    """実際のStreamlitアプリ(app.py)上で、(1)最適化済みの月度から未計算の別月度へ
    切り替えてもクラッシュせず結果がリセットされること、(2)月度ごとの希望休データが
    互いに独立して保存・復元されることを、AppTestで実機に近い形で確認する。
    """
    issues = []
    # このチェックは実際の app.py (本番の data/ ディレクトリを参照) を通すため、
    # 既存の実運用データを壊さないよう、対象月度の保存ファイルを一時退避し、
    # 終了後に必ず元の内容へ復元する。
    target_paths = [utils.saved_kyuka_path_for(y, m) for y, m in [(2026, 9), (2026, 11)]] + [
        utils.LATEST_SHIFT_PATH,
        utils.LATEST_SHIFT_META_PATH,
    ]
    backups: dict[str, bytes] = {}
    for p in target_paths:
        if os.path.exists(p):
            with open(p, "rb") as f:
                backups[p] = f.read()
            os.remove(p)

    try:
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file("app.py", default_timeout=120)
        at.run()
        at.sidebar.button[0].click().run()
        if list(at.exception):
            issues.append(f"初回最適化で例外: {at.exception}")

        sid = at.session_state["staff_df"].loc[at.session_state["staff_df"]["name"] == "生駒", "staff_id"].iloc[0]
        sept_df = pd.concat(
            [
                at.session_state["requests_df"],
                pd.DataFrame([{"staff_id": sid, "name": "生駒", "date": dt.date(2026, 9, 20), "kind": "希望休"}]),
            ],
            ignore_index=True,
        )
        at.session_state["requests_df"] = sept_df
        at.run()

        month_selectbox = at.sidebar.selectbox[0]
        month_selectbox.select(11).run()
        if list(at.exception):
            issues.append(f"未計算月度(11月)への切り替えで例外(=クラッシュ再現): {at.exception}")
        if at.session_state["solve_result"] is not None:
            issues.append("月度切り替え後もsolve_resultが古い月度のまま残っている")
        if len(at.session_state["requests_df"]) != 0:
            issues.append("月度切り替え後、希望休データが空になっていない(月度分離の不備)")

        tab3 = at.tabs[2]
        if not any("最適化を実行" in i.value for i in tab3.info):
            issues.append("未計算月度でタブ3に案内メッセージが表示されない")
        if list(at.exception):
            issues.append(f"タブ3描画時に例外: {at.exception}")

        month_selectbox2 = at.sidebar.selectbox[0]
        month_selectbox2.select(9).run()
        restored = at.session_state["requests_df"]
        if len(restored) != 1 or restored.iloc[0]["name"] != "生駒":
            issues.append(f"9月度に戻した際の希望休データが正しく復元されない(件数={len(restored)})")
    finally:
        # 一時退避したファイルを元に戻す。テストが新規作成したファイルのうち
        # 元々存在しなかったものは削除する。
        for p in target_paths:
            if p in backups:
                os.makedirs(os.path.dirname(p), exist_ok=True)
                with open(p, "wb") as f:
                    f.write(backups[p])
            elif os.path.exists(p):
                os.remove(p)

    return issues


# ---------------------------------------------------------------------------
# メイン実行
# ---------------------------------------------------------------------------
def main():
    print("=" * 70)
    print("シフト最適化アプリ 網羅的整合性検証")
    print("=" * 70)

    dates_df, staff_df, requests_df, result, base_holiday_quota = setup_default_scenario()
    print(f"ソルバーステータス: {result.status_name} (feasible={result.is_feasible})")
    if not result.is_feasible:
        print("!!! ソルバーが解を返しませんでした。以降の検証を中止します。")
        return 1

    record(
        "1. 正社員・店長の出勤日数(基準出勤日数の厳守)",
        check_seishain_workdays(result, staff_df, dates_df, base_holiday_quota),
    )

    real_shortages, skill_issues, duplicates = check_store_requirements(result, staff_df, dates_df)
    record("2a. 店舗成立要件(不足許容日を除く)", real_shortages)
    record("2b. 測定・加工スキル保持者の同席", skill_issues)
    record("2c. 同日重複配置の非存在(AI結果)", duplicates)
    record("2d. 名古屋中川店 土日祝 社員2名以上", check_nakagawa_weekend(result, staff_df, dates_df))

    record("3a. 個別スタッフの勤務制限(若松/若林/中村)", check_individual_restrictions(result, staff_df, dates_df))
    record("3b. 嘱託・パートの契約勤務日数レンジ", check_day_ranges(result, staff_df))
    record(
        "3c. 稲沢店における若松優先配置(竹内=稲沢+若松=他店の禁止)",
        check_wakamatsu_takeuchi_inazawa_priority(result, staff_df),
    )

    record(
        "4a. 有休代替候補の妥当性(正社員除外/上限未到達)",
        check_leave_availability_integrity(result, staff_df, dates_df, requests_df),
    )
    record("4b. 既に休みのスタッフが候補に出現しないか", check_leave_candidate_excludes_already_off(staff_df, dates_df))

    record("5a. 希望休・有休データのローカル永続化", check_requests_persistence())
    record("5b. 最適化結果のローカル永続化", check_shift_result_persistence())
    record("5c. 手動編集後のアラート連動", check_manual_edit_alert_sync(result, staff_df, dates_df))

    record("6a. 別期間(10/10を含まない月)でのクラッシュ非発生と基本要件", check_different_period_no_crash())
    record("6b. 特別休業日を含む期間の整合性", check_special_closure_days_integration())
    record("6c. 大量休み希望衝突時のクラッシュ非発生", check_heavy_conflict_no_crash())
    record("7a. check_manual_shift_alerts の異常入力耐性", check_manual_shift_alerts_robustness())
    record(
        "7b. 月度切り替え時のクラッシュ非発生・希望休データ分離",
        check_month_switch_no_crash_and_kyuka_isolation(),
    )

    print()
    print("-" * 70)
    print("検証結果一覧")
    print("-" * 70)
    n_pass = 0
    n_fail = 0
    for name, passed, detail in RESULTS:
        mark = "✅ PASS" if passed else "❌ NG"
        if passed:
            n_pass += 1
        else:
            n_fail += 1
        print(f"[{mark}] {name}")
        if not passed:
            print(f"         └ {detail}")

    print("-" * 70)
    print(f"合計: {len(RESULTS)}件 / PASS {n_pass}件 / NG {n_fail}件")
    print("=" * 70)
    return 0 if n_fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
