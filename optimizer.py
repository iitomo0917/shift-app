"""
眼鏡店7店舗 シフト最適化エンジン (Google OR-Tools CP-SAT)

方針:
  - 「社長」はスタッフマスタに一切存在しないため、変数レベルで現場出勤ゼロが保証される。
  - 店長7名は各自の店舗に完全固定（他店ヘルプ不可）。勤務可能店舗の集合自体を
    home_store 1店舗のみに制限する(=変数レベルのハード制約)ことで、ペナルティの
    掛け忘れ等に依存しない確実な固定を実現している。
  - 店舗の必要人員体制・スキル要件・嘱託の最低勤務日数・パートの月間稼働日数レンジは、
    達成できない場合でもソルバーが FAIL(INFEASIBLE) しないよう、すべて
    「不足/超過スラック変数＋重いペナルティ」で表現する（不足時アラート要件への対応）。
  - 各店舗の「1日あたり社員は最大2名まで」「許可されたパターン以外の3〜4名体制は
    禁止」は、社員数の上限キャップ・店舗ごとの合計人数キャップとして真のハード制約
    (スラックなし)で表現する。これにより体制が勝手に膨らむことは構造的に発生しない。
  - 定休日/絶対休/有休確定/最大6連勤/1日1店舗/店舗ごとの人員上限キャップのみは
    真のハード制約として扱う。
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass

import pandas as pd
from ortools.sat.python import cp_model

import utils

# ---------------------------------------------------------------------------
# ペナルティ重み
# ---------------------------------------------------------------------------
W_SHORTAGE = 1000       # 店舗の必要人員(社員/パート組合せ)が満たせない
W_SKILL = 500           # 測定・加工スキル保有者が1人もいない
W_SHOKUTAKU_MIN = 800   # 嘱託の期間内最低勤務日数を満たせない
W_PART_RANGE = 700      # パートの期間内勤務日数が許容レンジ[min,max]を外れた
W_HOLIDAY_DEV = 900     # 店長・正社員の出勤日数が規定(21日、有休確定者は20日等)から
                        # 1日でもズレることへの重いペナルティ。店舗の必要人員/スキル要件
                        # (W_SHORTAGE/W_SKILL)には次ぐが、他の全てのソフト優先度
                        # (パート活用・嘱託ボーナス・優先店舗など)を確実に上回るよう、
                        # 意図的に高く設定している(=事実上のハード制約として機能させる)。
W_SOFT_REQUEST = 15     # 希望休/有休申請を無視して出勤させた
W_SOFT_REQUEST_PRIORITY = 40  # 有給取得可能日に申請された有休申請を無視した場合の重い罰則
W_HELP = 1              # 正社員が主所属以外の店舗に出た(ヘルプ)回数
W_CONSEC5 = 3           # 嘱託が5連勤以上になった場合
W_WEEKEND_MAX = 2       # 土日出勤の最大値(公平化)
W_PREFERRED_STORE = 8   # 優先店舗(例: 吉田の天白植田店)以外への配属(軽微なペナルティ)
W_PAIR_REQUIREMENT = 1200  # 期間内必須ペア勤務(例: 辻本+尾澤@大治店)が0回だった場合
W_SHOKUTAKU_BONUS = 6   # 嘱託が1日出勤するごとの報酬(負の目的関数寄与)。
                        # 最低日数で頭打ちにせず、上限日数まで積極的に組み込み、
                        # 正社員の有休取得可能枠(店舗の余剰人員)を創出するための重み。
W_PART_BONUS = 5        # パート(尾澤/野道/柴田)が1日出勤するごとの報酬(負の目的関数寄与)。
                        # 上限日数付近まで積極的に活用し、正社員の有休枠創出につなげる。
W_STORE_PATTERN_PREF = 5  # 店舗の平日体制パターンに関するソフトな優先度。
                          # 正の値=社員2名体制を優先(パート併用を軽く抑制)、
                          # 負の値=社員1名+パート体制を優先(パート併用を軽く推奨)。

BIG_M = 10


@dataclass
class SolveResult:
    status_name: str
    is_feasible: bool
    shift_df: pd.DataFrame
    shortages: list[dict]
    staff_summary_df: pd.DataFrame
    objective_value: float
    solver_wall_time: float


def _day_work_expr(work_vars, sid, d, allowed_stores):
    terms = [work_vars[(sid, d, st)] for st in allowed_stores if (sid, d, st) in work_vars]
    return cp_model.LinearExpr.Sum(terms) if terms else 0


def solve_shift(
    dates_df: pd.DataFrame,
    staff_df: pd.DataFrame,
    requests_df: pd.DataFrame,
    base_holiday_quota: int,
    time_limit_sec: int = 30,
) -> SolveResult:
    model = cp_model.CpModel()

    business_days = list(dates_df.loc[dates_df["is_business_day"], "date"])
    all_days = list(dates_df["date"])
    closed_days_count = int((~dates_df["is_business_day"]).sum())
    # 土曜・日曜・祝日のうち営業日であるものの集合(店舗別の土日祝限定ルールに使用)
    weekend_holiday_days: set = set(
        dates_df.loc[dates_df["is_business_day"] & dates_df["is_weekend_or_holiday"], "date"]
    )
    # 祝日を含まない、純粋な土曜・日曜の営業日集合(若松の大治店優先配置ルール用)
    saturday_sunday_days: set = {d for d in business_days if d.weekday() in (5, 6)}

    # 店長7名（各自店舗に完全固定・他店ヘルプ不可）を除く、店長以外の正社員のみ
    seishain = staff_df[staff_df["emp_type"] == "正社員"]
    # 店長+正社員：基準公休日数の消化対象となる月給スタッフ
    salaried_staff = staff_df[staff_df["emp_type"].isin(["店長", "正社員"])]
    shokutaku = staff_df[staff_df["emp_type"] == "嘱託"]
    # 店舗の必要体制カウントに算入する「社員」区分（店長を含む）
    employee_staff = staff_df[staff_df["emp_type"].isin(utils.EMPLOYEE_TYPES)]
    part_staff = staff_df[staff_df["emp_type"] == "パート"]

    # 実際に使用する勤務可能店舗は、DataFrame上の値を鵜呑みにせず
    # ここで再導出する（店長の自店舗固定＝ハード制約を二重に保証するため）。
    effective_allowed: dict[str, list[str]] = {
        row.staff_id: utils.derive_allowed_stores(row) for row in staff_df.itertuples()
    }

    staff_allowed = effective_allowed
    staff_name = {row.staff_id: row.name for row in staff_df.itertuples()}
    staff_type = {row.staff_id: row.emp_type for row in staff_df.itertuples()}
    staff_skill = {row.staff_id: bool(row.has_skill) for row in staff_df.itertuples()}
    staff_home = {row.staff_id: getattr(row, "home_store", None) for row in staff_df.itertuples()}
    staff_role = {row.staff_id: getattr(row, "part_role", None) for row in staff_df.itertuples()}
    name_to_id = {row.name: row.staff_id for row in staff_df.itertuples()}

    # --- 1. 出勤変数 -------------------------------------------------------
    work_vars: dict[tuple[str, dt.date, str], cp_model.IntVar] = {}
    for row in staff_df.itertuples():
        sid = row.staff_id
        for d in business_days:
            for st in effective_allowed[sid]:
                work_vars[(sid, d, st)] = model.NewBoolVar(f"w_{sid}_{d.isoformat()}_{st}")

    # --- 2. 1日1店舗まで -----------------------------------------------------
    for row in staff_df.itertuples():
        sid = row.staff_id
        for d in business_days:
            terms = [work_vars[(sid, d, st)] for st in effective_allowed[sid]]
            model.Add(sum(terms) <= 1)

    # --- 3. 定休日は全員休み(変数自体を作らないため自動的に満たされる) -------------
    #     (closed daysの変数は存在しない)

    # --- 4. 絶対休・有休確定はハード制約 -----------------------------------
    hard_off_dates: dict[str, set] = {}
    soft_off_dates: dict[str, dict] = {}
    if not requests_df.empty:
        for _, r in requests_df.iterrows():
            sid, d, kind = r["staff_id"], r["date"], r["kind"]
            if sid not in staff_allowed:
                continue
            if kind in utils.HARD_OFF_KINDS:
                hard_off_dates.setdefault(sid, set()).add(d)
            elif kind in utils.SOFT_OFF_KINDS:
                soft_off_dates.setdefault(sid, {})[d] = kind

    for sid, dates in hard_off_dates.items():
        for d in dates:
            for st in staff_allowed.get(sid, []):
                if (sid, d, st) in work_vars:
                    model.Add(work_vars[(sid, d, st)] == 0)

    # --- 4b. 土日祝の店舗別配置禁止(ハード制約) ------------------------------
    #     例: 若松の新蟹江店(検査技能の観点)、野道の大治店(平日のみ稼働)。
    for row in staff_df.itertuples():
        forbidden = getattr(row, "weekend_holiday_forbidden_stores", None) or []
        if not forbidden:
            continue
        sid = row.staff_id
        for d in weekend_holiday_days:
            for st in forbidden:
                if (sid, d, st) in work_vars:
                    model.Add(work_vars[(sid, d, st)] == 0)

    # --- 4c. 純粋な土曜・日曜(祝日除く)の店舗別配置禁止(ハード制約) --------------
    #     例: 若松は土日に稲沢店へは配置しない(=土日に出勤する場合は大治店優先)。
    for row in staff_df.itertuples():
        forbidden_ss = getattr(row, "saturday_sunday_forbidden_stores", None) or []
        if not forbidden_ss:
            continue
        sid = row.staff_id
        for d in saturday_sunday_days:
            for st in forbidden_ss:
                if (sid, d, st) in work_vars:
                    model.Add(work_vars[(sid, d, st)] == 0)

    # --- 4d. 手動指定シフト(特定日・特定店舗のロースター固定、ハード制約) -----------
    #     例: 10/10(土) 徳重店=中村+山岡を確定配置、稲沢店=生駒のみ(2人目は
    #     不足として許容)。指定した氏名は必ずその店舗に出勤させ、それ以外の
    #     スタッフは同日同店舗への配置を禁止する(=ロースターを完全固定)。
    #     必要人数に満たない指定は、既存の店舗別不足スラック機構がそのまま
    #     ペナルティ付きで吸収するため、ソルバーが落ちることはない。
    manual_fixed_ids_by_store_day: dict[tuple[dt.date, str], set[str]] = {}
    for spec in utils.MANUAL_STORE_ASSIGNMENTS:
        d = spec["date"]
        store = spec["store"]
        if d not in business_days:
            continue
        fixed_ids = {name_to_id[n] for n in spec["names"] if n in name_to_id}
        manual_fixed_ids_by_store_day[(d, store)] = fixed_ids
        for sid in fixed_ids:
            if (sid, d, store) in work_vars:
                model.Add(work_vars[(sid, d, store)] == 1)
        if spec.get("exclude_others", True):
            # 指定人数で意図的に不足させたい場合など、それ以外のスタッフの
            # 同日同店舗への配置も禁止する(=ロースターを完全固定)。
            for sid in staff_df["staff_id"]:
                if sid in fixed_ids:
                    continue
                if (sid, d, store) in work_vars:
                    model.Add(work_vars[(sid, d, store)] == 0)

    # --- 4e. 稲沢店における若松・竹内の優先度制御(ハード制約) --------------------
    #     同一営業日に若松・竹内がともに出勤する場合、「竹内を稲沢店に配置し、
    #     かつ若松を稲沢店以外に配置する」という組み合わせを禁止する。
    #     これにより、若松が出勤する日に竹内が稲沢店へ入るなら若松も稲沢店に
    #     入らざるを得なくなる(=事実上、若松が稲沢店へ優先配置され、竹内は
    #     他店舗へ回る)。どちらか一方が休みの日は該当項が自動的に0になるため、
    #     制約は実質的に無効化され、既存の若松の土日祝配置禁止(4b)・純粋な
    #     土日の稲沢店配置禁止(4c)・竹内のスキル/移動可能店舗制約とも矛盾なく
    #     共存する(いずれも変数への追加の禁止であり、==1のような強制ではない
    #     ため、既存のFEASIBLE解の探索可能性を損なわない)。
    sid_wakamatsu = name_to_id.get("若松")
    sid_takeuchi = name_to_id.get("竹内")
    if sid_wakamatsu and sid_takeuchi:
        for d in business_days:
            takeuchi_inazawa = work_vars.get((sid_takeuchi, d, "稲沢店"))
            if takeuchi_inazawa is None:
                continue
            wakamatsu_other_terms = [
                work_vars[(sid_wakamatsu, d, st)]
                for st in effective_allowed[sid_wakamatsu]
                if st != "稲沢店" and (sid_wakamatsu, d, st) in work_vars
            ]
            if not wakamatsu_other_terms:
                continue
            model.Add(takeuchi_inazawa + cp_model.LinearExpr.Sum(wakamatsu_other_terms) <= 1)

    # --- 5. 最大6連勤(7日ウィンドウの合計<=6) --------------------------------
    for row in staff_df.itertuples():
        sid = row.staff_id
        for i in range(len(all_days) - 6):
            window = all_days[i : i + 7]
            expr_terms = []
            for d in window:
                if d in business_days:
                    expr_terms.append(_day_work_expr(work_vars, sid, d, effective_allowed[sid]))
            if expr_terms:
                model.Add(cp_model.LinearExpr.Sum(expr_terms) <= 6)

    penalty_terms: list = []
    shortage_records: dict[tuple, dict] = {}

    # --- 6. 店舗別 必要人員体制(不足はスラックで許容) ------------------------
    for d in business_days:
        for store in utils.STORES:
            employees_here = [
                sid for sid in employee_staff["staff_id"] if (sid, d, store) in work_vars
            ]
            e_expr = cp_model.LinearExpr.Sum([work_vars[(sid, d, store)] for sid in employees_here]) if employees_here else 0

            if store in utils.GENERAL_STORES:
                # 稲沢店・新蟹江店・極楽店: 社員2名固定。パート不可・3名以上は完全禁止。
                model.Add(e_expr <= 2)  # 上限キャップ(ハード) — 3名以上への膨張を構造的に禁止
                shortfall_e2 = model.NewIntVar(0, 2, f"sf_{store}_{d}_e2")
                model.Add(e_expr + shortfall_e2 >= 2)
                penalty_terms.append((shortfall_e2, W_SHORTAGE))
                shortage_records[(d, store)] = {"type": "general", "vars": {"不足人数": shortfall_e2}}

            elif store in utils.COMBO_STORE_PART_ROLES:
                roles = utils.COMBO_STORE_PART_ROLES[store]
                part_here = [
                    sid
                    for sid in part_staff["staff_id"]
                    if staff_role.get(sid) in roles and (sid, d, store) in work_vars
                ]
                p_expr = cp_model.LinearExpr.Sum([work_vars[(sid, d, store)] for sid in part_here]) if part_here else 0

                # 社員は最大2名まで(上限キャップ・ハード) — これにより「社員2名+パート1名」
                # を超える体制(3名超)は構造的に発生しない。
                model.Add(e_expr <= 2)

                if store == "大治店":
                    # 大治店は必ず2名体制(3名以上禁止)。パートは1名まで(尾澤 or 野道のどちらか)。
                    model.Add(p_expr <= 1)
                    model.Add(e_expr + p_expr <= 2)

                combo = model.NewBoolVar(f"combo_{store}_{d}")

                if store == "天白植田店":
                    if d in weekend_holiday_days:
                        # 土日祝は「社員1名+パート柴田」の2名体制を禁止し、
                        # 必ず社員2名以上(2名 or 2名+柴田の3名)とする(ハード制約)。
                        model.Add(combo == 0)
                    else:
                        # 平日は3名体制を完全禁止し、必ず2名体制のみとする(ハード制約)。
                        model.Add(e_expr + p_expr <= 2)
                elif store == "名古屋中川店":
                    if d in weekend_holiday_days:
                        # 土日祝は「社員1名+パート尾澤」の2名体制を禁止し、必ず社員2名
                        # 以上(2名 or 2名+尾澤の3名)とする(ハード制約)。平日は従来通り
                        # 社員1名+パート尾澤の2名体制も許可する。
                        model.Add(combo == 0)

                sf_e2 = model.NewIntVar(0, 2, f"sf_{store}_{d}_e2")
                sf_e1 = model.NewIntVar(0, 1, f"sf_{store}_{d}_e1")
                sf_p1 = model.NewIntVar(0, 1, f"sf_{store}_{d}_p1")
                model.Add(e_expr + sf_e2 >= 2).OnlyEnforceIf(combo.Not())
                model.Add(e_expr + sf_e1 >= 1).OnlyEnforceIf(combo)
                model.Add(p_expr + sf_p1 >= 1).OnlyEnforceIf(combo)
                penalty_terms += [(sf_e2, W_SHORTAGE), (sf_e1, W_SHORTAGE), (sf_p1, W_SHORTAGE)]
                shortage_records[(d, store)] = {
                    "type": "combo",
                    "vars": {"社員2名不足(社員のみ体制)": sf_e2, "社員1名不足(組合せ体制)": sf_e1, "パート不足(組合せ体制)": sf_p1},
                    "combo": combo,
                }

                # --- 店舗体制の優先度(ソフト): 売上規模・客数負荷に応じたバランス調整 ---
                if store == "大治店":
                    # パート活用推奨店舗: 平日・土日ともに「社員1名+パート」を推奨(報酬)。
                    penalty_terms.append((combo, -W_STORE_PATTERN_PREF))
                elif store == "天白植田店" and d not in weekend_holiday_days:
                    # パート活用推奨店舗: 平日は「社員1名+パート柴田」を推奨(報酬)。
                    penalty_terms.append((combo, -W_STORE_PATTERN_PREF))
                elif store == "名古屋中川店" and d not in weekend_holiday_days:
                    # 高負荷店舗: 平日でも極力「社員2名体制」を優先(パート併用は軽く抑制)。
                    penalty_terms.append((combo, W_STORE_PATTERN_PREF))

            elif store == utils.TOKUSHIGE_STORE:
                b_here = [
                    sid for sid in part_staff["staff_id"]
                    if staff_role.get(sid) == "B" and (sid, d, store) in work_vars
                ]
                c_here = [
                    sid for sid in part_staff["staff_id"]
                    if staff_role.get(sid) == "C" and (sid, d, store) in work_vars
                ]
                b_expr = cp_model.LinearExpr.Sum([work_vars[(sid, d, store)] for sid in b_here]) if b_here else 0
                c_expr = cp_model.LinearExpr.Sum([work_vars[(sid, d, store)] for sid in c_here]) if c_here else 0

                # 社員は最大2名まで、かつ「社員+パートB+パートC」の合計は最大3名まで
                # (上限キャップ・ハード)。社員2名+パート2名(=4名)は構造的に禁止。
                model.Add(e_expr <= 2)
                model.Add(e_expr + b_expr + c_expr <= 3)

                combo = model.NewBoolVar(f"combo_{store}_{d}")
                if d in weekend_holiday_days:
                    # 土日祝は「社員1名+パートB+パートC」の3名体制を禁止し、
                    # 必ず社員2名以上とする(ハード制約)。社員2名+パート1名の
                    # 3名体制(パターン②の亜種)は禁止対象外のため許可される。
                    model.Add(combo == 0)
                sf_e2 = model.NewIntVar(0, 2, f"sf_{store}_{d}_e2")
                sf_e1 = model.NewIntVar(0, 1, f"sf_{store}_{d}_e1")
                sf_b = model.NewIntVar(0, 1, f"sf_{store}_{d}_b")
                sf_c = model.NewIntVar(0, 1, f"sf_{store}_{d}_c")
                model.Add(e_expr + sf_e2 >= 2).OnlyEnforceIf(combo.Not())
                model.Add(e_expr + sf_e1 >= 1).OnlyEnforceIf(combo)
                model.Add(b_expr + sf_b >= 1).OnlyEnforceIf(combo)
                model.Add(c_expr + sf_c >= 1).OnlyEnforceIf(combo)
                penalty_terms += [(sf_e2, W_SHORTAGE), (sf_e1, W_SHORTAGE), (sf_b, W_SHORTAGE), (sf_c, W_SHORTAGE)]
                if d not in weekend_holiday_days:
                    # 高負荷店舗: 平日でも極力「社員2名体制」を優先(パート併用は軽く抑制)。
                    penalty_terms.append((combo, W_STORE_PATTERN_PREF))
                shortage_records[(d, store)] = {
                    "type": "tokushige",
                    "vars": {
                        "社員2名不足(社員のみ体制)": sf_e2,
                        "社員1名不足(3名体制)": sf_e1,
                        "パートB不足(3名体制)": sf_b,
                        "パートC不足(3名体制)": sf_c,
                    },
                    "combo": combo,
                }

            # --- スキル要件: 店舗に出勤中の誰か1名は測定・加工スキル保有者 ---
            all_here = [sid for sid in staff_df["staff_id"] if (sid, d, store) in work_vars]
            staffed_expr = cp_model.LinearExpr.Sum([work_vars[(sid, d, store)] for sid in all_here]) if all_here else 0
            skilled_here = [sid for sid in all_here if staff_skill.get(sid)]
            skill_expr = cp_model.LinearExpr.Sum([work_vars[(sid, d, store)] for sid in skilled_here]) if skilled_here else 0

            staffed_bool = model.NewBoolVar(f"staffed_{store}_{d}")
            model.Add(staffed_expr >= staffed_bool)
            model.Add(staffed_expr <= BIG_M * staffed_bool)
            skill_shortfall = model.NewIntVar(0, 1, f"sf_{store}_{d}_skill")
            model.Add(skill_expr + skill_shortfall >= staffed_bool)
            penalty_terms.append((skill_shortfall, W_SKILL))
            shortage_records[(d, store)].setdefault("vars", {})["スキル保有者不在"] = skill_shortfall

    # --- 7. 希望休・有休申請(ソフト) -----------------------------------------
    availability_df = utils.compute_paid_leave_availability(dates_df, staff_df, requests_df)
    priority_dates = set(availability_df.loc[availability_df["paid_leave_slots"] > 0, "date"])

    for sid, date_kind_map in soft_off_dates.items():
        allowed = staff_allowed.get(sid, [])
        for d, kind in date_kind_map.items():
            expr = _day_work_expr(work_vars, sid, d, allowed)
            if isinstance(expr, int) and expr == 0:
                continue
            if kind == "有休申請" and d in priority_dates:
                weight = W_SOFT_REQUEST_PRIORITY
            else:
                weight = W_SOFT_REQUEST
            penalty_terms.append((expr, weight))

    # --- 8. 嘱託: 期間内勤務日数レンジ[min,max]の遵守(不足/超過はスラックで許容) --
    shokutaku_range_info: dict[str, tuple[cp_model.IntVar, cp_model.IntVar, int, int]] = {}
    for row in shokutaku.itertuples():
        sid = row.staff_id
        min_days = int(row.min_workdays or 0)
        max_days = int(row.max_workdays) if pd.notna(getattr(row, "max_workdays", None)) else len(business_days)
        # 特別休業日の追加等で営業日数が想定より少ない期間でも、レンジが物理的に
        # 達成不可能な数値のまま残らないよう、営業日数を超えないクリップを行う
        # (恒常的な巨大な不足数の表示や無用なペナルティの蓄積を避けるため)。
        min_days = min(min_days, len(business_days))
        max_days = min(max_days, len(business_days))
        total_expr = cp_model.LinearExpr.Sum(
            [_day_work_expr(work_vars, sid, d, effective_allowed[sid]) for d in business_days]
        )
        under = model.NewIntVar(0, min_days, f"sf_min_{sid}")
        over = model.NewIntVar(0, len(business_days), f"sf_max_{sid}")
        model.Add(total_expr + under >= min_days)
        model.Add(total_expr - over <= max_days)
        penalty_terms.append((under, W_SHOKUTAKU_MIN))
        penalty_terms.append((over, W_SHOKUTAKU_MIN))
        # 嘱託は最低日数で頭打ちにせず、[min,max]レンジの上限に向けて積極的に
        # 出勤させる(負の重み=報酬)。over側の重いペナルティ(W_SHOKUTAKU_MIN)が
        # 上限を超えないよう常に上回るため、maxを超えて増やされることはない。
        penalty_terms.append((total_expr, -W_SHOKUTAKU_BONUS))
        shokutaku_range_info[sid] = (under, over, min_days, max_days)

    # --- 8b. パート: 期間内勤務日数レンジ[min,max]の遵守(不足/超過はスラックで許容) --
    #     リソースを使い切らせすぎず、かつ最低限は稼働させるための制約。
    #     絶対休等との衝突でレンジを満たせない場合でもソルバーを落とさないよう、
    #     重いペナルティ付きスラックで表現する。
    part_range_info: dict[str, tuple[cp_model.IntVar, cp_model.IntVar, int, int]] = {}
    for row in part_staff.itertuples():
        sid = row.staff_id
        role = staff_role.get(sid)
        min_days, max_days = utils.PART_ROLE_WORKDAY_RANGE.get(role, (0, len(business_days)))
        # 特別休業日の追加等で営業日数が減っても物理的に不可能なレンジのままに
        # ならないよう、営業日数でクリップする。
        min_days = min(min_days, len(business_days))
        max_days = min(max_days, len(business_days))
        total_expr = cp_model.LinearExpr.Sum(
            [_day_work_expr(work_vars, sid, d, effective_allowed[sid]) for d in business_days]
        )
        under = model.NewIntVar(0, min_days, f"sf_part_under_{sid}")
        over = model.NewIntVar(0, len(business_days), f"sf_part_over_{sid}")
        model.Add(total_expr + under >= min_days)
        model.Add(total_expr - over <= max_days)
        penalty_terms.append((under, W_PART_RANGE))
        penalty_terms.append((over, W_PART_RANGE))
        if role in ("A", "D", "E"):
            # 尾澤(A)・柴田(D)・野道(E): 希望休以外の空き日は上限日数付近まで積極的に
            # 出勤させる(負の重み=報酬)。over側の重いペナルティが上限超過を防ぐ。
            penalty_terms.append((total_expr, -W_PART_BONUS))
        part_range_info[sid] = (under, over, min_days, max_days)

    # --- 9. 嘱託: 5連勤回避(ソフト) ------------------------------------------
    for row in shokutaku.itertuples():
        sid = row.staff_id
        for i in range(len(all_days) - 4):
            window = [d for d in all_days[i : i + 5] if d in business_days]
            if not window:
                continue
            expr = cp_model.LinearExpr.Sum([_day_work_expr(work_vars, sid, d, effective_allowed[sid]) for d in window])
            excess = model.NewIntVar(0, 5, f"excess5_{sid}_{i}")
            model.Add(expr - 4 <= excess)
            penalty_terms.append((excess, W_CONSEC5))

    # --- 10. 正社員(店長以外): 主所属以外(ヘルプ)出勤の最小化(ソフト) ------------
    #     店長は allowed_stores が home_store のみなので help_terms は常に空
    #     (=このループの対象外と等価)。念のため店長も含めて回すが影響はない。
    for row in salaried_staff.itertuples():
        sid = row.staff_id
        home = row.home_store
        help_terms = [
            work_vars[(sid, d, st)]
            for d in business_days
            for st in effective_allowed[sid]
            if st != home and (sid, d, st) in work_vars
        ]
        if help_terms:
            penalty_terms.append((cp_model.LinearExpr.Sum(help_terms), W_HELP))

    # --- 10b. 優先店舗への配属(ソフト) ----------------------------------------
    #     例: 吉田(嘱託)は出勤する日、可能な限り天白植田店へ配属する。
    #     他店舗の欠員等でやむを得ず他店舗へ出た場合も出勤自体は許可し、
    #     軽微なペナルティのみを課す。
    for row in staff_df.itertuples():
        preferred_store = getattr(row, "preferred_store", None)
        if not preferred_store:
            continue
        sid = row.staff_id
        other_store_terms = [
            work_vars[(sid, d, st)]
            for d in business_days
            for st in effective_allowed[sid]
            if st != preferred_store and (sid, d, st) in work_vars
        ]
        if other_store_terms:
            penalty_terms.append((cp_model.LinearExpr.Sum(other_store_terms), W_PREFERRED_STORE))

    # --- 11. 店長+正社員: 出勤日数を規定日数(21日/有休確定者は20日等)に厳格化 -----
    #     W_HOLIDAY_DEV を非常に重く設定しているため、事実上のハード制約として
    #     機能する(店舗人員・スキル要件のみを上回って優先される)。それでも数式上は
    #     スラック変数を介した制約のため、極端な休み希望の衝突があってもソルバーが
    #     INFEASIBLE で落ちることはなく、必ず解が得られる。
    target_off = max(0, base_holiday_quota - closed_days_count)
    n_business_days = len(business_days)
    holiday_dev_info = {}
    for row in salaried_staff.itertuples():
        sid = row.staff_id
        total_work_expr = cp_model.LinearExpr.Sum(
            [_day_work_expr(work_vars, sid, d, effective_allowed[sid]) for d in business_days]
        )

        if n_business_days >= 1:
            # 店長・正社員11名は、定休日を除いた基準公休の「残り1日」を必ず
            # 個別公休(出勤なし)として取得できるようにする(ハード制約)。
            # その日の店舗体制は他店舗からのヘルプ(正社員/嘱託)で維持される
            # (店舗別必要体制のハード制約/不足ペナルティが別途保証する)。
            model.Add(total_work_expr <= n_business_days - 1)

        preferred_off = getattr(row, "preferred_off_date", None)
        if pd.notna(preferred_off) and preferred_off in business_days:
            # 個別公休の希望日が指定されている場合は、その日を確定公休として
            # ハード制約化する(希望日が無指定の場合はAIが上のキャップの範囲内で
            # 全社体制のバランスを見て自動的に1日を選ぶ)。
            for st_ in effective_allowed[sid]:
                if (sid, preferred_off, st_) in work_vars:
                    model.Add(work_vars[(sid, preferred_off, st_)] == 0)

        confirmed_paid_leave = 0
        if not requests_df.empty:
            confirmed_paid_leave = int(
                (
                    (requests_df["staff_id"] == sid)
                    & (requests_df["kind"] == "有休確定")
                    & (requests_df["date"].isin(business_days))
                ).sum()
            )
        # off_business_days = n_business_days - total_work
        # koukyuu_equivalent = off_business_days - confirmed_paid_leave
        dev_pos = model.NewIntVar(0, n_business_days, f"dev_pos_{sid}")
        dev_neg = model.NewIntVar(0, n_business_days, f"dev_neg_{sid}")
        model.Add(
            (n_business_days - total_work_expr) - confirmed_paid_leave - target_off == dev_pos - dev_neg
        )
        penalty_terms.append((dev_pos, W_HOLIDAY_DEV))
        penalty_terms.append((dev_neg, W_HOLIDAY_DEV))
        holiday_dev_info[sid] = (dev_pos, dev_neg, confirmed_paid_leave)

    # --- 12. 土日出勤の公平化(ソフト・ミニマックス) ----------------------------
    weekend_days = [d for d in business_days if d.weekday() in (5, 6)]
    if weekend_days:
        max_weekend = model.NewIntVar(0, len(weekend_days), "max_weekend")
        for row in employee_staff.itertuples():
            sid = row.staff_id
            wk_expr = cp_model.LinearExpr.Sum(
                [_day_work_expr(work_vars, sid, d, effective_allowed[sid]) for d in weekend_days]
            )
            model.Add(max_weekend >= wk_expr)
        penalty_terms.append((max_weekend, W_WEEKEND_MAX))

    # --- 13. 期間内必須ペア勤務(不足はスラックで許容) --------------------------
    #     例: 辻本+尾澤@大治店を期間内に最低1回は同日勤務させる。
    #     氏名で名寄せするため、氏名が変更されて解決できない組み合わせはスキップする
    #     (安全側に倒し、ソルバーを落とさない)。name_to_id は冒頭で構築済みのものを再利用。
    for spec in utils.REQUIRED_PAIR_WORKDAYS:
        name_a, name_b = spec["names"]
        pair_store = spec["store"]
        min_occurrences = spec["min_occurrences"]
        sid_a = name_to_id.get(name_a)
        sid_b = name_to_id.get(name_b)
        if not sid_a or not sid_b:
            continue

        pair_vars = []
        for d in business_days:
            if (sid_a, d, pair_store) in work_vars and (sid_b, d, pair_store) in work_vars:
                va = work_vars[(sid_a, d, pair_store)]
                vb = work_vars[(sid_b, d, pair_store)]
                pv = model.NewBoolVar(f"pair_{sid_a}_{sid_b}_{pair_store}_{d.isoformat()}")
                model.Add(pv <= va)
                model.Add(pv <= vb)
                model.Add(pv >= va + vb - 1)
                pair_vars.append(pv)

        pair_total = cp_model.LinearExpr.Sum(pair_vars) if pair_vars else 0
        pair_shortfall = model.NewIntVar(0, min_occurrences, f"sf_pair_{sid_a}_{sid_b}_{pair_store}")
        model.Add(pair_total + pair_shortfall >= min_occurrences)
        penalty_terms.append((pair_shortfall, W_PAIR_REQUIREMENT))

    # --- 目的関数 -------------------------------------------------------------
    model.Minimize(cp_model.LinearExpr.Sum([var * w for var, w in penalty_terms]))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_sec
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    status_name = solver.StatusName(status)
    is_feasible = status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    if not is_feasible:
        return SolveResult(status_name, False, pd.DataFrame(), [], pd.DataFrame(), 0.0, solver.WallTime())

    # --- 結果抽出 ---------------------------------------------------------
    shift_rows = []
    for (sid, d, st), var in work_vars.items():
        if solver.Value(var) == 1:
            shift_rows.append(
                {
                    "date": d,
                    "store": st,
                    "staff_id": sid,
                    "name": staff_name[sid],
                    "emp_type": staff_type[sid],
                }
            )
    shift_df = pd.DataFrame(shift_rows).sort_values(["date", "store", "emp_type"]) if shift_rows else pd.DataFrame(
        columns=["date", "store", "staff_id", "name", "emp_type"]
    )

    # --- 不足アラート -------------------------------------------------------
    manual_spec_by_store_day = {(s["date"], s["store"]): s for s in utils.MANUAL_STORE_ASSIGNMENTS}
    shortages = []
    for (d, store), rec in shortage_records.items():
        for label, var in rec["vars"].items():
            amount = solver.Value(var)
            if amount > 0:
                manual_spec = manual_spec_by_store_day.get((d, store))
                if manual_spec is not None and label == "不足人数":
                    # 手動指定による意図的な人員不足(例: 稲沢店を生駒1名体制のまま
                    # 2人目を無理に埋めない)は、その旨が分かる説明文を優先表示する。
                    fixed_names = "、".join(manual_spec["names"])
                    needed_total = len(manual_spec["names"]) + amount
                    shortages.append(
                        {
                            "date": d,
                            "store": store,
                            "内容": label,
                            "不足数": amount,
                            "原因候補": f"{fixed_names}{len(manual_spec['names'])}名体制 / {needed_total}人目未配置",
                        }
                    )
                    continue
                cause_names = []
                for sid in staff_df["staff_id"]:
                    off_kind = None
                    if sid in hard_off_dates and d in hard_off_dates[sid]:
                        off_kind = "絶対休/有休確定"
                    elif sid in soft_off_dates and d in soft_off_dates[sid]:
                        off_kind = soft_off_dates[sid][d]
                    if off_kind is None:
                        continue
                    relevant = (
                        staff_home.get(sid) == store
                        or store in staff_allowed.get(sid, [])
                        and staff_type.get(sid) in utils.EMPLOYEE_TYPES
                    )
                    if relevant:
                        cause_names.append(f"{staff_name[sid]}({off_kind})")
                shortages.append(
                    {
                        "date": d,
                        "store": store,
                        "内容": label,
                        "不足数": amount,
                        "原因候補": "、".join(cause_names) if cause_names else "(希望休以外の要因/純粋な人員不足)",
                    }
                )
    shortages.sort(key=lambda r: (r["date"], r["store"]))

    # --- スタッフサマリー -----------------------------------------------------
    summary_rows = []
    for row in staff_df.itertuples():
        sid = row.staff_id
        work_days = shift_df[shift_df["staff_id"] == sid] if not shift_df.empty else pd.DataFrame()
        weekend_cnt = int(work_days["date"].apply(lambda d: d.weekday() in (5, 6)).sum()) if not work_days.empty else 0
        help_cnt = 0
        if row.emp_type in ("店長", "正社員", "嘱託") and not work_days.empty:
            help_cnt = int((work_days["store"] != row.home_store).sum())

        range_note = "-"
        range_source = shokutaku_range_info.get(sid) or part_range_info.get(sid)
        if range_source:
            under, over, min_days, max_days = range_source
            u, o = solver.Value(under), solver.Value(over)
            if u > 0:
                range_note = f"下限-{u}(規定{min_days}〜{max_days}日)"
            elif o > 0:
                range_note = f"上限+{o}(規定{min_days}〜{max_days}日)"
            else:
                range_note = f"OK({min_days}〜{max_days}日)"

        summary_rows.append(
            {
                "staff_id": sid,
                "name": row.name,
                "区分": row.emp_type,
                "出勤日数": len(work_days),
                "土日出勤日数": weekend_cnt,
                "ヘルプ出勤日数": help_cnt,
                "勤務日数レンジ(嘱託/パート)": range_note,
            }
        )
    staff_summary_df = pd.DataFrame(summary_rows)

    return SolveResult(
        status_name=status_name,
        is_feasible=True,
        shift_df=shift_df,
        shortages=shortages,
        staff_summary_df=staff_summary_df,
        objective_value=solver.ObjectiveValue(),
        solver_wall_time=solver.WallTime(),
    )
