"""
逵ｼ髀｡蠎・蠎苓・ 繧ｷ繝輔ヨ譛驕ｩ蛹悶い繝励Μ - 蜈ｱ騾壹Θ繝ｼ繝・ぅ繝ｪ繝・ぅ

縺薙・繝｢繧ｸ繝･繝ｼ繝ｫ縺梧球蠖薙☆繧狗ｯ・峇:
  - 繧ｷ繝輔ヨ蟇ｾ雎｡譛滄俣・域ｯ取怦16譌･縲懃ｿ梧怦15譌･・峨・繧ｫ繝ｬ繝ｳ繝繝ｼ逕滓・縺ｨ蝟ｶ讌ｭ蛹ｺ蛻・愛螳・  - 蠎苓・繝槭せ繧ｿ繝ｻ繧ｹ繧ｿ繝・ヵ繝槭せ繧ｿ縺ｮ繝・ヵ繧ｩ繝ｫ繝医ョ繝ｼ繧ｿ
  - 譛臥ｵｦ莨第嚊縺ｮ縲悟叙蠕怜庄閭ｽ譌･繝ｻ莠ｺ謨ｰ譫縲阪・閾ｪ蜍慕ｮ怜・縺ｨ譯亥・譁・函謌・  - Excel・・xlsx・峨お繧ｯ繧ｹ繝昴・繝茨ｼ亥ｺ苓・蛻･譌･蛻･繧ｷ繝輔ヨ陦ｨ / 繧ｹ繧ｿ繝・ヵ蛻･蜃ｺ蜍､荳隕ｧ陦ｨ・・
縲檎､ｾ髟ｷ縲阪・迴ｾ蝣ｴ蜃ｺ蜍､繧ｼ繝ｭ縺檎ｵｶ蟇ｾ隕∽ｻｶ縺ｮ縺溘ａ縲√せ繧ｿ繝・ヵ繝槭せ繧ｿ繝ｻ繧ｷ繝輔ヨ螟画焚縺ｮ縺・★繧後↓繧・荳蛻・匳蝣ｴ縺励↑縺・ｼ・縺昴ｂ縺昴ｂ繝・・繧ｿ縺ｨ縺励※謖√◆縺ｪ縺・ｼ峨・"""

from __future__ import annotations

import calendar
import colorsys
import csv
import datetime as dt
import io
import json
import os
import re
import time
import unicodedata
from dataclasses import dataclass, field
from typing import Optional

import jpholiday
import pandas as pd

# ---------------------------------------------------------------------------
# 蠎苓・繝槭せ繧ｿ
# ---------------------------------------------------------------------------

STORES = [
    "遞ｲ豐｢蠎・,
    "螟ｧ豐ｻ蠎・,
    "蜷榊商螻倶ｸｭ蟾晏ｺ・,
    "譁ｰ陝ｹ豎溷ｺ・,
    "蠕ｳ驥榊ｺ・,
    "讌ｵ讌ｽ蠎・,
    "螟ｩ逋ｽ讀咲伐蠎・,
]

# 繝代・繝医・蠖ｹ蜑ｲ -> 蜍､蜍吝庄閭ｽ蠎苓・・育ｵｶ蟇ｾ蛻ｶ髯撰ｼ・PART_ROLE_ALLOWED_STORES = {
    "A": ["螟ｧ豐ｻ蠎・, "蜷榊商螻倶ｸｭ蟾晏ｺ・],
    "B": ["蠕ｳ驥榊ｺ・],
    "C": ["蠕ｳ驥榊ｺ・],
    "D": ["螟ｩ逋ｽ讀咲伐蠎・],
    "E": ["螟ｧ豐ｻ蠎・],
}

# 縲檎､ｾ蜩｡1蜷搾ｼ九ヱ繝ｼ繝・蜷阪堺ｽ灘宛縺梧ｭ｣蠑上↓險ｱ蜿ｯ縺輔ｌ縺ｦ縺・ｋ蠎苓・縺ｨ縲∝ｯｾ蠢懊☆繧九ヱ繝ｼ繝亥ｽｹ蜑ｲ
# ・亥ｾｳ驥榊ｺ励・迚ｹ谿翫Ν繝ｼ繝ｫ縺ｮ縺溘ａ蛻･謇ｱ縺・ｼ・COMBO_STORE_PART_ROLES = {
    "螟ｧ豐ｻ蠎・: ["A", "E"],
    "蜷榊商螻倶ｸｭ蟾晏ｺ・: ["A"],
    "螟ｩ逋ｽ讀咲伐蠎・: ["D"],
}

TOKUSHIGE_STORE = "蠕ｳ驥榊ｺ・
TOKUSHIGE_PART_ROLES = ["B", "C"]

GENERAL_STORES = [s for s in STORES if s not in COMBO_STORE_PART_ROLES and s != TOKUSHIGE_STORE]

# 縲檎､ｾ蜩｡縲阪→縺励※蠎苓・縺ｮ蠢・ｦ∽ｽ灘宛繧ｫ繧ｦ繝ｳ繝医↓邂怜・縺輔ｌ繧句玄蛻・ｼ医ヱ繝ｼ繝医ｒ髯､縺擾ｼ・EMPLOYEE_TYPES = ["蠎鈴聞", "豁｣遉ｾ蜩｡", "蝌ｱ險・]

# 蠎苓・縺斐→縺ｮ1譌･縺ゅ◆繧頑怙螟ｧ莠ｺ謨ｰ繧ｭ繝｣繝・・(optimizer.py縺ｮ繝上・繝牙宛邏・→荳閾ｴ縺輔○繧・縲・# 縺薙ｌ縺ｯ繧ｽ繝ｫ繝舌・蛛ｴ縺ｮ螳滄圀縺ｮ莠ｺ蜩｡荳企剞(2蜷・3蜷・縺ｧ縺ゅｊ縲∫判髱｢繝ｻExcel縺ｮ陦ｨ遉ｺ陦梧焚縺ｨ縺ｯ
# 蛻･讎ょｿｵ縺ｧ縺ゅｋ(陦ｨ遉ｺ陦梧焚縺ｯ荳九・STORE_SLOT_ROWS繧貞盾辣ｧ)縲・STORE_MAX_HEADCOUNT = {
    store: (3 if store in (TOKUSHIGE_STORE, "蜷榊商螻倶ｸｭ蟾晏ｺ・, "螟ｩ逋ｽ讀咲伐蠎・) else 2) for store in STORES
}

# 逕ｻ髱｢荳翫・謇句虚邱ｨ髮・ユ繝ｼ繝悶Ν繝ｻExcel繧ｨ繧ｯ繧ｹ繝昴・繝医・荳｡譁ｹ縺ｧ縲∝・7蠎苓・繧剃ｸ蠕九・譫縲崎｡ｨ遉ｺ縺ｫ
# 邨ｱ荳縺吶ｋ縺溘ａ縺ｮ陦梧焚縲ょｮ滄圀縺ｮ莠ｺ蜩｡荳企剞(STORE_MAX_HEADCOUNT縲・蜷・3蜷・繧定ｶ・∴繧区棧縺ｯ縲・# 郢∝ｿ呎律縺ｮ蠢懈抄繧ｹ繧ｿ繝・ヵ霑ｽ蜉繧・脂遯√″隱ｿ謨ｴ縺ｮ荳譎ら噪縺ｪ蜿励￠逧ｿ縺ｨ縺励※遨ｺ谺・・縺ｾ縺ｾ陦ｨ遉ｺ縺輔ｌ繧・# (繧ｽ繝ｫ繝舌・蛛ｴ縺ｮ蠎苓・蛻･莠ｺ蜩｡荳企剞繧ｭ繝｣繝・・繧貞､画峩縺吶ｋ繧ゅ・縺ｧ縺ｯ縺ｪ縺・縲・STORE_SLOT_ROWS = 4

# 繝代・繝医・蠖ｹ蜑ｲ縺斐→縺ｮ縲∵悄髢灘・(16譌･縲懃ｿ梧怦15譌･)蜍､蜍呎律謨ｰ縺ｮ險ｱ螳ｹ遽・峇 (min, max)
# 繝ｪ繧ｽ繝ｼ繧ｹ繧剃ｽｿ縺・・繧峨○縺吶℃縺壹√°縺､譛菴朱剞縺ｯ遞ｼ蜒阪＆縺帙ｋ縺溘ａ縺ｮ繝上・繝牙宛邏・驥阪・繝翫Ν繝・ぅ)縲・PART_ROLE_WORKDAY_RANGE = {
    "A": (10, 13),  # 蟆ｾ貔､・亥､ｧ豐ｻ蠎・蜷榊商螻倶ｸｭ蟾晏ｺ暦ｼ・    "B": (12, 13),  # 荳榊虚驥趣ｼ亥ｾｳ驥榊ｺ暦ｼ・    "C": (11, 11),  # 蜑咲伐・亥ｾｳ驥榊ｺ暦ｼ会ｼ壼ｮ溽ｸｾ繝・・繧ｿ縺ｫ貅匁侠縺・1譌･蝗ｺ螳・    "D": (11, 14),  # 譟ｴ逕ｰ・亥､ｩ逋ｽ讀咲伐蠎暦ｼ・    "E": (12, 13),  # 驥朱％・亥､ｧ豐ｻ蠎暦ｼ・}

WEEKDAY_JP = ["譛・, "轣ｫ", "豌ｴ", "譛ｨ", "驥・, "蝨・, "譌･"]

NORMAL_HOURS = "10:00縲・9:00"
SHORT_HOURS = "10:00縲・7:00"


# ---------------------------------------------------------------------------
# 繧ｫ繝ｬ繝ｳ繝繝ｼ逕滓・
# ---------------------------------------------------------------------------

def get_period_dates(year: int, month: int) -> list[dt.date]:
    """蟇ｾ雎｡蟷ｴ譛医・縲・6譌･蟋九∪繧翫懃ｿ梧怦15譌･邱繧√阪・譌･莉倥Μ繧ｹ繝医ｒ霑斐☆縲・""
    start = dt.date(year, month, 16)
    if month == 12:
        next_year, next_month = year + 1, 1
    else:
        next_year, next_month = year, month + 1
    end = dt.date(next_year, next_month, 15)
    days = (end - start).days
    return [start + dt.timedelta(days=i) for i in range(days + 1)]


def last_tuesday_of_month(year: int, month: int) -> dt.date:
    """謖・ｮ壹＠縺滓圜譛医↓縺翫￠繧九梧怙邨ら↓譖懈律縲阪ｒ霑斐☆縲・""
    last_day = calendar.monthrange(year, month)[1]
    d = dt.date(year, month, last_day)
    while d.weekday() != 1:  # Tuesday == 1
        d -= dt.timedelta(days=1)
    return d


def is_weekend_or_holiday(d: dt.date) -> bool:
    """蝨滓屆繝ｻ譌･譖懊√∪縺溘・譌･譛ｬ縺ｮ逾晄律縺九←縺・°繧貞愛螳壹☆繧具ｼ亥悄譌･逾昴・驟咲ｽｮ蛻ｶ邏・畑・峨・""
    return d.weekday() in (5, 6) or jpholiday.is_holiday(d)


DEFAULT_SPECIAL_CLOSURE_LABEL = "縺顔寔繝ｻ蟷ｴ譛ｫ蟷ｴ蟋狗ｭ・
DEFAULT_FORCED_OPEN_LABEL = "閾ｨ譎ょ霧讌ｭ"

def classify_days(
    dates: list[dt.date],
    special_closure_dates: list[dt.date] | dict[dt.date, str] | None = None,
    forced_open_dates: list[dt.date] | dict[dt.date, str] | None = None,
) -> pd.DataFrame:
    """蜷・律縺ｮ譖懈律繝ｻ螳壻ｼ第律/迚ｹ蛻･蝟ｶ讌ｭ/騾壼ｸｸ蝟ｶ讌ｭ/迚ｹ蛻･莨第･ｭ譌･縺ｮ蛹ｺ蛻・ｒ蛻､螳壹☆繧九・
    繝ｫ繝ｼ繝ｫ:
      - 豈朱ｱ豌ｴ譖懈律: 螳壻ｼ第律
      - 轣ｫ譖懈律: 縲後◎縺ｮ證ｦ譛医・譛邨ら↓譖懈律縲阪・縺ｿ迚ｹ蛻･蝟ｶ讌ｭ(10-17譎・縲√◎繧御ｻ･螟悶・螳壻ｼ第律
      - 荳願ｨ倅ｻ･螟・ 騾壼ｸｸ蝟ｶ讌ｭ(10-19譎・
      - special_closure_dates 縺ｧ謖・ｮ壹＆繧後◆譌･(縺顔寔繝ｻ蟷ｴ譛ｫ蟷ｴ蟋狗ｭ峨・莉ｻ諢上・蜈ｨ蠎嶺ｸ譁我ｼ第･ｭ譌･)縺ｯ縲・        騾壼ｸｸ縺ｪ繧牙霧讌ｭ譌･(譛邨ら↓譖懈律縺ｮ遏ｭ邵ｮ蝟ｶ讌ｭ繧貞性繧)縺ｨ縺ｪ繧区律縺ｧ縺ゅ▲縺ｦ繧ょｼｷ蛻ｶ逧・↓
        縲檎音蛻･莨第･ｭ譌･縲阪→縺励※莨第･ｭ謇ｱ縺・↓縺吶ｋ(=蜈ｬ莨第律謨ｰ縺ｮ險育ｮ怜ｼ上↓縲檎音蛻･莨第･ｭ譌･謨ｰ縲・        縺ｨ縺励※蛻･譫縺ｧ蜉邂励＆繧後ｋ)縲よ里縺ｫ螳壻ｼ第律縺ｮ譌･繧呈欠螳壹＠縺ｦ繧ゆｺ碁㍾繧ｫ繧ｦ繝ｳ繝医・縺励↑縺・・      - special_closure_dates 縺ｯ繝ｪ繧ｹ繝・蜈ｨ譌･縲後♀逶・・蟷ｴ譛ｫ蟷ｴ蟋狗ｭ峨阪→縺・≧豎守畑逅・罰縺ｫ縺ｪ繧・
        縺ｧ繧ゅ＋譌･莉・ 逅・罰}縺ｮ霎樊嶌(萓・ {date(2026,12,31): "險育判蟷ｴ莨・})縺ｧ繧よｸ｡縺帙ｋ縲・        霎樊嶌縺ｧ逅・罰繧呈欠螳壹＠縺滓律縺ｯ縲］ote蛻励↓縺昴・逅・罰縺後◎縺ｮ縺ｾ縺ｾ蜿肴丐縺輔ｌ繧・        (萓・ 縲檎音蛻･莨第･ｭ(險育判蟷ｴ莨・縲・縲ゅ％繧後・陦ｨ遉ｺ繝ｻExcel蜃ｺ蜉帑ｸ翫・逅・罰繝ｩ繝吶Ν縺ｮ
        驕輔＞縺ｫ縺吶℃縺壹√＞縺壹ｌ繧りｧ｣譫仙ｯｾ雎｡縺ｮ蝟ｶ讌ｭ譌･縺九ｉ縺ｯ髯､螟悶＆繧後ｋ轤ｹ縺ｯ蜷後§縲・      - forced_open_dates 縺ｧ謖・ｮ壹＆繧後◆譌･縺ｯ縲∵悽譚･縺ｯ螳壻ｼ第律(豈朱ｱ豌ｴ譖懊・譛邨ら↓譖應ｻ･螟悶・
        轣ｫ譖・縺ｨ縺ｪ繧区律縺ｧ縺ゅ▲縺ｦ繧ゅ∝ｼｷ蛻ｶ逧・↓縲碁壼ｸｸ蝟ｶ讌ｭ縲・騾壼ｸｸ蝟ｶ讌ｭ譎る俣)縺ｨ縺励※
        蝟ｶ讌ｭ譌･謇ｱ縺・↓縺吶ｋ(=蟷ｴ譛ｫ蟷ｴ蟋区・縺代↓螟牙援逧・↓蝟ｶ讌ｭ縺吶ｋ縲∫ｭ峨・繧､繝ｬ繧ｮ繝･繝ｩ繝ｼ蟇ｾ蠢・縲・        special_closure_dates 縺ｨ forced_open_dates 縺ｮ荳｡譁ｹ縺ｫ蜷後§譌･縺悟性縺ｾ繧後ｋ蝣ｴ蜷医・
        forced_open_dates(蝟ｶ讌ｭ縺吶ｋ)繧貞━蜈医☆繧九よ欠螳壼ｽ｢蠑上・special_closure_dates
        縺ｨ蜷梧ｧ倥∵律莉倥Μ繧ｹ繝医∪縺溘・{譌･莉・ 逅・罰}縺ｮ霎樊嶌縺ｮ縺・★繧後ｂ貂｡縺帙ｋ縲・    """
    if isinstance(special_closure_dates, dict):
        special_labels = dict(special_closure_dates)
    else:
        special_labels = {d: DEFAULT_SPECIAL_CLOSURE_LABEL for d in (special_closure_dates or [])}
    special_set = set(special_labels.keys())

    if isinstance(forced_open_dates, dict):
        forced_open_labels = dict(forced_open_dates)
    else:
        forced_open_labels = {d: DEFAULT_FORCED_OPEN_LABEL for d in (forced_open_dates or [])}
    forced_open_set = set(forced_open_labels.keys())

    # 譛滄俣蜀・↓逋ｻ蝣ｴ縺励≧繧区圜譛医◎繧後◇繧後・譛邨ら↓譖懈律繧剃ｺ句燕險育ｮ・    months = sorted({(d.year, d.month) for d in dates})
    last_tue = {(y, m): last_tuesday_of_month(y, m) for y, m in months}

    rows = []
    for d in dates:
        weekday = d.weekday()  # Mon=0 ... Sun=6
        if weekday == 2:
            day_type, hours, note = "螳壻ｼ第律", "-", "豈朱ｱ豌ｴ譖懷ｮ壻ｼ・
        elif weekday == 1:
            if d == last_tue[(d.year, d.month)]:
                day_type, hours, note = "迚ｹ蛻･蝟ｶ讌ｭ(遏ｭ邵ｮ)", SHORT_HOURS, "譛亥・譛邨ら↓譖懈律"
            else:
                day_type, hours, note = "螳壻ｼ第律", "-", "轣ｫ譖懷ｮ壻ｼ・譛邨ら↓譖應ｻ･螟・"
        else:
            day_type, hours, note = "騾壼ｸｸ蝟ｶ讌ｭ", NORMAL_HOURS, ""

        is_special_closure = False
        if d in special_set and day_type != "螳壻ｼ第律":
            reason = special_labels.get(d) or DEFAULT_SPECIAL_CLOSURE_LABEL
            day_type, hours, note = "迚ｹ蛻･莨第･ｭ譌･", "-", f"迚ｹ蛻･莨第･ｭ({reason})"
            is_special_closure = True

        if d in forced_open_set:
            reason = forced_open_labels.get(d) or DEFAULT_FORCED_OPEN_LABEL
            day_type, hours, note = "騾壼ｸｸ蝟ｶ讌ｭ", NORMAL_HOURS, f"閾ｨ譎ょ霧讌ｭ({reason})"
            is_special_closure = False

        holiday_name = jpholiday.is_holiday_name(d) or ""
        rows.append(
            {
                "date": d,
                "weekday_jp": WEEKDAY_JP[weekday],
                "day_type": day_type,
                "hours": hours,
                "note": note,
                "is_closed": day_type in ("螳壻ｼ第律", "迚ｹ蛻･莨第･ｭ譌･"),
                "is_business_day": day_type not in ("螳壻ｼ第律", "迚ｹ蛻･莨第･ｭ譌･"),
                "is_weekend_or_holiday": is_weekend_or_holiday(d),
                "holiday_name": holiday_name,
                "is_special_closure": is_special_closure,
            }
        )
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 繧ｹ繧ｿ繝・ヵ繝槭せ繧ｿ・医ョ繝輔か繝ｫ繝亥､繝ｻ縺吶＄縺ｫ蜍穂ｽ懈､懆ｨｼ縺ｧ縺阪ｋ繧医≧縺ｫ逕ｨ諢擾ｼ・# ---------------------------------------------------------------------------

def default_staff_df() -> pd.DataFrame:
    """蠎鈴聞7蜷搾ｼ区ｭ｣遉ｾ蜩｡4蜷搾ｼ句亞險・蜷搾ｼ九ヱ繝ｼ繝・蜷・= 19蜷阪・繝・ヵ繧ｩ繝ｫ繝医・繧ｹ繧ｿ繧堤函謌舌☆繧九・
    遉ｾ髟ｷ縺ｯ莉墓ｧ倅ｸ翫す繝輔ヨ隕∝藤縺九ｉ螳悟・髯､螟悶・縺溘ａ縲√％縺ｮ繝槭せ繧ｿ縺ｫ縺ｯ荳蛻・匳蝣ｴ縺励↑縺・・    蠎鈴聞縺ｯ蜷・・縺ｮ蠎苓・縺ｫ螳悟・蝗ｺ螳夲ｼ井ｻ門ｺ苓・繝倥Ν繝嶺ｸ榊庄・昴ワ繝ｼ繝牙宛邏・ｼ峨〒縲・    蠎鈴聞莉･螟悶・豁｣遉ｾ蜩｡4蜷阪→蝌ｱ險・蜷阪・縺ｿ縺悟・蠎苓・縺ｸ縺ｮ繝倥Ν繝怜・蜍､繧定ｨｱ蜿ｯ縺輔ｌ繧九・    """
    rows = []

    # 蠎鈴聞7蜷搾ｼ壼推蠎苓・縺ｫ螳悟・蝗ｺ螳壹∽ｻ門ｺ励・繝ｫ繝嶺ｸ榊庄縲∵ｸｬ螳壹・蜉蟾･繧ｹ繧ｭ繝ｫ菫晄怏
    tencho_spec = [
        ("逕滄ｧ・, "遞ｲ豐｢蠎・),
        ("霎ｻ譛ｬ", "螟ｧ豐ｻ蠎・),
        ("蜀・伐", "蜷榊商螻倶ｸｭ蟾晏ｺ・),
        ("蟆乗棊", "譁ｰ陝ｹ豎溷ｺ・),
        ("蜉阯､", "蠕ｳ驥榊ｺ・),
        ("逕ｰ荳ｭ", "讌ｵ讌ｽ蠎・),
        ("髟ｷ轢ｬ", "螟ｩ逋ｽ讀咲伐蠎・),
    ]
    for i, (name, home) in enumerate(tencho_spec):
        rows.append(
            dict(
                staff_id=f"T{i+1:02d}",
                name=name,
                emp_type="蠎鈴聞",
                home_store=home,
                allowed_stores=[home],
                can_help=False,
                custom_allowed_stores=None,
                part_role=None,
                has_skill=True,
                min_workdays=None,
            )
        )

    # 豁｣遉ｾ蜩｡4蜷搾ｼ亥ｺ鈴聞莉･螟厄ｼ会ｼ夂悄逕ｰ縺ｯ蜈ｨ蠎苓・繝倥Ν繝怜庄閭ｽ縲・    # 闍･譫励・蜍､蜍吝庄閭ｽ蠎苓・縺後梧･ｵ讌ｽ蠎・荳ｻ謇螻・/蠕ｳ驥榊ｺ・螟ｩ逋ｽ讀咲伐蠎励阪・3蠎苓・縺ｮ縺ｿ縺ｫ髯仙ｮ・    # (遞ｲ豐｢蠎励・螟ｧ豐ｻ蠎励・蜷榊商螻倶ｸｭ蟾晏ｺ励・譁ｰ陝ｹ豎溷ｺ励∈縺ｮ驟咲ｽｮ縺ｯ螳悟・遖∵ｭ｢)縲・    # 荳ｭ譚代・遞ｲ豐｢蠎励∈縺ｮ驟咲ｽｮ繧堤ｦ∵ｭ｢(蜍､蜍吝庄閭ｽ蠎苓・縺九ｉ髯､螟・縲・    # 闍･譚ｾ縺ｯ蜍､蜍吝庄閭ｽ蠎苓・縺後梧眠陝ｹ豎溷ｺ・螟ｧ豐ｻ蠎・遞ｲ豐｢蠎・蜷榊商螻倶ｸｭ蟾晏ｺ励阪↓髯仙ｮ壹＆繧後・    # 縺九▽貂ｬ螳壹・蜉蟾･繧ｹ繧ｭ繝ｫ繧剃ｿ晄怏縺励↑縺・=蜷悟ｸｭ縺吶ｋ繧ｹ繧ｿ繝・ヵ縺ｫ隕√せ繧ｭ繝ｫ菫晄怏閠・∬ｩｳ邏ｰ縺ｯ
    # 蠎苓・蠢・ｦ∽ｽ灘宛縺ｮ繧ｹ繧ｭ繝ｫ隕∽ｻｶ縺梧里縺ｫ荳闊ｬ隗｣縺ｨ縺励※菫晁ｨｼ縺吶ｋ)縲・    seishain_spec = [
        ("逵溽伐", "蜷榊商螻倶ｸｭ蟾晏ｺ・, True, None),
        ("闍･譫・, "讌ｵ讌ｽ蠎・, True, ["讌ｵ讌ｽ蠎・, "蠕ｳ驥榊ｺ・, "螟ｩ逋ｽ讀咲伐蠎・]),
        ("荳ｭ譚・, "蠕ｳ驥榊ｺ・, True, [s for s in STORES if s != "遞ｲ豐｢蠎・]),
        ("闍･譚ｾ", "譁ｰ陝ｹ豎溷ｺ・, False, ["譁ｰ陝ｹ豎溷ｺ・, "螟ｧ豐ｻ蠎・, "遞ｲ豐｢蠎・, "蜷榊商螻倶ｸｭ蟾晏ｺ・]),
    ]
    for i, (name, home, has_skill, custom_stores) in enumerate(seishain_spec):
        rows.append(
            dict(
                staff_id=f"S{i+1:02d}",
                name=name,
                emp_type="豁｣遉ｾ蜩｡",
                home_store=home,
                allowed_stores=list(custom_stores) if custom_stores else list(STORES),
                can_help=True,
                custom_allowed_stores=list(custom_stores) if custom_stores else None,
                part_role=None,
                has_skill=has_skill,
                min_workdays=None,
            )
        )

    # 蝌ｱ險・蜷搾ｼ壽悄髢灘・蜍､蜍呎律謨ｰ縺ｮ遽・峇[min,max]縺ゅｊ縲らｫｹ蜀・・蜈ｨ7蠎苓・縺ｫ繝倥Ν繝怜・蜍､蜿ｯ閭ｽ縲・    shokutaku_spec = [
        ("蜷臥伐", "螟ｩ逋ｽ讀咲伐蠎・, True, 16, 18, None),
        ("遶ｹ蜀・, "螟ｩ逋ｽ讀咲伐蠎・, True, 18, 20, None),
        ("螻ｱ蟯｡", "遞ｲ豐｢蠎・, True, 18, 20, None),
    ]
    for i, (name, home, has_skill, min_days, max_days, custom_stores) in enumerate(shokutaku_spec):
        rows.append(
            dict(
                staff_id=f"K{i+1:02d}",
                name=name,
                emp_type="蝌ｱ險・,
                home_store=home,
                allowed_stores=list(custom_stores) if custom_stores else list(STORES),
                can_help=True,
                custom_allowed_stores=list(custom_stores) if custom_stores else None,
                part_role=None,
                has_skill=has_skill,
                min_workdays=min_days,
                max_workdays=max_days,
            )
        )

    # 繝代・繝・蜷搾ｼ壼共蜍吝庄閭ｽ蠎苓・縺檎ｵｶ蟇ｾ蛻ｶ髯・    part_spec = [
        ("蟆ｾ貔､", "A"),
        ("荳榊虚驥・, "B"),
        ("蜑咲伐", "C"),
        ("譟ｴ逕ｰ", "D"),
        ("驥朱％", "E"),
    ]
    for i, (name, role) in enumerate(part_spec):
        allowed = PART_ROLE_ALLOWED_STORES[role]
        rows.append(
            dict(
                staff_id=f"P{i+1:02d}",
                name=name,
                emp_type="繝代・繝・,
                home_store=allowed[0],
                allowed_stores=list(allowed),
                can_help=False,
                custom_allowed_stores=None,
                part_role=role,
                has_skill=False,
                min_workdays=None,
            )
        )

    df = pd.DataFrame(rows)
    # 蛟句挨蜈ｬ莨・谿九ｊ1譌･)縺ｮ蟶梧悍譌･縲よ里螳壹・譛ｪ謖・ｮ・None)=AI縺瑚・蜍輔〒譛驕ｩ驟榊・縺吶ｋ縲・    # 螳滄圀縺ｫ驕ｩ逕ｨ蟇ｾ雎｡縺ｨ縺ｪ繧九・縺ｯ emp_type 縺後悟ｺ鈴聞縲阪梧ｭ｣遉ｾ蜩｡縲阪・11蜷阪・縺ｿ縲・    df["preferred_off_date"] = None

    # 蝨滓律逾昴・迚ｹ螳壹・蠎苓・縺ｸ縺ｮ驟咲ｽｮ繧堤ｦ∵ｭ｢縺吶ｋ繧ｹ繧ｿ繝・ヵ(繝上・繝牙宛邏・縲・    #   闍･譚ｾ: 讀懈渊謚閭ｽ荳崎ｶｳ縺ｮ縺溘ａ蝨滓律逾昴・蜷榊商螻倶ｸｭ蟾晏ｺ鈴・鄂ｮ繧堤ｦ∵ｭ｢(蟷ｳ譌･縺ｮ縺ｿ)縲・    #        譁ｰ陝ｹ豎溷ｺ励↓縺､縺・※縺ｯ縲∽ｻ･蜑阪・縺薙％縺ｫ蜷ｫ繧√◆螳悟・遖∵ｭ｢縺縺｣縺溘′縲∽ｻ門ｺ苓・縺ｮ
    #        繧ｷ繝輔ヨ縺後←縺・＠縺ｦ繧ょ沂縺ｾ繧峨↑縺・=逵溘・莠ｺ蜩｡荳崎ｶｳ縺檎函縺倥ｋ)蝣ｴ蜷医↓髯舌ｊ
    #        驟榊ｱ槭ｒ隱阪ａ縺溘＞縺ｨ縺ｮ隕∵悍縺ｫ繧医ｊ縲∽ｸ九・weekend_holiday_avoid_stores縺ｸ
    #        遘ｻ蜍輔＠縺・螳悟・遖∵ｭ｢縺ｧ縺ｯ縺ｪ縺上・㍾縺・・繝翫Ν繝・ぅ莉倥″縺ｮ縲梧怙邨よ焔谿ｵ縲阪↓螟画峩)縲・    #   驥朱％: 螟ｧ豐ｻ蠎励・蝨滓律逾晏・蜍､繧堤ｦ∵ｭ｢(蟷ｳ譌･縺ｮ縺ｿ遞ｼ蜒・縲・    df["weekend_holiday_forbidden_stores"] = [[] for _ in range(len(df))]
    df.loc[df["name"] == "闍･譚ｾ", "weekend_holiday_forbidden_stores"] = df.loc[
        df["name"] == "闍･譚ｾ", "weekend_holiday_forbidden_stores"
    ].apply(lambda _: ["蜷榊商螻倶ｸｭ蟾晏ｺ・])
    df.loc[df["name"] == "驥朱％", "weekend_holiday_forbidden_stores"] = df.loc[
        df["name"] == "驥朱％", "weekend_holiday_forbidden_stores"
    ].apply(lambda _: ["螟ｧ豐ｻ蠎・])

    # 蝨滓律逾昴・縲悟次蜑・咲音螳壹・蠎苓・縺ｸ縺ｮ驟咲ｽｮ繧帝∩縺代◆縺・′縲∽ｻ門ｺ苓・縺ｮ莠ｺ蜩｡荳崎ｶｳ縺後←縺・＠縺ｦ繧・    # 隗｣豸医〒縺阪↑縺・ｴ蜷医↓髯舌ｊ驟榊ｱ槭ｒ險ｱ螳ｹ縺吶ｋ縲√→縺・≧繧ｽ繝輔ヨ蛻ｶ邏・驥阪＞繝壹リ繝ｫ繝・ぅ)縲・    # weekend_holiday_forbidden_stores(螳悟・遖∵ｭ｢繝ｻ繝上・繝牙宛邏・縺ｨ縺ｯ逡ｰ縺ｪ繧翫・    # 逵溘↓蠢・ｦ√↑蝣ｴ蜷・=荳崎ｶｳ隗｣豸・縺ｫ縺ｯ繧ｽ繝ｫ繝舌・縺御ｾ句､也噪縺ｫ縺薙％縺ｸ驟榊ｱ槭〒縺阪ｋ縲・    #   闍･譚ｾ: 讀懈渊謚閭ｽ縺ｮ隕ｳ轤ｹ縺九ｉ譁ｰ陝ｹ豎溷ｺ励・譛ｬ譚･驕ｿ縺代◆縺・′縲∽ｻ門ｺ苓・縺ｮ繧ｷ繝輔ヨ縺後←縺・＠縺ｦ繧・    #        蝓九∪繧峨↑縺・ｴ蜷医・縲梧怙邨よ焔谿ｵ縲阪→縺励※驟榊ｱ槭ｒ險ｱ螳ｹ縺吶ｋ(2026/9譎らせ縺ｮ隕∵悍)縲・    df["weekend_holiday_avoid_stores"] = [[] for _ in range(len(df))]
    df.loc[df["name"] == "闍･譚ｾ", "weekend_holiday_avoid_stores"] = df.loc[
        df["name"] == "闍･譚ｾ", "weekend_holiday_avoid_stores"
    ].apply(lambda _: ["譁ｰ陝ｹ豎溷ｺ・])

    # 蜃ｺ蜍､譌･縺ｫ縺ｧ縺阪ｋ髯舌ｊ縺薙・蠎苓・縺ｸ驟榊ｱ槭＠縺ｦ縺ｻ縺励＞縲√→縺・≧繧ｽ繝輔ヨ縺ｪ蜆ｪ蜈亥ｺ苓・縲・    #   蜷臥伐: 螟ｩ逋ｽ讀咲伐蠎励ｒ蜆ｪ蜈磯・螻・莉門ｺ励・繝ｫ繝励・繧・・繧貞ｾ励↑縺・ｴ蜷医・縺ｿ霆ｽ蠕ｮ縺ｪ繝壹リ繝ｫ繝・ぅ)縲・    #   螻ｱ蟯｡: 荳ｻ謇螻槭・遞ｲ豐｢蠎励ｒ譛蜆ｪ蜈磯・螻・遶ｹ蜀・・闍･譚ｾ遲峨・繝倥Ν繝励ｈ繧雁━蜈医＆縺帙ｋ)縲・    #   蟆ｾ貔､: 荳ｻ謇螻槭・蜷榊商螻倶ｸｭ蟾晏ｺ励ｒ蜆ｪ蜈磯・螻槭よ勸谿ｵ縺ｯ蜷榊商螻倶ｸｭ蟾晏ｺ励↓逡吶ａ縲∝､ｧ豐ｻ蠎励∈縺ｮ
    #        謖ｯ譖ｿ(霎ｻ譛ｬ+蟆ｾ貔､縺ｮ邨・粋縺・縺ｯ莉門ｺ苓・縺ｮ莠ｺ蜩｡荳崎ｶｳ縺後←縺・＠縺ｦ繧りｧ｣豸医〒縺阪↑縺・    #        蝣ｴ蜷医・縲梧怙邨よ焔谿ｵ縲阪→縺励※縺ｮ縺ｿ逋ｺ蜍輔＆縺帙◆縺・2026/9譎らせ縺ｮ隕∵悍)縲・    df["preferred_store"] = None
    df.loc[df["name"] == "蜷臥伐", "preferred_store"] = "螟ｩ逋ｽ讀咲伐蠎・
    df.loc[df["name"] == "螻ｱ蟯｡", "preferred_store"] = "遞ｲ豐｢蠎・
    df.loc[df["name"] == "蟆ｾ貔､", "preferred_store"] = "蜷榊商螻倶ｸｭ蟾晏ｺ・

    # 蝨滓屆繝ｻ譌･譖・逾晄律縺ｯ蜷ｫ縺ｾ縺ｪ縺・縺ｫ髯舌ｊ迚ｹ螳壹・蠎苓・縺ｸ縺ｮ驟咲ｽｮ繧堤ｦ∵ｭ｢縺吶ｋ繧ｹ繧ｿ繝・ヵ(繝上・繝牙宛邏・縲・    #   闍･譚ｾ: 蝨滓律縺ｯ螟ｧ豐ｻ蠎励ｒ蜆ｪ蜈磯・鄂ｮ縺ｨ縺吶ｋ縺溘ａ縲∵ｮ九ｋ遞ｲ豐｢蠎励ｒ遖∵ｭ｢(=蝨滓律縺ｯ螟ｧ豐ｻ蠎怜崋螳・縲・    #        (譁ｰ陝ｹ豎溷ｺ励・ weekend_holiday_forbidden_stores 縺ｧ譌｢縺ｫ遖∵ｭ｢貂医∩)
    df["saturday_sunday_forbidden_stores"] = [[] for _ in range(len(df))]
    df.loc[df["name"] == "闍･譚ｾ", "saturday_sunday_forbidden_stores"] = df.loc[
        df["name"] == "闍･譚ｾ", "saturday_sunday_forbidden_stores"
    ].apply(lambda _: ["遞ｲ豐｢蠎・])

    return df


# ---------------------------------------------------------------------------
# 繧ｹ繧ｿ繝・ヵ蜷咲ｰｿ縺ｮ豌ｸ邯壼喧(Google繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝・
# ---------------------------------------------------------------------------

def load_staff_roster() -> pd.DataFrame | None:
    """Google繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医↓菫晏ｭ俶ｸ医∩縺ｮ繧ｹ繧ｿ繝・ヵ蜷咲ｰｿ繧定ｪｭ縺ｿ霎ｼ繧縲・
    菫晏ｭ倥ョ繝ｼ繧ｿ縺檎┌縺・謗･邯壹お繝ｩ繝ｼ縺ｮ蝣ｴ蜷医・ None 繧定ｿ斐☆
    (蜻ｼ縺ｳ蜃ｺ縺怜・縺ｯ default_staff_df() 縺ｫ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ縺吶ｋ縺薙→)縲・    """
    try:
        return sheets_store.load_staff_roster()
    except Exception:
        return None


def save_staff_roster(staff_df: pd.DataFrame) -> None:
    """繧ｹ繧ｿ繝・ヵ蜷咲ｰｿ繧竪oogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医∈菫晏ｭ倥☆繧九・""
    try:
        sheets_store.save_staff_roster(staff_df)
    except Exception:
        pass


# 譛滄俣蜀・↓譛菴・蝗槭・迚ｹ螳壹・2蜷阪ｒ蜷御ｸ蠎苓・縺ｧ蜷梧律蜍､蜍吶＆縺帙ｋ縲√→縺・≧邨・∩蜷医ｏ縺幄ｦ∽ｻｶ縲・# (荳崎ｶｳ譎ゅ・繧ｽ繝ｫ繝舌・繧定誠縺ｨ縺輔★驥阪＞繝壹リ繝ｫ繝・ぅ縺ｧ隴ｦ蜻翫☆繧九た繝輔ヨ蛹悶ワ繝ｼ繝牙宛邏・
REQUIRED_PAIR_WORKDAYS = [
    {"names": ("霎ｻ譛ｬ", "蟆ｾ貔､"), "store": "螟ｧ豐ｻ蠎・, "min_occurrences": 1},
]

# 迚ｹ螳壽律繝ｻ迚ｹ螳壼ｺ苓・縺ｮ驟咲ｽｮ繧呈焔蜍輔〒遒ｺ螳壹＆縺帙ｋ(繝上・繝牙宛邏・縲よ欠螳壹＠縺滓ｰ丞錐繧貞ｿ・★蜃ｺ蜍､
# 縺輔○縲√◎繧御ｻ･螟悶・繧ｹ繧ｿ繝・ヵ縺ｯ蜷梧律蜷悟ｺ苓・縺ｸ縺ｮ驟咲ｽｮ繧堤ｦ∵ｭ｢縺吶ｋ(=繝ｭ繝ｼ繧ｹ繧ｿ繝ｼ繧貞ｮ悟・蝗ｺ螳・縲・# 蠢・ｦ∽ｺｺ謨ｰ縺ｫ貅縺溘↑縺・欠螳壹・蝣ｴ蜷医・縲√◎縺ｮ蛻・□縺台ｸ崎ｶｳ繧ｹ繝ｩ繝・け縺檎ｫ九■縲∽ｸ崎ｶｳ繧｢繝ｩ繝ｼ繝医↓
# 蜿肴丐縺輔ｌ繧・繧ｽ繝ｫ繝舌・縺ｯ關ｽ縺｡縺ｪ縺・縲ょｯｾ雎｡譛滄俣縺ｫ隧ｲ蠖捺律縺悟性縺ｾ繧後↑縺・ｴ蜷医・辟｡隕悶＆繧後ｋ縲・#
# 驕主悉縺ｫ縲・026/10/10縺ｮ蠕ｳ驥榊ｺ・荳ｭ譚・螻ｱ蟯｡縲∫ｨｲ豐｢蠎・逕滄ｧ・蜷阪・縺ｿ縲阪→縺・≧荳譎ら噪縺ｪ
# 繝・ヰ繝・げ逕ｨ縺ｮ蝗ｺ螳壽欠螳壹′縺薙％縺ｫ蜈･縺｣縺ｦ縺・◆縺後√◎縺ｮ蠕瑚ｿｽ蜉縺励◆縲檎ｨｲ豐｢蠎励・螻ｱ蟯｡繧・# 譛蜆ｪ蜈磯・鄂ｮ縺吶ｋ縲阪→縺・≧蜆ｪ蜈磯・ｽ阪Ν繝ｼ繝ｫ縺ｨ遏帷崟縺吶ｋ(螻ｱ蟯｡繧貞ｼｷ蛻ｶ逧・↓蠕ｳ驥榊ｺ励∈
# 騾√▲縺ｦ縺励∪縺・◆繧・縺薙→縺悟愛譏弱＠縺溘◆繧√∵￡荵・噪縺ｪ迚ｹ萓九→縺励※縺ｯ荳埼←蛻・→蛻､譁ｭ縺・# 遐ｴ譽・＠縺溘る壼ｸｸ縺ｯ遨ｺ縺ｮ縺ｾ縺ｾ縺ｫ縺励※縺翫″縲∫音螳壽律繧呈焔蜍輔〒蝗ｺ螳壹＠縺溘＞荳譎ら噪縺ｪ
# 髴隕√′逕溘§縺溷ｴ蜷医・縺ｿ縲∵悄髢薙ｒ蛹ｺ蛻・▲縺ｦ荳譎ら噪縺ｫ霑ｽ蜉縺吶ｋ縺薙→縲・MANUAL_STORE_ASSIGNMENTS: list[dict] = []


def derive_allowed_stores(row) -> list[str]:
    """emp_type繝ｻ驛ｨ鄂ｲ蝗ｺ螳壹Ν繝ｼ繝ｫ繝ｻ蛟句挨縺ｮ蜍､蜍吝庄閭ｽ蠎苓・蛻ｶ髯舌°繧牙共蜍吝庄閭ｽ蠎苓・繧貞ｰ主・縺吶ｋ縲・
    蜆ｪ蜈磯・ｽ・
      1. 蠎鈴聞: 閾ｪ蠎苓・蝗ｺ螳・繝上・繝牙宛邏・縲ゆｻ悶・險ｭ螳壹↓髢｢繧上ｉ縺・home_store 1蠎苓・縺ｮ縺ｿ縲・      2. 繝代・繝・ 蠖ｹ蜑ｲ縺斐→縺ｮ邨ｶ蟇ｾ蛻ｶ髯・蜍､蜍吝庄閭ｽ蠎苓・縺ｮ邨ｶ蟇ｾ蛻ｶ髯・縲・      3. custom_allowed_stores 縺瑚ｨｭ螳壹＆繧後※縺・ｋ蝣ｴ蜷・萓・ 闍･譚ｾ繝ｻ遶ｹ蜀・: 縺昴・蠎苓・
         繝ｪ繧ｹ繝医↓髯仙ｮ・繝上・繝牙宛邏・縲ゆｻ門ｺ苓・繝倥Ν繝苓ｨｱ蜿ｯ繝輔Λ繧ｰ繧医ｊ蜆ｪ蜈医＆繧後ｋ縲・      4. 縺昴ｌ莉･螟・豁｣遉ｾ蜩｡繝ｻ蝌ｱ險励・讓呎ｺ悶こ繝ｼ繧ｹ): 莉門ｺ苓・繝倥Ν繝苓ｨｱ蜿ｯ繝輔Λ繧ｰ縺ｫ蠕薙＞縲・         蜈ｨ蠎苓・ or 荳ｻ謇螻槫ｺ苓・縺ｮ縺ｿ縲・    """

    def _get(key, default=None):
        if isinstance(row, dict):
            return row.get(key, default)
        return getattr(row, key, default)

    emp_type = _get("emp_type")
    home_store = _get("home_store")
    if emp_type == "蠎鈴聞":
        return [home_store]
    if emp_type == "繝代・繝・:
        part_role = _get("part_role")
        return list(PART_ROLE_ALLOWED_STORES.get(part_role, [home_store]))
    custom_stores = _get("custom_allowed_stores")
    if isinstance(custom_stores, list) and len(custom_stores) > 0:
        return list(custom_stores)
    can_help = _get("can_help", True)
    return list(STORES) if can_help else [home_store]


def default_requests_df() -> pd.DataFrame:
    """蟶梧悍莨代・譛臥ｵｦ逕ｳ隲九・蜈･蜉帙ユ繝ｼ繝悶Ν縺ｮ遨ｺ髮帛ｽ｢縲・""
    return pd.DataFrame(
        {
            "staff_id": pd.Series(dtype="str"),
            "name": pd.Series(dtype="str"),
            "date": pd.Series(dtype="object"),
            "kind": pd.Series(dtype="str"),  # 蟶梧悍莨・/ 邨ｶ蟇ｾ莨・/ 譛臥ｵｦ逕ｳ隲・        }
    )


REQUEST_KINDS = ["蟶梧悍莨・, "邨ｶ蟇ｾ莨・, "譛臥ｵｦ逕ｳ隲・]
HARD_OFF_KINDS = {"邨ｶ蟇ｾ莨・, "譛臥ｵｦ逕ｳ隲・}
SOFT_OFF_KINDS = {"蟶梧悍莨・}

# 譌ｧ陦ｨ險倥°繧峨・繝・・繧ｿ遘ｻ陦檎畑繧ｨ繧､繝ｪ繧｢繧ｹ縲る℃蜴ｻ縺ｫ菫晏ｭ倥＆繧後◆繝ｭ繧ｰ/CSV縺ｫ縺ｯ縲・#   (1) 縲梧怏莨代崎｡ｨ險・迴ｾ陦後・縲梧怏邨ｦ縲崎｡ｨ險倥∈謾ｹ遘ｰ縺吶ｋ蜑・
#   (2) 縲梧怏邨ｦ遒ｺ螳壹榊玄蛻・迴ｾ陦後〒縺ｯ蟒・ｭ｢縺輔ｌ縲梧怏邨ｦ逕ｳ隲九阪↓邨ｱ蜷域ｸ医∩)
# 縺ｮ縺・★繧後°縺後◎縺ｮ縺ｾ縺ｾ谿九▲縺ｦ縺・ｋ蜿ｯ閭ｽ諤ｧ縺後≠繧九◆繧√∬ｪｭ縺ｿ霎ｼ縺ｿ譎ゅ↓蠢・★縺薙・陦ｨ繧・# 騾壹＠縺ｦ迴ｾ陦後・豁｣隕剰｡ｨ險・縲梧怏邨ｦ逕ｳ隲九・縺ｸ豁｣隕丞喧縺吶ｋ(=菫晏ｭ俶ｸ医∩縺ｮ邨ｶ蟇ｾ蛻ｶ邏・′縲・# 陦ｨ險伜､画峩繝ｻ蛹ｺ蛻・ｵｱ蜷医□縺代〒繝上・繝牙宛邏・°繧牙､悶ｌ縺ｦ縺励∪縺・ｺ区腐繧帝亟縺・縲・_LEGACY_KIND_ALIASES = {
    "譛我ｼ醍筏隲・: "譛臥ｵｦ逕ｳ隲・,
    "譛我ｼ醍｢ｺ螳・: "譛臥ｵｦ逕ｳ隲・,  # 譌ｧ陦ｨ險・縺九▽ 蟒・ｭ｢蛹ｺ蛻・莠碁㍾縺ｮ遘ｻ陦・
    "譛臥ｵｦ遒ｺ螳・: "譛臥ｵｦ逕ｳ隲・,  # 蛹ｺ蛻・ｵｱ蜷・縲梧怏邨ｦ遒ｺ螳壹阪・蟒・ｭ｢縺励梧怏邨ｦ逕ｳ隲九阪↓荳譛ｬ蛹・
}


def _normalize_kind_value(kind: str) -> str:
    """kind譁・ｭ怜・繧偵∵立陦ｨ險・譛我ｼ・繝ｻ蟒・ｭ｢蛹ｺ蛻・譛臥ｵｦ遒ｺ螳・繧貞性繧√※迴ｾ陦後・豁｣隕剰｡ｨ險・    (譛臥ｵｦ逕ｳ隲・縺ｸ豁｣隕丞喧縺吶ｋ縲・""
    return _LEGACY_KIND_ALIASES.get(kind, kind)


# ---------------------------------------------------------------------------
# 蟶梧悍莨代・譛我ｼ大・蜉帙ユ繝ｼ繝悶Ν縺ｮ繝ｭ繝ｼ繧ｫ繝ｫ豌ｸ邯壼喧
# ---------------------------------------------------------------------------

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
SAVED_KYUKA_PATH = os.path.join(DATA_DIR, "saved_kyuka.csv")  # 譌ｧ繝ｻ譛亥ｺｦ髱槫・髮｢縺ｮ繝・ヵ繧ｩ繝ｫ繝医ヱ繧ｹ(蠕梧婿莠呈鋤逕ｨ)


def saved_kyuka_path_for(year: int, month: int) -> str:
    """譛亥ｺｦ蛻･(繧ｷ繝輔ヨ譛滄俣髢句ｧ句ｹｴ譛・縺ｮ蟶梧悍莨代・譛我ｼ代ョ繝ｼ繧ｿ縺ｮ菫晏ｭ倥ヱ繧ｹ繧定ｿ斐☆縲・
    譛ｪ譚･縺ｮ譛亥ｺｦ(2縲・繝ｶ譛亥・遲・縺ｮ蟶梧悍莨代ｒ蜈郁｡悟・蜉帙・菫晏ｭ倥〒縺阪ｋ繧医≧縲∵怦蠎ｦ縺斐→縺ｫ
    螳悟・縺ｫ迢ｬ遶九＠縺溘ヵ繧｡繧､繝ｫ縺ｧ邂｡逅・☆繧九・    """
    return os.path.join(DATA_DIR, f"saved_kyuka_{int(year)}_{int(month):02d}.csv")


def save_requests_to_disk(requests_df: pd.DataFrame, path: str = SAVED_KYUKA_PATH) -> None:
    """蟶梧悍莨代・譛我ｼ大・蜉帙ユ繝ｼ繝悶Ν繧偵Ο繝ｼ繧ｫ繝ｫCSV縺ｸ蜊ｳ譎ゆｿ晏ｭ倥☆繧九・
    逕ｻ髱｢縺ｧ縺ｮ謇句・蜉帙・邱ｨ髮・，SV荳諡ｬ繧､繝ｳ繝昴・繝医・縺・★繧後・螟画峩蠕後↓繧ょ他縺ｳ蜃ｺ縺吶％縺ｨ縺ｧ縲・    繧｢繝励Μ縺ｮ蜀崎ｵｷ蜍輔ｄ蛻･繝悶Λ繧ｦ繧ｶ繝ｻ蛻･遶ｯ譛ｫ縺九ｉ縺ｮ繧｢繧ｯ繧ｻ繧ｹ譎ゅ↓繧ょ・螳ｹ縺悟ｾｩ蜈・＆繧後ｋ縲・    """
    os.makedirs(os.path.dirname(path), exist_ok=True)
    out = requests_df.copy()
    if not out.empty:
        out["date"] = out["date"].apply(lambda d: d.isoformat() if isinstance(d, dt.date) else d)
    out.reindex(columns=["staff_id", "name", "date", "kind"]).to_csv(path, index=False, encoding="utf-8-sig")


def load_requests_from_disk(path: str = SAVED_KYUKA_PATH) -> pd.DataFrame:
    """繝ｭ繝ｼ繧ｫ繝ｫ菫晏ｭ俶ｸ医∩縺ｮ蟶梧悍莨代・譛我ｼ代ョ繝ｼ繧ｿ繧定ｪｭ縺ｿ霎ｼ繧縲ょｭ伜惠縺励↑縺代ｌ縺ｰ遨ｺ縺ｮ髮帛ｽ｢繧定ｿ斐☆縲・""
    if not os.path.exists(path):
        return default_requests_df()
    try:
        df = pd.read_csv(path, dtype={"staff_id": str, "name": str, "kind": str}, encoding="utf-8-sig")
    except Exception:
        return default_requests_df()
    if df.empty or "date" not in df.columns:
        return default_requests_df()
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    df = df.dropna(subset=["date"])
    df = df.reindex(columns=["staff_id", "name", "date", "kind"]).reset_index(drop=True)
    if "kind" in df.columns:
        df["kind"] = df["kind"].apply(lambda k: _normalize_kind_value(k) if isinstance(k, str) else k)
    return df


def clear_saved_requests(path: str = SAVED_KYUKA_PATH) -> None:
    """繝ｭ繝ｼ繧ｫ繝ｫ菫晏ｭ俶ｸ医∩縺ｮ蟶梧悍莨代・譛我ｼ代ョ繝ｼ繧ｿ繧貞炎髯､縺吶ｋ(繝ｪ繧ｻ繝・ヨ讖溯・逕ｨ)縲・""
    if os.path.exists(path):
        os.remove(path)


# ---------------------------------------------------------------------------
# 蟶梧悍莨代・譛我ｼ代・縲悟句挨逕ｳ隲九・霑ｽ險伜梛繝ｭ繧ｰ縲咲ｮ｡逅・#
# 譌ｧ譁ｹ蠑・save_requests_to_disk)縺ｯ縲∫判髱｢荳翫・蜈ｨ繧ｹ繧ｿ繝・ヵ蛻・・蜈･蜉帙ユ繝ｼ繝悶Ν繧呈ｯ主屓
# 縺ｾ繧九＃縺ｨ1縺､縺ｮCSV縺ｸ荳頑嶌縺堺ｿ晏ｭ倥＠縺ｦ縺・◆縲ゅ％縺ｮ縺溘ａ縲、縺輔ｓ縺ｮ遶ｯ譛ｫ縺悟商縺・せ繝翫ャ繝・# 繧ｷ繝ｧ繝・ヨ繧剃ｿ晄戟縺励◆縺ｾ縺ｾ(=B縺輔ｓ縺瑚ｿｽ蜉縺励◆逕ｳ隲九ｒ縺ｾ縺遏･繧峨↑縺・憾諷九・縺ｾ縺ｾ)菴輔ｉ縺・# 縺ｮ逅・罰縺ｧ蜀堺ｿ晏ｭ倥＆繧後ｋ縺ｨ縲。縺輔ｓ縺ｮ逕ｳ隲九＃縺ｨ豸医∴縺ｦ縺励∪縺・ｫｶ蜷医′襍ｷ縺薙ｊ蠕励◆縲・#
# 譁ｰ譁ｹ蠑上〒縺ｯ縲∫筏隲・莉ｶ縺斐→縺ｫ縲瑚ｪｰ縺後・縺・▽繝ｻ菴輔ｒ逕ｳ隲九＠縺溘°縲阪ｒ1陦後→縺励※繝ｭ繧ｰCSV
# (data/kyuka_requests_{year}_{month:02d}.csv)縺ｸ霑ｽ險倥☆繧九・縺ｿ縺ｨ縺励∵里蟄倩｡後・
# 隱ｭ縺ｿ霎ｼ縺ｿ竊呈嶌縺肴鋤縺遺・蜈ｨ菴謎ｿ晏ｭ倥・荳蛻・｡後ｏ縺ｪ縺・ゅ％繧後↓繧医ｊ縲∬､・焚繝悶Λ繧ｦ繧ｶ/遶ｯ譛ｫ縺九ｉ
# 蜷梧凾縺ｫ騾∽ｿ｡縺輔ｌ縺ｦ繧ゆｺ偵＞縺ｮ陦後ｒ荳頑嶌縺阪・豸亥悉縺吶ｋ縺薙→縺後↑縺上↑繧・=霑ｽ險倥・蜴溽炊逧・↓
# 陦晉ｪ√＠縺ｪ縺・縲ょ叙豸医ｂ縲悟叙豸医咲ｨｮ蛻･縺ｮ陦後ｒ霑ｽ險倥☆繧玖ｫ也炊蜑企勁縺ｨ縺励※謇ｱ縺・・# 迴ｾ蝨ｨ縺ｮ譛牙柑縺ｪ逕ｳ隲倶ｸ隕ｧ縺ｯ縲・繧ｹ繧ｿ繝・ヵ蜷・ 譌･莉・縺斐→縺ｫ繝ｭ繧ｰ縺ｮ譛譁ｰ陦後ｒ謗｡逕ｨ縺吶ｋ縺薙→
# 縺ｧ驛ｽ蠎ｦ蜀肴ｧ狗ｯ峨☆繧・=繧､繝吶Φ繝医た繝ｼ繧ｷ繝ｳ繧ｰ/霑ｽ險倥Ο繧ｰ縺ｮ閠・∴譁ｹ)縲・# ---------------------------------------------------------------------------

KYUKA_LOG_COLUMNS = ["staff_name", "date", "request_type", "updated_at"]
CANCELLED_REQUEST_TYPE = "蜿匁ｶ・  # 繝ｭ繧ｰ荳翫・蜿匁ｶ・隲也炊蜑企勁)繝槭・繧ｫ繝ｼ


# 豕ｨ險・ 莉･蜑阪・縺薙％縺ｫ縲∵怦蠎ｦ蛻･繝ｭ繝ｼ繧ｫ繝ｫCSV/JSON繝輔ぃ繧､繝ｫ縺ｸ菫晏ｭ倥☆繧句ｮ溯｣・# (kyuka_log_path_for繝ｻspecial_days_path_for繝ｻ繝輔ぃ繧､繝ｫ繝ｭ繝・け繝ｻ
# append_kyuka_request遲・縺後≠縺｣縺溘′縲ヾtreamlit Community Cloud縺ｯ繧｢繝励Μ縺ｮ
# 蜀崎ｵｷ蜍輔・蜀阪ョ繝励Ο繧､縺ｮ縺溘・縺ｫ繝ｭ繝ｼ繧ｫ繝ｫ繝輔ぃ繧､繝ｫ繧ｷ繧ｹ繝・Β縺ｸ縺ｮ譖ｸ縺崎ｾｼ縺ｿ繧貞・縺ｦ
# 豸亥悉縺励※縺励∪縺・=豌ｸ邯壼喧縺輔ｌ縺ｪ縺・縺薙→縺檎｢ｺ隱阪＆繧後◆縺溘ａ縲；oogle繧ｹ繝励Ξ繝・ラ
# 繧ｷ繝ｼ繝医∈菫晏ｭ倥☆繧区婿蠑・sheets_store.py)縺ｫ鄂ｮ縺肴鋤縺医◆縲ゆｻ･荳九・髢｢謨ｰ鄒､縺ｯ縲・# 蜷悟錐縺ｮ髢｢謨ｰ縺ｨ縺励※蠑輔″邯壹″蜻ｼ縺ｳ蜃ｺ縺帙ｋ縺後∝・驛ｨ縺ｧ縺ｯsheets_store邨檎罰縺ｧGoogle
# 繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医ｒ隱ｭ縺ｿ譖ｸ縺阪☆繧九・
import sheets_store


def _sheets_setup_error(exc: Exception) -> str:
    return (
        "Google繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医∈縺ｮ謗･邯壹↓螟ｱ謨励＠縺ｾ縺励◆縲４ecrets縺ｮ險ｭ螳・
        "([gcp_service_account]繝ｻspreadsheet_url)繧偵＃遒ｺ隱阪￥縺縺輔＞縲・
        f"(隧ｳ邏ｰ: {exc})"
    )


def load_special_days(year: int, month: int) -> tuple[dict[dt.date, str], dict[dt.date, str]]:
    """謖・ｮ壽怦蠎ｦ縺ｮ迚ｹ蛻･莨第･ｭ譌･繝ｻ閾ｨ譎ょ霧讌ｭ譌･縺ｮ險ｭ螳壹ｒ縲；oogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医°繧芽ｪｭ縺ｿ霎ｼ繧縲・
    譛ｪ險ｭ螳壹・謗･邯壹お繝ｩ繝ｼ縺ｮ蝣ｴ蜷医・縲∽ｸ｡譁ｹ縺ｨ繧らｩｺ縺ｮ霎樊嶌(=菴輔ｂ謖・ｮ壹↑縺・繧定ｿ斐☆
    (繧｢繝励Μ蜈ｨ菴薙ｒ繧ｯ繝ｩ繝・す繝･縺輔○縺ｪ縺・◆繧√・螳牙・蛛ｴ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ)縲・    """
    try:
        df = sheets_store.load_special_days_all()
    except Exception:
        return {}, {}
    if df.empty:
        return {}, {}
    df = df[(df["year"].astype(str) == str(year)) & (df["month"].astype(str) == str(month))]
    special_closure_map: dict[dt.date, str] = {}
    forced_open_map: dict[dt.date, str] = {}
    for _, r in df.iterrows():
        try:
            d = dt.date.fromisoformat(r["date"])
        except (TypeError, ValueError):
            continue
        if r["category"] == "special_closure":
            special_closure_map[d] = r["reason"]
        elif r["category"] == "forced_open":
            forced_open_map[d] = r["reason"]
    return special_closure_map, forced_open_map


def save_special_days(
    year: int,
    month: int,
    special_closure_map: dict[dt.date, str],
    forced_open_map: dict[dt.date, str],
) -> None:
    """謖・ｮ壽怦蠎ｦ縺ｮ迚ｹ蛻･莨第･ｭ譌･繝ｻ閾ｨ譎ょ霧讌ｭ譌･縺ｮ險ｭ螳壹ｒGoogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医∈菫晏ｭ倥☆繧九・""
    sheets_store.save_special_days(year, month, special_closure_map, forced_open_map)


def append_kyuka_request(staff_name: str, date: dt.date, request_type: str, _period_key: str | None = None) -> None:
    """蛟句挨縺ｮ蟶梧悍莨代・譛臥ｵｦ逕ｳ隲・縺ｾ縺溘・蜿匁ｶ・繧・莉ｶ縲；oogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医・繝ｭ繧ｰ縺ｸ霑ｽ險倥☆繧九・
    `_period_key` 縺ｯ譌ｧ繝ｭ繝ｼ繧ｫ繝ｫ繝輔ぃ繧､繝ｫ譁ｹ蠑上・蜷肴ｮ九・蠑墓焚縺ｧ縲ヾheets譁ｹ蠑上〒縺ｯ
    菴ｿ逕ｨ縺励↑縺・蜈ｨ譛滄俣蜈ｱ騾壹・1繧ｷ繝ｼ繝医↓菫晏ｭ倥☆繧九◆繧・縲よ里蟄倥・蜻ｼ縺ｳ蜃ｺ縺礼ｮ・園繧・    螟画峩縺帙★縺ｫ貂医・繧医≧縲∽ｺ呈鋤諤ｧ縺ｮ縺溘ａ縺ｫ蠑墓焚縺縺第ｮ九＠縺ｦ縺ゅｋ縲・    """
    sheets_store.append_kyuka_request(staff_name, date, request_type)


def load_kyuka_log() -> pd.DataFrame:
    """繝ｭ繧ｰ縺ｮ蜈ｨ陦・螻･豁ｴ繝ｻ蜿匁ｶ郁｡後ｒ蜷ｫ繧)繧偵◎縺ｮ縺ｾ縺ｾ隱ｭ縺ｿ霎ｼ繧縲・""
    return sheets_store.load_kyuka_log()


def compute_current_requests_from_log(log_df: pd.DataFrame, staff_df: pd.DataFrame) -> pd.DataFrame:
    """霑ｽ險伜梛繝ｭ繧ｰ縺九ｉ縲∫樟蝨ｨ譛牙柑縺ｪ蟶梧悍莨代・譛我ｼ代・荳隕ｧ(requests_df蠖｢蠑・繧貞・讒狗ｯ峨☆繧九・
    蜷御ｸ縺ｮ(繧ｹ繧ｿ繝・ヵ蜷・ 譌･莉・縺ｫ繝ｭ繧ｰ陦後′隍・焚縺ゅｋ蝣ｴ蜷医・縲√Ο繧ｰ縺ｮ霑ｽ險倬・〒譛蠕後・陦・    (=譛譁ｰ縺ｮ逕ｳ隲句・螳ｹ)縺ｮ縺ｿ繧呈治逕ｨ縺吶ｋ縲よ怙譁ｰ陦後′縲悟叙豸医阪□縺｣縺溷ｴ蜷医・縲√◎縺ｮ
    邨・∩蜷医ｏ縺帙ｒ邨先棡縺九ｉ髯､螟悶☆繧・隲也炊蜑企勁)縲・    """
    if log_df.empty:
        return default_requests_df()

    df = log_df.dropna(subset=["staff_name", "date"]).copy()
    if df.empty:
        return default_requests_df()
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    df = df.dropna(subset=["date"])
    if df.empty:
        return default_requests_df()
    # 驕主悉縺ｫ譌ｧ陦ｨ險・譛我ｼ醍筏隲・譛我ｼ醍｢ｺ螳・縺ｧ霑ｽ險倥＆繧後◆繝ｭ繧ｰ陦後ｂ縲∵眠陦ｨ險・譛臥ｵｦ)縺ｸ
    # 豁｣隕丞喧縺励※縺九ｉ謗｡逕ｨ縺吶ｋ(陦ｨ險伜､画峩縺縺代〒繝上・繝牙宛邏・°繧牙､悶ｌ縺ｪ縺・ｈ縺・↓縺吶ｋ)縲・    df["request_type"] = df["request_type"].apply(lambda k: _normalize_kind_value(k) if isinstance(k, str) else k)

    # groupby(...).last() 縺ｯ繧ｰ繝ｫ繝ｼ繝怜・縺ｮ譛邨ょ・迴ｾ陦・=繝ｭ繧ｰ荳翫〒譛繧よ眠縺励＞逕ｳ隲・繧呈治逕ｨ縺吶ｋ縲・    latest = df.groupby(["staff_name", "date"], as_index=False, sort=False).last()
    latest = latest[latest["request_type"] != CANCELLED_REQUEST_TYPE]
    if latest.empty:
        return default_requests_df()

    name_to_id = dict(zip(staff_df["name"], staff_df["staff_id"]))
    latest = latest.copy()
    latest["staff_id"] = latest["staff_name"].map(name_to_id)
    latest = latest.dropna(subset=["staff_id"])
    latest = latest.rename(columns={"staff_name": "name", "request_type": "kind"})
    return latest.reindex(columns=["staff_id", "name", "date", "kind"]).reset_index(drop=True)


def load_current_requests(year: int, month: int, staff_df: pd.DataFrame) -> pd.DataFrame:
    """謖・ｮ壽怦蠎ｦ縺ｮ縲檎樟蝨ｨ譛牙柑縺ｪ縲榊ｸ梧悍莨代・譛我ｼ台ｸ隕ｧ繧偵；oogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝井ｸ翫・

    霑ｽ險伜梛繝ｭ繧ｰ(蜈ｨ譛滄俣蜈ｱ騾壹・1繧ｷ繝ｼ繝・縺九ｉ蜀肴ｧ狗ｯ峨＠縺ｦ霑斐☆縲ゅΟ繧ｰ縺ｯ蜈ｨ譛滄俣蛻・′
    1縺､縺ｮ繧ｷ繝ｼ繝医↓縺ｾ縺ｨ縺ｾ縺｣縺ｦ縺・ｋ縺溘ａ縲・寔險亥ｾ後↓謖・ｮ壽怦蠎ｦ縺ｮ譌･莉倥□縺代∈邨槭ｊ霎ｼ繧縲・    """
    log_df = load_kyuka_log()
    current_df = compute_current_requests_from_log(log_df, staff_df)
    if current_df.empty:
        return current_df
    period_dates = set(get_period_dates(year, month))
    return current_df[current_df["date"].isin(period_dates)].reset_index(drop=True)


def sync_admin_requests_edit(current_df: pd.DataFrame, edited_df: pd.DataFrame, path: str | None = None) -> None:
    """邂｡逅・・↓繧医ｋ繝槭ヨ繝ｪ繧ｯ繧ｹ陦ｨ縺ｮ荳諡ｬ邱ｨ髮・ｒ縲∬ｿｽ險伜梛繝ｭ繧ｰ縺ｸ縺ｮ蟾ｮ蛻・ｿｽ險倥↓螟画鋤縺励※菫晏ｭ倥☆繧九・
    `current_df`(邱ｨ髮・燕縺ｮ髮・ｨ育憾諷・縺ｨ`edited_df`(邱ｨ髮・ｾ後・迥ｶ諷・繧呈ｯ碑ｼ・＠縲・    螳滄圀縺ｫ蛟､縺悟､牙喧縺励◆(繧ｹ繧ｿ繝・ヵ蜷・ 譌･莉・縺ｮ邨・∩蜷医ｏ縺帙↓縺､縺・※縺ｮ縺ｿ譁ｰ縺励＞繝ｭ繧ｰ陦・    (霑ｽ蜉/螟画峩縺ｯ譁ｰ遞ｮ蛻･縲∝炎髯､縺ｯ縲悟叙豸医・繧定ｿｽ險倥☆繧九ょ､牙喧縺ｮ縺ｪ縺・筏隲九↓縺ｯ荳蛻・    隗ｦ繧後↑縺・◆繧√√％縺ｮ邱ｨ髮・→蜷梧凾縺ｫ莉悶・繧ｹ繧ｿ繝・ヵ縺碁∽ｿ｡縺励◆蛟句挨逕ｳ隲九ｒ蟾ｻ縺崎ｾｼ繧薙〒
    豸医＠縺ｦ縺励∪縺・％縺ｨ縺後↑縺・・
    驥崎ｦ・ `current_df` 縺ｫ縺ｯ縲∫ｮ｡逅・・′螳滄圀縺ｫ邱ｨ髮・ｒ蟋九ａ縺滓凾轤ｹ縺ｮ繧ｹ繝翫ャ繝励す繝ｧ繝・ヨ
    (=繝槭ヨ繝ｪ繧ｯ繧ｹ陦ｨ繧呈緒逕ｻ縺励◆髫帙↓菴ｿ縺｣縺溽憾諷・繧呈ｸ｡縺吶％縺ｨ縲ゅ％縺ｮ髢｢謨ｰ繧貞他縺ｶ逶ｴ蜑阪↓
    謾ｹ繧√※繝・ぅ繧ｹ繧ｯ縺九ｉ譛譁ｰ迥ｶ諷九ｒ隱ｭ縺ｿ逶ｴ縺励※貂｡縺励※縺ｯ縺ｪ繧峨↑縺・りｪｭ縺ｿ逶ｴ縺励※縺励∪縺・→縲・    邂｡逅・・′邱ｨ髮・＠縺ｦ縺・ｋ髢薙↓莉悶・繧ｹ繧ｿ繝・ヵ縺碁∽ｿ｡縺励◆譁ｰ隕冗筏隲九′縲檎ｷｨ髮・燕蠕後〒豸医∴縺・    蟾ｮ蛻・阪→隱､隱崎ｭ倥＆繧後∝叙豸医→縺励※荳頑嶌縺阪＆繧後※縺励∪縺・・    """
    cur_map = {(r["name"], r["date"]): r["kind"] for _, r in current_df.iterrows()} if not current_df.empty else {}
    new_map = {(r["name"], r["date"]): r["kind"] for _, r in edited_df.iterrows()} if not edited_df.empty else {}

    for key in set(cur_map) | set(new_map):
        old_kind = cur_map.get(key)
        new_kind = new_map.get(key)
        if old_kind == new_kind:
            continue
        name, date = key
        append_kyuka_request(name, date, new_kind if new_kind is not None else CANCELLED_REQUEST_TYPE, path)


def clear_kyuka_log(current_df: pd.DataFrame, _path: str | None = None) -> None:
    """謖・ｮ壽怦蠎ｦ縺ｮ縲檎樟蝨ｨ譛牙柑縺ｪ縲榊ｸ梧悍莨代・譛我ｼ醍筏隲九ｒ縺吶∋縺ｦ蜿匁ｶ医↓縺吶ｋ(繝ｪ繧ｻ繝・ヨ讖溯・逕ｨ)縲・
    繝ｭ繧ｰ閾ｪ菴薙・霑ｽ險伜梛縺ｧGoogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝井ｸ翫↓縺吶∋縺ｦ縺ｮ螻･豁ｴ繧剃ｿ晄戟縺励※縺・ｋ縺溘ａ縲・    縺薙％縺ｧ縺ｯ`current_df`(繝ｪ繧ｻ繝・ヨ蟇ｾ雎｡譛滄俣縺ｮ迴ｾ蝨ｨ譛牙柑縺ｪ逕ｳ隲倶ｸ隕ｧ)縺ｫ蜷ｫ縺ｾ繧後ｋ
    蜷・繧ｹ繧ｿ繝・ヵ蜷・ 譌･莉・縺ｫ縺､縺・※縲悟叙豸医阪Ο繧ｰ陦後ｒ霑ｽ險倥☆繧句ｽ｢縺ｧ螳溽樟縺吶ｋ
    (驕主悉縺ｮ螻･豁ｴ縺ｯ豸医＆縺壹∫屮譟ｻ繝ｭ繧ｰ縺ｨ縺励※谿九☆)縲・
    `_path` 縺ｯ譌ｧ繝ｭ繝ｼ繧ｫ繝ｫ繝輔ぃ繧､繝ｫ譁ｹ蠑上・蜷肴ｮ九・蠑墓焚縺ｧ縲ヾheets譁ｹ蠑上〒縺ｯ菴ｿ逕ｨ縺励↑縺・・    """
    if current_df is None or current_df.empty:
        return
    active_requests = [(r["name"], r["date"]) for _, r in current_df.iterrows()]
    sheets_store.clear_kyuka_requests_in_range(active_requests)


def kyuka_matrix_date_labels(dates_df: pd.DataFrame) -> list[str]:
    """蟶梧悍莨代・繝医Μ繧ｯ繧ｹ陦ｨ縺ｮ蛻苓ｦ句・縺・蟇ｾ雎｡譛滄俣縺ｮ蜈ｨ譌･莉・繧定ｿ斐☆縲・""
    return [f"{d.month}/{d.day}({wd})" for d, wd in zip(dates_df["date"], dates_df["weekday_jp"])]


def build_kyuka_requests_wide(
    requests_df: pd.DataFrame,
    staff_df: pd.DataFrame,
    dates_df: pd.DataFrame,
) -> pd.DataFrame:
    """迴ｾ蝨ｨ縺ｮ逕ｳ隲倶ｸ隕ｧ(髟ｷ蠖｢蠑・繧偵∫ｮ｡逅・・畑繝槭ヨ繝ｪ繧ｯ繧ｹ陦ｨ(繧ｹ繧ｿ繝・ヵﾃ玲律莉・縺ｫ螟画鋤縺吶ｋ縲・
    譛ｪ逕ｳ隲九・繧ｻ繝ｫ縺ｯ蠎苓・蛻･譌･蛻･繧ｷ繝輔ヨ陦ｨ縺ｮ謇句虚邱ｨ髮・→蜷後§陦ｨ險・BLANK_LABEL=縲鯉ｼ育ｩｺ逋ｽ・峨・
    縺ｧ蝓九ａ繧・陦ｨ遉ｺ繝ｻ邱ｨ髮・・繝ｫ繝繧ｦ繝ｳ縺ｮ驕ｸ謚櫁い繧貞・遉ｾ逧・↓邨ｱ荳縺吶ｋ縺溘ａ)縲・    """
    date_labels = kyuka_matrix_date_labels(dates_df)
    lookup: dict[tuple[str, dt.date], str] = {}
    if not requests_df.empty:
        for _, r in requests_df.iterrows():
            lookup[(r["name"], r["date"])] = r["kind"]

    wide = pd.DataFrame({"繧ｹ繧ｿ繝・ヵ蜷・: staff_df["name"].tolist()})
    for label, d in zip(date_labels, dates_df["date"]):
        wide[label] = [lookup.get((nm, d), BLANK_LABEL) for nm in staff_df["name"]]
    return wide


def kyuka_requests_wide_to_long(
    wide_df: pd.DataFrame,
    staff_df: pd.DataFrame,
    dates_df: pd.DataFrame,
) -> pd.DataFrame:
    """邂｡逅・・畑繝槭ヨ繝ｪ繧ｯ繧ｹ陦ｨ(邱ｨ髮・ｾ・繧偵∫筏隲倶ｸ隕ｧ(髟ｷ蠖｢蠑・縺ｸ螟画鋤縺吶ｋ縲・
    遨ｺ谺・・譛ｪ遏･縺ｮ遞ｮ蛻･蛟､縺ｯ辟｡隕悶☆繧・繧ｯ繝ｩ繝・す繝･縺励↑縺・縲・    邂｡逅・・ｷｨ髮・畑繧ｰ繝ｪ繝・ラ縺瑚牡莉倥″邨ｵ譁・ｭ嶺ｻ倥″縺ｮ陦ｨ險・KYUKA_KIND_EMOJI_LABELS)縺ｧ
    貂｡縺輔ｌ縺溷ｴ蜷医ｂ縲√％縺薙〒繝励Ξ繝ｼ繝ｳ縺ｪ遞ｮ蛻･譁・ｭ怜・(縲悟ｸ梧悍莨代咲ｭ・縺ｸ豁｣隕丞喧縺励※縺九ｉ
    蛻､螳壹・菫晏ｭ倥☆繧・邨ｵ譁・ｭ苓｡ｨ險倥・縺ゅ￥縺ｾ縺ｧ邱ｨ髮・判髱｢縺ｮ陦ｨ遉ｺ荳翫・蟾･螟ｫ縺ｧ縺ゅｊ縲∽ｿ晏ｭ倥ョ繝ｼ繧ｿ
    縺ｫ縺ｯ蜿肴丐縺輔○縺ｪ縺・◆繧・縲・    """
    date_labels = kyuka_matrix_date_labels(dates_df)
    label_to_date = dict(zip(date_labels, dates_df["date"]))
    name_to_id = dict(zip(staff_df["name"], staff_df["staff_id"]))

    rows = []
    for _, r in wide_df.iterrows():
        name = r.get("繧ｹ繧ｿ繝・ヵ蜷・)
        if name not in name_to_id:
            continue
        for label in date_labels:
            val = str(r.get(label, "")).strip()
            val = KYUKA_EMOJI_TO_KIND.get(val, val)
            if val not in REQUEST_KINDS:
                continue
            rows.append({"staff_id": name_to_id[name], "name": name, "date": label_to_date[label], "kind": val})
    return pd.DataFrame(rows, columns=["staff_id", "name", "date", "kind"])


# ---------------------------------------------------------------------------
# 譛驕ｩ蛹也ｵ先棡繝ｻ謇句虚邱ｨ髮・す繝輔ヨ縺ｮ豌ｸ邯壼喧(Google繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝・
# ---------------------------------------------------------------------------
# 莉･蜑阪・縺薙％縺ｫ繝ｭ繝ｼ繧ｫ繝ｫCSV/JSON繝輔ぃ繧､繝ｫ縺ｸ菫晏ｭ倥☆繧句ｮ溯｣・′縺ゅ▲縺溘′縲ヾtreamlit
# Community Cloud縺ｯ繧｢繝励Μ縺ｮ蜀崎ｵｷ蜍輔・蜀阪ョ繝励Ο繧､縺ｮ縺溘・縺ｫ繝ｭ繝ｼ繧ｫ繝ｫ繝輔ぃ繧､繝ｫ繧ｷ繧ｹ繝・Β
# 縺ｸ縺ｮ譖ｸ縺崎ｾｼ縺ｿ繧貞・縺ｦ豸亥悉縺励※縺励∪縺・◆繧√；oogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医∈菫晏ｭ倥☆繧区婿蠑・# (sheets_store.py)縺ｫ鄂ｮ縺肴鋤縺医◆縲る未謨ｰ蜷阪・繧ｷ繧ｰ繝阪メ繝｣縺ｯ蜿ｯ閭ｽ縺ｪ髯舌ｊ邯ｭ謖√＠縺ｦ縺・ｋ縲・

def save_shift_result_to_disk(shift_df: pd.DataFrame, meta: dict, *args, **kwargs) -> None:
    """譛驕ｩ蛹也ｵ先棡(縺ｾ縺溘・謇句虚邱ｨ髮・ｾ後・譛譁ｰ繧ｷ繝輔ヨ)繧竪oogle繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医∈蜊ｳ譎ゆｿ晏ｭ倥☆繧九・
    譛驕ｩ蛹悶・螳溯｡悟ｮ御ｺ・凾繝ｻ謇句虚邱ｨ髮・・遒ｺ螳壽凾繝ｻ繝ｪ繧ｻ繝・ヨ譎ゅ・縺・★繧後°繧峨ｂ蜻ｼ縺ｳ蜃ｺ縺吶％縺ｨ縺ｧ縲・    繝悶Λ繧ｦ繧ｶ繧帝哩縺倥◆繧雁挨繝悶Λ繧ｦ繧ｶ/蛻･遶ｯ譛ｫ縺九ｉ繧｢繧ｯ繧ｻ繧ｹ縺励◆蝣ｴ蜷医〒繧ょｾｩ蜈・〒縺阪ｋ縲・    """
    try:
        sheets_store.save_shift_result(shift_df, meta)
    except Exception:
        # 菫晏ｭ伜､ｱ謨玲凾繧ゅい繝励Μ蜈ｨ菴薙・繧ｯ繝ｩ繝・す繝･縺輔○縺ｪ縺・逕ｻ髱｢荳翫・迥ｶ諷九・菫晄戟縺輔ｌ繧・縲・        pass


def load_shift_result_from_disk(*args, **kwargs) -> tuple[pd.DataFrame | None, dict]:
    """Google繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医↓菫晏ｭ俶ｸ医∩縺ｮ譛譁ｰ繧ｷ繝輔ヨ邨先棡縺ｨ繝｡繧ｿ諠・ｱ繧定ｪｭ縺ｿ霎ｼ繧縲・
    菫晏ｭ倥ョ繝ｼ繧ｿ縺檎┌縺・謗･邯壹お繝ｩ繝ｼ縺ｮ蝣ｴ蜷医・ (None, {}) 繧定ｿ斐☆(繧ｯ繝ｩ繝・す繝･縺励↑縺・縲・    """
    try:
        return sheets_store.load_shift_result()
    except Exception:
        return None, {}


def clear_saved_shift_result(*args, **kwargs) -> None:
    """Google繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝医↓菫晏ｭ俶ｸ医∩縺ｮ譛譁ｰ繧ｷ繝輔ヨ邨先棡繧貞炎髯､縺吶ｋ縲・""
    try:
        sheets_store.clear_shift_result()
    except Exception:
        pass


def build_simple_staff_summary(
    shift_df: pd.DataFrame,
    staff_df: pd.DataFrame,
    dates_df: pd.DataFrame,
) -> pd.DataFrame:
    """繧ｷ繝輔ヨ(髟ｷ蠖｢蠑・縺九ｉ縲∫ｰ｡譏鍋沿縺ｮ繧ｹ繧ｿ繝・ヵ蛻･繧ｵ繝槭Μ繝ｼ(蜃ｺ蜍､譌･謨ｰ繝ｻ蝨滓律蜃ｺ蜍､繝ｻ繝倥Ν繝・繧剃ｽ懊ｋ縲・
    繝ｭ繝ｼ繧ｫ繝ｫ菫晏ｭ倥ョ繝ｼ繧ｿ縺九ｉ蠕ｩ蜈・＠縺溘そ繝・す繝ｧ繝ｳ縺ｪ縺ｩ縲∝・縺ｮ繧ｽ繝ｫ繝舌・蜀・Κ螟画焚(蝌ｱ險・繝代・繝医・
    遞ｼ蜒肴律謨ｰ繝ｬ繝ｳ繧ｸ蛻､螳壹↑縺ｩ)縺檎┌縺・ｴ蜷医・邁｡譏楢｡ｨ遉ｺ逕ｨ縲・    """
    rows = []
    for row in staff_df.itertuples():
        sid = row.staff_id
        work_days = shift_df[shift_df["staff_id"] == sid] if not shift_df.empty else pd.DataFrame()
        weekend_cnt = int(work_days["date"].apply(lambda d: d.weekday() in (5, 6)).sum()) if not work_days.empty else 0
        help_cnt = 0
        if row.emp_type in ("蠎鈴聞", "豁｣遉ｾ蜩｡", "蝌ｱ險・) and not work_days.empty:
            help_cnt = int((work_days["store"] != row.home_store).sum())
        rows.append(
            {
                "staff_id": sid,
                "name": row.name,
                "蛹ｺ蛻・: row.emp_type,
                "蜃ｺ蜍､譌･謨ｰ": len(work_days),
                "蝨滓律蜃ｺ蜍､譌･謨ｰ": weekend_cnt,
                "繝倥Ν繝怜・蜍､譌･謨ｰ": help_cnt,
                "蜍､蜍呎律謨ｰ繝ｬ繝ｳ繧ｸ(蝌ｱ險・繝代・繝・": "-",
            }
        )
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 蟶梧悍莨代・譛我ｼ・CSV 荳諡ｬ繧､繝ｳ繝昴・繝茨ｼ上ユ繝ｳ繝励Ξ繝ｼ繝亥・蜉・# ---------------------------------------------------------------------------

_CSV_KIND_LOOKUP = {
    "蟶梧悍莨・: "蟶梧悍莨・,
    "莨・: "蟶梧悍莨・,
    "蜈ｬ莨・: "蟶梧悍莨・,
    "譛我ｼ・: "譛臥ｵｦ逕ｳ隲・,
    "譛臥ｵｦ": "譛臥ｵｦ逕ｳ隲・,
    "譛臥ｵｦ逕ｳ隲・: "譛臥ｵｦ逕ｳ隲・,
    "譛我ｼ醍筏隲・: "譛臥ｵｦ逕ｳ隲・,  # 譌ｧ陦ｨ險・蠕梧婿莠呈鋤)
    "譛臥ｵｦ遒ｺ螳・: "譛臥ｵｦ逕ｳ隲・,  # 蟒・ｭ｢蛹ｺ蛻・蠕梧婿莠呈鋤縲√梧怏邨ｦ逕ｳ隲九阪∈邨ｱ蜷・
    "譛我ｼ醍｢ｺ螳・: "譛臥ｵｦ逕ｳ隲・,  # 譌ｧ陦ｨ險倥・蟒・ｭ｢蛹ｺ蛻・蠕梧婿莠呈鋤縲√梧怏邨ｦ逕ｳ隲九阪∈邨ｱ蜷・
    "邨ｶ蟇ｾ莨・: "邨ｶ蟇ｾ莨・,
}


def _normalize_date_token(s: str) -> str:
    """CSV荳ｭ縺ｮ譌･莉倩｡ｨ險倥・謠ｺ繧・蜈ｨ隗呈焚蟄励・譖懈律繧ｫ繝・さ繝ｻ蟷ｴ譛域律陦ｨ險倡ｭ・繧呈ｭ｣隕丞喧縺吶ｋ縲・""
    s = unicodedata.normalize("NFKC", str(s)).strip()
    s = re.sub(r"\(.*?\)", "", s)  # "(豌ｴ)" 遲峨・譖懈律繧ｫ繝・さ繧帝勁蜴ｻ
    s = re.sub(r"・・*?・・, "", s)
    s = s.replace("蟷ｴ", "/").replace("譛・, "/")
    s = s.replace("譌･", "")
    return s.strip()


def _date_label_variants(d: dt.date) -> set[str]:
    """1縺､縺ｮ譌･莉倥′蜿悶ｊ縺・ｋ莉｣陦ｨ逧・↑譁・ｭ怜・陦ｨ險倥・繝舌Μ繧ｨ繝ｼ繧ｷ繝ｧ繝ｳ繧貞・謖吶☆繧九・""
    variants = set()
    for m in (str(d.month), f"{d.month:02d}"):
        for day in (str(d.day), f"{d.day:02d}"):
            variants.add(f"{m}/{day}")
            variants.add(f"{m}-{day}")
    variants.add(d.isoformat())
    variants.add(f"{d.year}/{d.month}/{d.day}")
    variants.add(f"{d.year}/{d.month:02d}/{d.day:02d}")
    variants.add(f"{d.year}-{d.month:02d}-{d.day:02d}")
    return variants


def _build_date_lookup(dates_df: pd.DataFrame) -> dict[str, dt.date]:
    lookup: dict[str, dt.date] = {}
    for d in dates_df["date"]:
        for v in _date_label_variants(d):
            lookup[v] = d
    return lookup


def generate_requests_csv_template(staff_df: pd.DataFrame, dates_df: pd.DataFrame) -> bytes:
    """蜈ｨ繧ｹ繧ｿ繝・ヵﾃ怜ｯｾ雎｡譛滄俣蜈ｨ譌･莉倥・縲∫ｩｺ縺ｮ蟶梧悍莨大・蜉帷畑CSV繝・Φ繝励Ξ繝ｼ繝医ｒ逕滓・縺吶ｋ縲・
    譌･莉倥・繝医Μ繧ｯ繧ｹ蠖｢蠑・1蛻礼岼=繧ｹ繧ｿ繝・ヵ蜷阪・蛻礼岼莉･髯・蜷・律莉・縲ゅそ繝ｫ縺ｯ遨ｺ谺・・縺ｾ縺ｾ
    驟榊ｸ・＠縲∝茜逕ｨ閠・↓縲悟ｸ梧悍莨代阪梧怏莨代阪檎ｵｶ蟇ｾ莨代阪・縺・★繧後°繧定ｨ伜・縺励※繧ゅｉ縺・・    """
    date_labels = [f"{d.month}/{d.day}({wd})" for d, wd in zip(dates_df["date"], dates_df["weekday_jp"])]
    template_df = pd.DataFrame({"繧ｹ繧ｿ繝・ヵ蜷・: staff_df["name"].tolist()})
    for label in date_labels:
        template_df[label] = ""
    return template_df.to_csv(index=False).encode("utf-8-sig")


def parse_requests_csv(
    file_bytes: bytes,
    staff_df: pd.DataFrame,
    dates_df: pd.DataFrame,
) -> tuple[pd.DataFrame, list[str]]:
    """蟶梧悍莨代・譛我ｼ舛SV繧定ｪｭ縺ｿ霎ｼ縺ｿ縲〉equests_df蠖｢蠑・staff_id/name/date/kind)縺ｫ螟画鋤縺吶ｋ縲・
    莉･荳九・2蠖｢蠑上ｒ閾ｪ蜍募愛蛻･縺励※蜿励￠莉倥￠繧・繧ｯ繝ｩ繝・す繝･縺帙★縲∬ｪ崎ｭ倥〒縺阪↑縺・｡・蛻・蛟､縺ｯ
    隴ｦ蜻翫Γ繝・そ繝ｼ繧ｸ縺ｫ髮・ｴ・＠縺ｦ繧ｹ繧ｭ繝・・縺吶ｋ):
      - 邵ｦ謖√■3蛻怜ｽ｢蠑・ 繧ｹ繧ｿ繝・ヵ蜷・ 譌･莉・ 遞ｮ蛻･(蟶梧悍莨・譛我ｼ・邨ｶ蟇ｾ莨・遲・
      - 譌･莉倥・繝医Μ繧ｯ繧ｹ蠖｢蠑・ 1蛻礼岼=繧ｹ繧ｿ繝・ヵ蜷阪・蛻礼岼莉･髯・蜷・律莉倥√そ繝ｫ=遞ｮ蛻･
    """
    warnings: list[str] = []
    empty = pd.DataFrame(columns=["staff_id", "name", "date", "kind"])

    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            raw = pd.read_csv(io.BytesIO(file_bytes), dtype=str, keep_default_na=False, encoding=encoding)
            break
        except UnicodeDecodeError:
            raw = None
            continue
        except Exception as e:
            return empty, [f"CSV縺ｮ隱ｭ縺ｿ霎ｼ縺ｿ縺ｫ螟ｱ謨励＠縺ｾ縺励◆: {e}"]
    else:
        return empty, ["CSV縺ｮ譁・ｭ励さ繝ｼ繝峨ｒ隱崎ｭ倥〒縺阪∪縺帙ｓ縺ｧ縺励◆(UTF-8縺ｾ縺溘・Shift-JIS縺ｧ菫晏ｭ倥＠縺ｦ縺上□縺輔＞)縲・]

    if raw is None or raw.shape[1] < 2:
        return empty, ["CSV縺ｫ蜊∝・縺ｪ蛻励′縺ゅｊ縺ｾ縺帙ｓ(譛菴・蛻怜ｿ・ｦ√〒縺・縲・]

    name_to_id = dict(zip(staff_df["name"], staff_df["staff_id"]))
    date_lookup = _build_date_lookup(dates_df)
    rows: list[dict] = []
    col_names = list(raw.columns)
    name_col = col_names[0]

    if raw.shape[1] <= 3:
        # 邵ｦ謖√■3蛻怜ｽ｢蠑・ 繧ｹ繧ｿ繝・ヵ蜷・ 譌･莉・ 遞ｮ蛻･
        date_col = col_names[1] if len(col_names) > 1 else None
        kind_col = col_names[2] if len(col_names) > 2 else None
        for i, r in raw.iterrows():
            name = str(r[name_col]).strip()
            date_raw = str(r[date_col]).strip() if date_col else ""
            kind_raw = str(r[kind_col]).strip() if kind_col else ""
            if not name and not date_raw and not kind_raw:
                continue
            if name not in name_to_id:
                warnings.append(f"{i + 2}陦檎岼: 繧ｹ繧ｿ繝・ヵ蜷阪鶏name}縲阪′隕九▽縺九ｊ縺ｾ縺帙ｓ縲ゅ％縺ｮ陦後・繧ｹ繧ｭ繝・・縺励∪縺励◆縲・)
                continue
            d = date_lookup.get(_normalize_date_token(date_raw))
            if d is None:
                warnings.append(f"{i + 2}陦檎岼: 譌･莉倥鶏date_raw}縲阪ｒ隱崎ｭ倥〒縺阪∪縺帙ｓ縲ゅ％縺ｮ陦後・繧ｹ繧ｭ繝・・縺励∪縺励◆縲・)
                continue
            kind = _CSV_KIND_LOOKUP.get(kind_raw.strip())
            if kind is None:
                warnings.append(
                    f"{i + 2}陦檎岼: 遞ｮ蛻･縲鶏kind_raw}縲阪ｒ隱崎ｭ倥〒縺阪∪縺帙ｓ"
                    "(蟶梧悍莨・譛我ｼ・邨ｶ蟇ｾ莨代・縺・★繧後°繧呈欠螳壹＠縺ｦ縺上□縺輔＞)縲ゅ％縺ｮ陦後・繧ｹ繧ｭ繝・・縺励∪縺励◆縲・
                )
                continue
            rows.append({"staff_id": name_to_id[name], "name": name, "date": d, "kind": kind})
    else:
        # 譌･莉倥・繝医Μ繧ｯ繧ｹ蠖｢蠑・ 1蛻礼岼=繧ｹ繧ｿ繝・ヵ蜷阪・蛻礼岼莉･髯・蜷・律莉伜・
        resolved_date_cols: dict[str, dt.date] = {}
        for col in col_names[1:]:
            d = date_lookup.get(_normalize_date_token(str(col)))
            if d is None:
                warnings.append(f"蛻苓ｦ句・縺励鶏col}縲阪ｒ譌･莉倥→縺励※隱崎ｭ倥〒縺阪∪縺帙ｓ縲ゅ％縺ｮ蛻励・繧ｹ繧ｭ繝・・縺励∪縺励◆縲・)
                continue
            resolved_date_cols[col] = d

        for i, r in raw.iterrows():
            name = str(r[name_col]).strip()
            if not name:
                continue
            if name not in name_to_id:
                warnings.append(f"{i + 2}陦檎岼: 繧ｹ繧ｿ繝・ヵ蜷阪鶏name}縲阪′隕九▽縺九ｊ縺ｾ縺帙ｓ縲ゅ％縺ｮ陦後・繧ｹ繧ｭ繝・・縺励∪縺励◆縲・)
                continue
            for col, d in resolved_date_cols.items():
                val = str(r[col]).strip()
                if not val or val in ("0", "-", "・・, "nan", "NaN"):
                    continue
                kind = _CSV_KIND_LOOKUP.get(val)
                if kind is None:
                    warnings.append(
                        f"{name} / {d.month}/{d.day}: 蛟､縲鶏val}縲阪ｒ隱崎ｭ倥〒縺阪∪縺帙ｓ"
                        "(蟶梧悍莨・譛我ｼ・邨ｶ蟇ｾ莨代・縺・★繧後°繧呈欠螳壹＠縺ｦ縺上□縺輔＞)縲ゅ％縺ｮ繧ｻ繝ｫ縺ｯ繧ｹ繧ｭ繝・・縺励∪縺励◆縲・
                    )
                    continue
                rows.append({"staff_id": name_to_id[name], "name": name, "date": d, "kind": kind})

    return pd.DataFrame(rows, columns=["staff_id", "name", "date", "kind"]), warnings


# ---------------------------------------------------------------------------
# 譛臥ｵｦ莨第嚊縲悟叙蠕怜庄閭ｽ譌･繝ｻ莠ｺ謨ｰ譫縲阪・閾ｪ蜍慕ｮ怜・
# ---------------------------------------------------------------------------

def min_required_employee_bodies_per_day() -> int:
    """1蝟ｶ讌ｭ譌･縺ゅ◆繧翫∝・7蠎苓・繧堤ｨｼ蜒阪＆縺帙ｋ縺ｮ縺ｫ譛菴朱剞蠢・ｦ√↑縲檎､ｾ蜩｡(豁｣遉ｾ蜩｡/蝌ｱ險・縲堺ｺｺ謨ｰ縲・
    繝代・繝井ｽｵ逕ｨ蜿ｯ閭ｽ蠎苓・縺ｯ繝代・繝医ｒ菴ｿ縺・燕謠撰ｼ育､ｾ蜩｡1蜷・繝代・繝・蜷搾ｼ峨〒譛蟆丞喧縲・    蠕ｳ驥榊ｺ励・繝代ち繝ｼ繝ｳ竭｡・育､ｾ蜩｡1蜷・繝代・繝・+繝代・繝・・峨ｒ菴ｿ縺・燕謠舌〒譛蟆丞喧縺吶ｋ縲・    """
    general = len(GENERAL_STORES) * 2
    combo = len(COMBO_STORE_PART_ROLES) * 1
    tokushige = 1
    return general + combo + tokushige


def compute_paid_leave_availability(
    dates_df: pd.DataFrame,
    staff_df: pd.DataFrame,
    requests_df: pd.DataFrame,
) -> pd.DataFrame:
    """蝟ｶ讌ｭ譌･縺斐→縺ｫ縲梧怏邨ｦ蜿門ｾ怜庄閭ｽ莠ｺ謨ｰ譫・育､ｾ蜩｡繝ｻ蝌ｱ險怜髄縺托ｼ峨阪ｒ閾ｪ蜍慕ｮ怜・縺吶ｋ縲・
    閠・∴譁ｹ:
      縺昴・譌･縺ｮ縲檎､ｾ蜩｡(蠎鈴聞+豁｣遉ｾ蜩｡+蝌ｱ險・縲阪・鬆ｭ謨ｰ縺ｮ縺・■縲∵里縺ｫ邨ｶ蟇ｾ莨・譛臥ｵｦ逕ｳ隲九〒
      謚懊￠縺ｦ縺・ｋ莠ｺ謨ｰ繧帝勁縺・◆縲檎ｨｼ蜒榊庄閭ｽ莠ｺ謨ｰ縲阪°繧峨∝・蠎鈴°蝟ｶ縺ｫ譛菴朱剞蠢・ｦ√↑
      莠ｺ謨ｰ(min_required_employee_bodies_per_day)繧貞ｷｮ縺怜ｼ輔＞縺滉ｽ吝鴨繧偵・      縺昴・譌･縺ｫ譁ｰ隕上〒譛我ｼ代ｒ蜑ｲ繧雁ｽ薙※繧峨ｌ繧倶ｺｺ謨ｰ譫縺ｨ縺ｿ縺ｪ縺吶・    """
    employee_staff = staff_df[staff_df["emp_type"].isin(EMPLOYEE_TYPES)]
    total_employees = len(employee_staff)
    min_required = min_required_employee_bodies_per_day()

    hard_off_by_date: dict[dt.date, int] = {}
    if not requests_df.empty:
        hard = requests_df[requests_df["kind"].isin(HARD_OFF_KINDS)]
        hard = hard[hard["staff_id"].isin(employee_staff["staff_id"])]
        hard_off_by_date = hard.groupby("date")["staff_id"].nunique().to_dict()

    rows = []
    for _, day in dates_df[dates_df["is_business_day"]].iterrows():
        d = day["date"]
        already_off = hard_off_by_date.get(d, 0)
        available = total_employees - already_off
        surplus = max(0, available - min_required)
        rows.append(
            {
                "date": d,
                "weekday_jp": day["weekday_jp"],
                "day_type": day["day_type"],
                "employee_headcount": total_employees,
                "already_confirmed_off": already_off,
                "min_required_employees": min_required,
                "paid_leave_slots": surplus,
            }
        )
    return pd.DataFrame(rows)


def generate_paid_leave_announcement(
    availability_df: pd.DataFrame,
    period_label: str,
) -> str:
    """繧ｹ繧ｿ繝・ヵ蜷代￠縺ｮ譛臥ｵｦ莨第嚊譯亥・繝・く繧ｹ繝医ｒ繝ｯ繝ｳ繧ｯ繝ｪ繝・け逕滓・縺吶ｋ縲・""
    available_days = availability_df[availability_df["paid_leave_slots"] > 0]

    lines = [
        f"縲須period_label}縲題ｨ育判蟷ｴ莨代・縺疲｡亥・",
        "",
        "荳玖ｨ倥・譌･縺ｯ莠ｺ蜩｡縺ｫ菴吝鴨縺後≠繧九◆繧√∵怏邨ｦ莨第嚊縺ｮ蜿門ｾ励ｒ蜆ｪ蜈育噪縺ｫ蜿励￠莉倥￠縺ｾ縺吶・,
        "蜿門ｾ励ｒ縺泌ｸ梧悍縺ｮ譁ｹ縺ｯ縲∵球蠖薙∪縺ｧ縺顔筏縺怜・縺上□縺輔＞・亥・逹繝ｻ隱ｿ謨ｴ縺ｮ荳翫∵ｱｺ螳壹＠縺ｾ縺呻ｼ峨・,
        "",
    ]
    if available_days.empty:
        lines.append("窶ｻ莉頑悄髢謎ｸｭ縲∽ｽ吝鴨繧堤｢ｺ菫昴〒縺阪ｋ譌･縺後≠繧翫∪縺帙ｓ縺ｧ縺励◆縲ょ句挨縺ｫ縺皮嶌隲・￥縺縺輔＞縲・)
    else:
        for _, row in available_days.iterrows():
            d: dt.date = row["date"]
            lines.append(
                f"繝ｻ{d.month}譛・d.day}譌･・・row['weekday_jp']}・・
                f"縲蜿門ｾ怜庄閭ｽ譫・嘴int(row['paid_leave_slots'])}蜷・
            )
        lines.append("")
        lines.append(f"蜷郁ｨ・{len(available_days)} 譌･髢薙∝叙蠕怜庄閭ｽ譫縺ゅｊ縲・)
    lines.append("")
    lines.append("窶ｻ荳願ｨ倅ｻ･螟悶・譌･縺ｧ繧ゅ∽ｺｺ蜩｡隱ｿ謨ｴ縺悟庄閭ｽ縺ｪ蝣ｴ蜷医・蜿門ｾ励〒縺阪ｋ縺薙→縺後≠繧翫∪縺吶・)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 譛驕ｩ蛹也ｵ先棡縺ｫ蝓ｺ縺･縺上梧怏邨ｦ莨第嚊 霑ｽ蜉蜿門ｾ怜庄閭ｽ譫縲阪・蜍慕噪蛻､螳・# ---------------------------------------------------------------------------

def _min_required_employees_for_store(
    store: str, part_roles_present: set[str], is_weekend_holiday: bool = False
) -> int:
    """縺昴・蠎苓・繝ｻ縺昴・譌･縺ｮ螳滄圀縺ｮ繝代・繝亥・蜍､迥ｶ豕√ｒ蜑肴署縺ｫ縲∵・遶九↓譛菴朱剞蠢・ｦ√↑遉ｾ蜩｡謨ｰ繧定ｿ斐☆縲・
    optimizer.py 縺ｮ蠎苓・蛻･繝上・繝牙宛邏・→蜷後§蛻､螳壼渕貅悶↓蜷医ｏ縺帙※縺・ｋ:
      - 遞ｲ豐｢蠎励・譁ｰ陝ｹ豎溷ｺ励・讌ｵ讌ｽ蠎・ 蟶ｸ縺ｫ2蜷・繝代・繝井ｸ榊庄縺ｮ縺溘ａ縲∵屆譌･繧貞撫繧上↑縺・縲・      - 螟ｧ豐ｻ蠎・ 險ｱ蜿ｯ繝代・繝・蟆ｾ貔､/驥朱％)縺ｮ縺・★繧後°縺悟・蜍､縺励※縺・ｌ縺ｰ1蜷阪√↑縺代ｌ縺ｰ2蜷・        (譖懈律繧貞撫繧上★蜷後§蛻､螳壹る㍽驕楢・霄ｫ縺ｯ蟷ｳ譌･縺ｮ縺ｿ蜃ｺ蜍､縺吶ｋ縺溘ａ縲∝悄譌･逾昴↓
        蠖ｼ縺悟・蜍､縺励※縺・ｋ迥ｶ諷九・縺昴ｂ縺昴ｂ逋ｺ逕溘＠縺ｪ縺・縲・      - 蜷榊商螻倶ｸｭ蟾晏ｺ励・螟ｩ逋ｽ讀咲伐蠎・ 蟷ｳ譌･縺ｯ險ｱ蜿ｯ繝代・繝医′蜃ｺ蜍､縺励※縺・ｌ縺ｰ1蜷阪・        縺ｪ縺代ｌ縺ｰ2蜷阪ょ悄譌･逾昴・縲檎､ｾ蜩｡1蜷・繝代・繝医阪・2蜷堺ｽ灘宛縺檎ｦ∵ｭ｢縺輔ｌ縺ｦ縺・ｋ縺溘ａ縲・        繝代・繝医・蜃ｺ蜍､譛臥┌縺ｫ髢｢繧上ｉ縺壼ｸｸ縺ｫ2蜷阪・      - 蠕ｳ驥榊ｺ・ 蟷ｳ譌･縺ｯ繝代・繝・繝ｻ繝代・繝・荳｡譁ｹ縺悟・蜍､縺励※縺・ｌ縺ｰ1蜷阪√◎繧御ｻ･螟悶・2蜷阪・        蝨滓律逾昴・縲檎､ｾ蜩｡1蜷・繝代・繝・蜷阪阪・3蜷堺ｽ灘宛縺檎ｦ∵ｭ｢縺輔ｌ縺ｦ縺・ｋ縺溘ａ縲・        繝代・繝医・蜃ｺ蜍､譛臥┌縺ｫ髢｢繧上ｉ縺壼ｸｸ縺ｫ2蜷阪・    """
    if store == TOKUSHIGE_STORE:
        if is_weekend_holiday:
            return 2
        return 1 if {"B", "C"} <= part_roles_present else 2
    if store in ("蜷榊商螻倶ｸｭ蟾晏ｺ・, "螟ｩ逋ｽ讀咲伐蠎・):
        if is_weekend_holiday:
            return 2
        return 1 if part_roles_present & set(COMBO_STORE_PART_ROLES[store]) else 2
    if store == "螟ｧ豐ｻ蠎・:
        return 1 if part_roles_present & set(COMBO_STORE_PART_ROLES[store]) else 2
    return 2


_LEAVE_REMOVABLE_TYPES = ("蠎鈴聞", "豁｣遉ｾ蜩｡")  # 縺薙・繝ｭ繧ｸ繝・け縺ｧ縲御ｼ代∩縺ｫ蝗槭☆縲榊呵｣懊→縺ｪ繧句玄蛻・# 莉｣譖ｿ隕∝藤(繝倥Ν繝玲兜蜈･)縺ｨ縺ｪ繧翫≧繧句玄蛻・ょｺ鈴聞繝ｻ豁｣遉ｾ蜩｡縺ｯ蜈ｨ蜩｡21譌･蜃ｺ蜍､(譛臥ｵｦ逕ｳ隲玖・・20譌･)縺ｧ
# 蝗ｺ螳壹＆繧後※縺翫ｊ縲∽ｻ｣譖ｿ縺ｫ蝗槭☆縺ｨ縺昴・蛻・□縺大・蜍､譌･謨ｰ縺瑚ｦ丞ｮ壹ｒ雜・∴縺ｦ縺励∪縺・◆繧√・# 莉｣譖ｿ蛟呵｣懊↓縺ｯ邨ｶ蟇ｾ縺ｫ蜷ｫ繧√↑縺・蝌ｱ險励・繝代・繝医・縺ｿ縺悟ｯｾ雎｡)縲・_LEAVE_SUBSTITUTE_TYPES = ("蝌ｱ險・, "繝代・繝・)


def compute_post_solve_leave_availability(
    shift_df: pd.DataFrame,
    staff_df: pd.DataFrame,
    dates_df: pd.DataFrame,
    requests_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """譛驕ｩ蛹悶仙ｾ後代・螳滄圀縺ｮ繧ｷ繝輔ヨ驟咲ｽｮ縺九ｉ縲∝推蝟ｶ讌ｭ譌･縺ｮ譛臥ｵｦ霑ｽ蜉蜿門ｾ怜庄閭ｽ譫繧堤ｮ怜・縺吶ｋ縲・
    莉｣譖ｿ繧ｷ繝溘Η繝ｬ繝ｼ繧ｷ繝ｧ繝ｳ蝙九・繧｢繝ｫ繧ｴ繝ｪ繧ｺ繝: 蜷・ｺ苓・繝ｻ蜷・霧讌ｭ譌･縺ｧ莉･荳九ｒ蛻､螳壹☆繧九・      繧ｱ繝ｼ繧ｹ1(譌｢蟄倥・菴吝臆): 迴ｾ蝨ｨ縺ｮ蝨ｨ邀堺ｺｺ謨ｰ縺梧・遶区怙菴惹ｺｺ謨ｰ繧剃ｸ雁屓縺｣縺ｦ縺・ｋ蛻・・縲・        縺昴・縺ｾ縺ｾ霑ｽ蜉縺ｮ譛我ｼ第棧縺ｨ縺ｪ繧・蠎鈴聞繝ｻ豁｣遉ｾ蜩｡縺悟ｯｾ雎｡縲・撼繧ｹ繧ｭ繝ｫ菫晄怏閠・ｒ蜆ｪ蜈・縲・      繧ｱ繝ｼ繧ｹ2(莉｣譖ｿ縺ｫ繧医ｋ蜑ｵ蜃ｺ): 縺｡繧・≧縺ｩ譛菴惹ｺｺ謨ｰ縺ｧ雜ｳ繧翫※縺・ｋ蝣ｴ蜷医〒繧ゅ・        縲悟ｮ滄圀縺ｫ縺昴・譌･縺昴・蠎苓・縺ｸ蜃ｺ蜍､縺励※縺・ｋ繧ｹ繧ｿ繝・ヵ縲阪・荳ｭ縺九ｉ蠎鈴聞繝ｻ豁｣遉ｾ蜩｡繧・        1蜷阪御ｼ代∩縲阪↓鄂ｮ縺肴鋤縺医◆縺ｨ莉ｮ螳壹＠縲√◎縺ｮ譌･縺ｩ縺薙↓繧ょ・蜍､縺励※縺翫ｉ縺壹√°縺､
        蟶梧悍莨代・邨ｶ蟇ｾ莨代・譛臥ｵｦ逕ｳ隲九・縺・★繧後ｂ蜈･縺｣縺ｦ縺・↑縺・=縺昴ｂ縺昴ｂ
        莨代∩縺溘＞繧上￠縺ｧ縺ｯ縺ｪ縺・豁｣遉ｾ蜩｡/蝌ｱ險励・荳ｭ縺九ｉ
          - 縺薙・蠎苓・縺ｸ縺ｮ蜍､蜍吶′險ｱ蜿ｯ縺輔ｌ縺ｦ縺・ｋ(譖懈律髯仙ｮ壹・遖∵ｭ｢繝ｫ繝ｼ繝ｫ繧ょ性繧)
          - 蝌ｱ險励・蝣ｴ蜷医・譛磯俣荳企剞譌･謨ｰ縺ｫ縺ｾ縺菴呵｣輔′縺ゅｋ
          - 霑ｽ蜉縺励※繧よ怙螟ｧ6騾｣蜍､繧定ｶ・∴縺ｪ縺・          - 繧ｹ繧ｭ繝ｫ隕∽ｻｶ(謚懊￠縺滉ｺｺ縺後せ繧ｭ繝ｫ菫晄怏閠・↑繧我ｻ｣譖ｿ繧ゅせ繧ｭ繝ｫ菫晄怏閠・繧呈ｺ縺溘☆
        莉｣譖ｿ蛟呵｣懊′1蜷阪〒繧りｦ九▽縺九ｌ縺ｰ縲√◎縺ｮ蛻・ｒ譛我ｼ大叙蠕怜庄閭ｽ譫縺ｨ縺励※險井ｸ翫☆繧九・    縺・★繧後・蠎苓・繧ゅ∵里縺ｫ謌千ｫ区怙菴惹ｺｺ謨ｰ繧剃ｸ句屓縺｣縺ｦ縺・ｋ(=荳崎ｶｳ險ｱ螳ｹ譌･)蝣ｴ蜷医・蟇ｾ雎｡螟悶→縺励・    縺昴・譌･繝ｻ縺昴・蠎苓・縺ｮ蛻､螳壹・縺ｿ繧偵せ繧ｭ繝・・縺吶ｋ(莉悶・譌･繝ｻ莉悶・蠎苓・縺ｮ險育ｮ励↓縺ｯ荳蛻・ｽｱ髻ｿ縺励↑縺・縲・    莉｣譖ｿ蛟呵｣懊・蜷梧律蜀・〒驥崎､・＠縺ｦ菴ｿ縺・屓縺輔↑縺・・    莨代∩縺檎｢ｺ螳壹＠縺ｦ縺・ｋ(蟶梧悍莨代′騾壹▲縺溽ｭ・繧ｹ繧ｿ繝・ヵ縺ｯ縲∵怏莨大呵｣懊↓繧ゆｻ｣譖ｿ隕∝藤縺ｫ繧・    荳蛻・匳蝣ｴ縺励↑縺・螳滓・縺ｨ遏帷崟縺吶ｋ陦ｨ遉ｺ繧帝亟縺・縲・    """
    business_days = dates_df.loc[dates_df["is_business_day"]]

    if shift_df.empty:
        return pd.DataFrame(
            [
                {
                    "date": day["date"],
                    "weekday_jp": day["weekday_jp"],
                    "day_type": day["day_type"],
                    "extra_leave_slots": 0,
                    "detail": "",
                }
                for _, day in business_days.iterrows()
            ]
        )

    staff_emp_type = dict(zip(staff_df["staff_id"], staff_df["emp_type"]))
    staff_has_skill = dict(zip(staff_df["staff_id"], staff_df["has_skill"]))
    staff_home = dict(zip(staff_df["staff_id"], staff_df["home_store"]))
    staff_part_role = dict(zip(staff_df["staff_id"], staff_df["part_role"]))
    staff_name = dict(zip(staff_df["staff_id"], staff_df["name"]))
    staff_max_workdays = dict(zip(staff_df["staff_id"], staff_df.get("max_workdays", pd.Series(dtype=float))))
    staff_allowed = {row.staff_id: set(derive_allowed_stores(row)) for row in staff_df.itertuples()}
    staff_weekend_holiday_forbidden = dict(zip(staff_df["staff_id"], staff_df.get("weekend_holiday_forbidden_stores")))
    staff_sat_sun_forbidden = dict(zip(staff_df["staff_id"], staff_df.get("saturday_sunday_forbidden_stores")))
    name_to_id = dict(zip(staff_df["name"], staff_df["staff_id"]))

    # 蟶梧悍莨代・邨ｶ蟇ｾ莨代・譛臥ｵｦ逕ｳ隲九・縺・★繧後°縺悟・縺｣縺ｦ縺・ｋ(staff_id, date)縺ｮ髮・粋縲・    # 縺薙・縺・★繧後°縺ｫ隧ｲ蠖薙☆繧九せ繧ｿ繝・ヵ縺ｯ縲√◎繧ゅ◎繧ゆｼ代∩縺溘＞/莨代∩縺檎｢ｺ螳壹＠縺ｦ縺・ｋ縺溘ａ縲・    # 莉｣譖ｿ隕∝藤縺ｨ縺励※縲悟他縺ｳ蜃ｺ縺吶榊呵｣懊↓縺ｯ邨ｶ蟇ｾ縺ｫ縺励↑縺・・    requested_off_pairs: set[tuple[str, dt.date]] = set()
    if requests_df is not None and not requests_df.empty:
        requested_off_pairs = set(zip(requests_df["staff_id"], requests_df["date"]))

    all_days = list(dates_df["date"])
    day_index = {d: i for i, d in enumerate(all_days)}

    work_days_by_staff: dict[str, dict] = {}
    if not shift_df.empty:
        for row in shift_df.itertuples():
            work_days_by_staff.setdefault(row.staff_id, {})[row.date] = row.store
    total_workdays_by_staff = {sid: len(dmap) for sid, dmap in work_days_by_staff.items()}

    def _is_idle(sid: str, d: dt.date) -> bool:
        return d not in work_days_by_staff.get(sid, {})

    def _has_day_capacity(sid: str) -> bool:
        """譛磯俣蜍､蜍呎律謨ｰ繝ｬ繝ｳ繧ｸ縺ｮ荳企剞縺ｫ縲√≠縺ｨ+1譌･蜍､蜍吶☆繧倶ｽ吝鴨縺後≠繧九°繧貞愛螳壹☆繧九・""
        etype = staff_emp_type.get(sid)
        if etype == "蝌ｱ險・:
            max_d = staff_max_workdays.get(sid)
            if pd.notna(max_d):
                return total_workdays_by_staff.get(sid, 0) < int(max_d)
            return True
        if etype == "繝代・繝・:
            role = staff_part_role.get(sid)
            _min_d, max_d = PART_ROLE_WORKDAY_RANGE.get(role, (0, len(dates_df)))
            return total_workdays_by_staff.get(sid, 0) < max_d
        return True

    def _six_consecutive_ok(sid: str, d: dt.date) -> bool:
        idx = day_index.get(d)
        if idx is None:
            return True
        worked = work_days_by_staff.get(sid, {})
        for start in range(max(0, idx - 6), idx + 1):
            end = start + 7
            if end > len(all_days):
                continue
            window = all_days[start:end]
            cnt = sum(1 for x in window if x == d or x in worked)
            if cnt > 6:
                return False
        return True

    def _store_allowed_on_date(sid: str, store: str, d: dt.date, is_wh: bool) -> bool:
        if store not in staff_allowed.get(sid, set()):
            return False
        if is_wh and store in (staff_weekend_holiday_forbidden.get(sid) or []):
            return False
        if d.weekday() in (5, 6) and store in (staff_sat_sun_forbidden.get(sid) or []):
            return False
        return True

    rows = []
    for _, day in business_days.iterrows():
        d = day["date"]
        is_wh = bool(day["is_weekend_or_holiday"])
        day_shift = shift_df[shift_df["date"] == d]
        total_slack = 0
        detail_parts = []
        used_substitutes_today: set[str] = set()
        # 縺薙・譌･縲∵焔蜍墓欠螳壹〒繝ｭ繝ｼ繧ｹ繧ｿ繝ｼ縺悟崋螳壹＆繧後※縺・ｋ繧ｹ繧ｿ繝・ヵ縺ｯ譛我ｼ大呵｣懊°繧蛾勁螟悶☆繧・        # (萓・ 10/10縺ｮ荳ｭ譚代・螻ｱ蟯｡繝ｻ逕滄ｧ偵・遒ｺ螳夐・鄂ｮ縺ｮ縺溘ａ縲√◎縺ｮ譌･縺ｯ莨代∩縺ｫ蝗槭○縺ｪ縺・縲・        locked_today = {
            name_to_id[n]
            for spec in MANUAL_STORE_ASSIGNMENTS
            if spec["date"] == d
            for n in spec["names"]
            if n in name_to_id
        }

        for store in STORES:
            store_shift = day_shift[day_shift["store"] == store]

            employees_here_all = []  # (sid, has_skill, is_home, etype)
            removable_candidates = []
            part_roles_present: set[str] = set()
            for sid in store_shift["staff_id"]:
                etype = staff_emp_type.get(sid)
                if etype in EMPLOYEE_TYPES:
                    entry = (sid, bool(staff_has_skill.get(sid)), staff_home.get(sid) == store, etype)
                    employees_here_all.append(entry)
                    if etype in _LEAVE_REMOVABLE_TYPES and sid not in locked_today:
                        removable_candidates.append(entry)
                elif etype == "繝代・繝・:
                    role = staff_part_role.get(sid)
                    if role:
                        part_roles_present.add(role)

            min_required = _min_required_employees_for_store(store, part_roles_present, is_wh)
            current_e = len(employees_here_all)
            if current_e < min_required:
                # 譌｢縺ｫ荳崎ｶｳ縺瑚ｨｱ螳ｹ縺輔ｌ縺ｦ縺・ｋ譌･(謇句虚謖・ｮ夂ｭ・縺ｯ縲√％繧御ｻ･荳頑ｸ帙ｉ縺吝愛螳壹ｒ陦後ｏ縺ｪ縺・・                continue

            surplus = max(0, current_e - min_required)
            removed_ids: set[str] = set()
            detail_names: list[str] = []

            # --- 繧ｱ繝ｼ繧ｹ1: 譌｢蟄倥・菴吝臆莠ｺ蜩｡繧偵◎縺ｮ縺ｾ縺ｾ譛我ｼ大呵｣懊↓縺吶ｋ ---
            if surplus > 0:
                sorted_candidates = sorted(removable_candidates, key=lambda p: (p[1], p[2]))
                skilled_remaining = sum(1 for e in employees_here_all if e[1])
                for sid, has_skill, _is_home, _etype in sorted_candidates:
                    if len(removed_ids) >= surplus:
                        break
                    if has_skill:
                        if skilled_remaining <= 1:
                            continue
                        skilled_remaining -= 1
                    removed_ids.add(sid)
                detail_names.extend(staff_name.get(s, s) for s in removed_ids)

            # --- 繧ｱ繝ｼ繧ｹ2: 縺｡繧・≧縺ｩ譛菴惹ｺｺ謨ｰ縺ｧ繧ゅ∽ｻ｣譖ｿ隕∝藤繧呈兜蜈･縺ｧ縺阪ｌ縺ｰ1蜷堺ｼ代ａ繧・---
            remaining_for_case2 = [c for c in removable_candidates if c[0] not in removed_ids]
            for sid, has_skill, _is_home, _etype in sorted(remaining_for_case2, key=lambda p: (p[1], p[2])):
                remaining_after_removal = [
                    e for e in employees_here_all if e[0] != sid and e[0] not in removed_ids
                ]
                if len(remaining_after_removal) >= min_required:
                    continue  # 螳溘・莉｣譖ｿ縺ｪ縺励〒繧りｶｳ繧翫※縺・◆(繧ｱ繝ｼ繧ｹ1縺ｧ諡ｾ縺・″繧後↑縺九▲縺溷・)

                remaining_skilled = sum(1 for e in remaining_after_removal if e[1])
                substitute_id = None
                for cand in staff_df.itertuples():
                    csid = cand.staff_id
                    if csid == sid or csid in removed_ids or csid in used_substitutes_today:
                        continue
                    if staff_emp_type.get(csid) not in _LEAVE_SUBSTITUTE_TYPES:
                        continue
                    if not _store_allowed_on_date(csid, store, d, is_wh):
                        continue
                    if not _is_idle(csid, d):
                        continue
                    if (csid, d) in requested_off_pairs:
                        continue  # 蟶梧悍莨・邨ｶ蟇ｾ莨・譛臥ｵｦ逕ｳ隲九′蜈･縺｣縺ｦ縺・ｋ莠ｺ縺ｯ蜻ｼ縺ｳ蜃ｺ縺輔↑縺・                    if not _has_day_capacity(csid):
                        continue
                    if has_skill and remaining_skilled == 0 and not bool(staff_has_skill.get(csid)):
                        continue  # 謚懊￠繧九・縺後せ繧ｭ繝ｫ菫晄怏閠・〒縲∽ｻ｣譖ｿ繧ゅせ繧ｭ繝ｫ縺ｪ縺励〒縺ｯ隕∽ｻｶ繧貞牡繧・                    if not _six_consecutive_ok(csid, d):
                        continue
                    substitute_id = csid
                    break

                if substitute_id is not None:
                    removed_ids.add(sid)
                    used_substitutes_today.add(substitute_id)
                    detail_names.append(
                        f"{staff_name.get(sid, sid)}竊畜staff_name.get(substitute_id, substitute_id)}莉｣譖ｿ蜿ｯ"
                    )

            if removed_ids:
                total_slack += len(removed_ids)
                detail_parts.append(f"{store}:{'縲・.join(detail_names)}")

        rows.append(
            {
                "date": d,
                "weekday_jp": day["weekday_jp"],
                "day_type": day["day_type"],
                "extra_leave_slots": total_slack,
                "detail": " / ".join(detail_parts),
            }
        )
    return pd.DataFrame(rows)


def leave_slot_badge(n: int) -> str:
    """譛臥ｵｦ霑ｽ蜉蜿門ｾ怜庄閭ｽ譫謨ｰ繧偵ヰ繝・ず逕ｨ縺ｮ遏ｭ縺・枚險縺ｫ螟画鋤縺吶ｋ縲・""
    return f"縺ゅ→{int(n)}蜷榊庄閭ｽ" if n > 0 else "貅譚ｯ(0蜷・"


def generate_post_solve_leave_announcement(
    availability_df: pd.DataFrame,
    period_label: str,
) -> str:
    """譛驕ｩ蛹也ｵ先棡縺ｫ蝓ｺ縺･縺上梧怏邨ｦ莨第嚊 霑ｽ蜉蜿門ｾ怜庄閭ｽ譌･縲阪・遉ｾ蜀・｡亥・譁・ｒ逕滓・縺吶ｋ縲・""
    available_days = availability_df[availability_df["extra_leave_slots"] > 0]

    lines = [
        f"縲須period_label}縲第怏邨ｦ莨第嚊 霑ｽ蜉蜿門ｾ怜庄閭ｽ譌･縺ｮ縺顔衍繧峨○",
        "",
        "莉･荳九・譌･遞九〒譛臥ｵｦ莨第嚊縺ｮ霑ｽ蜉蜿門ｾ励′蜿ｯ閭ｽ縺ｧ縺吶ょｸ梧悍閠・・逕ｳ隲九＠縺ｦ縺上□縺輔＞縲・,
        "",
    ]
    if available_days.empty:
        lines.append("窶ｻ迴ｾ蝨ｨ遒ｺ螳壹＠縺ｦ縺・ｋ繧ｷ繝輔ヨ縺ｧ縺ｯ縲∬ｿｽ蜉縺ｧ譛臥ｵｦ蜿門ｾ怜庄閭ｽ縺ｪ譌･縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・)
    else:
        for _, row in available_days.iterrows():
            d: dt.date = row["date"]
            lines.append(f"繝ｻ{d.month}/{d.day}・・row['weekday_jp']}・・ {leave_slot_badge(row['extra_leave_slots'])}")
        lines.append("")
        lines.append(f"蜷郁ｨ・{len(available_days)} 譌･髢薙∬ｿｽ蜉蜿門ｾ怜庄閭ｽ譫縺ゅｊ縲・)
    lines.append("")
    lines.append("窶ｻ蜈育捩鬆・・讌ｭ蜍咎・蜷医↓繧医ｊ隱ｿ謨ｴ縺輔○縺ｦ縺・◆縺縺丞ｴ蜷医′縺ゅｊ縺ｾ縺吶・)
    return "\n".join(lines)


_EMP_TYPE_SORT_ORDER = {"蠎鈴聞": 0, "豁｣遉ｾ蜩｡": 1, "蝌ｱ險・: 2, "繝代・繝・: 3}


# ---------------------------------------------------------------------------
# 繧ｷ繝輔ヨ邨先棡縺ｮ逕ｻ髱｢荳頑焔蜍慕ｷｨ髮・邇臥ｪ√″隱ｿ謨ｴ)
# ---------------------------------------------------------------------------

BLANK_LABEL = "・育ｩｺ逋ｽ・・

# 縲檎ｮ｡逅・・畑 蜈ｨ菴謎ｼ第嚊繝槭ヨ繝ｪ繝・け繧ｹ陦ｨ縲阪・髢ｲ隕ｧ繝｢繝ｼ繝・st.dataframe)縺ｧ菴ｿ縺・・# 繧ｻ繝ｫ縺ｮ蛟､縺斐→縺ｮ譁・ｭ苓牡縲ょ､縺ｮ諢丞袖縺御ｸ逶ｮ縺ｧ蛹ｺ蛻･縺ｧ縺阪ｋ繧医≧縺ｫ縺吶ｋ縺溘ａ縺ｮ驟崎牡縲・KYUKA_CELL_TEXT_COLORS: dict[str, str] = {
    BLANK_LABEL: "#B0B0B0",  # 譛ｪ逕ｳ隲・遨ｺ逋ｽ): 阮・＞繧ｰ繝ｬ繝ｼ縺ｧ逶ｮ遶九◆縺ｪ縺上☆繧・    "蟶梧悍莨・: "#1565C0",     # 髱・    "邨ｶ蟇ｾ莨・: "#E07B00",     # 繧ｪ繝ｬ繝ｳ繧ｸ
    "譛臥ｵｦ逕ｳ隲・: "#D81B7A",   # 繝斐Φ繧ｯ
}


def style_kyuka_wide_matrix(wide_df: pd.DataFrame):
    """邂｡逅・・畑繝槭ヨ繝ｪ繧ｯ繧ｹ陦ｨ(髢ｲ隕ｧ繝｢繝ｼ繝・縺ｫ縲∝､縺斐→縺ｮ譁・ｭ苓牡繧帝←逕ｨ縺励◆Styler繧定ｿ斐☆縲・
    st.dataframe 縺ｯpandas縺ｮStyler繧偵◎縺ｮ縺ｾ縺ｾ謠冗判縺ｧ縺阪ｋ縺溘ａ縲√％縺薙〒譁・ｭ苓牡縺縺代ｒ
    險ｭ螳壹＠縺欖tyler繧堤ｵ・∩遶九※繧九ょｯｾ雎｡縺ｯ譌･莉伜・縺ｮ縺ｿ(縲後せ繧ｿ繝・ヵ蜷阪榊・縺ｯ蟇ｾ雎｡螟・縲・    縺ｪ縺翫√％縺ｮ驟崎牡縺ｯ縲檎ｮ｡逅・・Δ繝ｼ繝峨〒邱ｨ髮・☆繧九阪メ繧ｧ繝・け縺薫FF縺ｮ髢ｲ隕ｧ陦ｨ遉ｺ縺ｫ縺ｮ縺ｿ
    驕ｩ逕ｨ縺輔ｌ繧九０N縺ｫ縺励◆髫帙↓陦ｨ遉ｺ縺輔ｌ繧・st.data_editor(邱ｨ髮・畑繧ｰ繝ｪ繝・ラ)縺ｯ縲・    Streamlit蛛ｴ縺ｮ謚陦鍋噪縺ｪ蛻ｶ邏・↓繧医ｊ縲∝､縺ｫ蠢懊§縺溘そ繝ｫ譁・ｭ苓牡縺ｮ謖・ｮ壹↓蟇ｾ蠢懊＠縺ｦ
    縺・↑縺・pandas Styler繝ｻcolumn_config縺ｮ縺・★繧後ｂ邱ｨ髮・Δ繝ｼ繝峨・繧ｻ繝ｫ譁・ｭ苓牡繧・    螟画峩縺吶ｋ謇区ｮｵ繧呈署萓帙＠縺ｦ縺・↑縺・縺溘ａ縲√％縺｡繧峨・蠕捺擂騾壹ｊ縺ｮ陦ｨ遉ｺ縺ｨ縺ｪ繧九・    """
    date_cols = [c for c in wide_df.columns if c != "繧ｹ繧ｿ繝・ヵ蜷・]

    def _color(val: str) -> str:
        color = KYUKA_CELL_TEXT_COLORS.get(val)
        return f"color: {color}" if color else ""

    return wide_df.style.map(_color, subset=date_cols)


# 縲檎ｮ｡逅・・Δ繝ｼ繝峨〒邱ｨ髮・☆繧九阪・邱ｨ髮・畑繧ｰ繝ｪ繝・ラ(st.data_editor)蟆ら畑縺ｮ陦ｨ險倥・# st.data_editor縺ｯ繧ｻ繝ｫ縺ｮ譁・ｭ苓牡謖・ｮ壹↓蟇ｾ蠢懊＠縺ｦ縺・↑縺・Styler繝ｻcolumn_config縺ｮ
# 縺・★繧後ｂ邱ｨ髮・Δ繝ｼ繝峨↓縺ｯ蜉ｹ縺九↑縺・縺溘ａ縲∽ｻ｣繧上ｊ縺ｫ濶ｲ莉倥″邨ｵ譁・ｭ励ｒ驕ｸ謚櫁い縺ｮ譁・ｭ怜・
# 縺昴・繧ゅ・縺ｫ蝓九ａ霎ｼ繧縺薙→縺ｧ縲∫ｷｨ髮・ｸｭ繧ゆｸ逶ｮ縺ｧ蛹ｺ蛻･縺ｧ縺阪ｋ繧医≧縺ｫ縺吶ｋ縲・# 菫晏ｭ倥＆繧後ｋ繝・・繧ｿ閾ｪ菴薙・蠕捺擂騾壹ｊ縲悟ｸ梧悍莨代咲ｭ峨・繝励Ξ繝ｼ繝ｳ縺ｪ譁・ｭ怜・縺ｮ縺ｾ縺ｾ螟峨ｏ繧峨↑縺・# (荳九・KYUKA_EMOJI_TO_KIND縺ｧ蠕ｩ蜈・＠縲〔yuka_requests_wide_to_long蛛ｴ縺ｧ豁｣隕丞喧縺吶ｋ)縲・KYUKA_KIND_EMOJI_LABELS: dict[str, str] = {
    "蟶梧悍莨・: "鳩蟶梧悍莨・,
    "邨ｶ蟇ｾ莨・: "泛邨ｶ蟇ｾ莨・,
    "譛臥ｵｦ逕ｳ隲・: "ｩｷ譛臥ｵｦ逕ｳ隲・,
}
KYUKA_EMOJI_TO_KIND: dict[str, str] = {v: k for k, v in KYUKA_KIND_EMOJI_LABELS.items()}


def to_emoji_labeled_kyuka_wide(wide_df: pd.DataFrame) -> pd.DataFrame:
    """邂｡逅・・ｷｨ髮・畑繝槭ヨ繝ｪ繧ｯ繧ｹ陦ｨ縺ｮ蛟､繧偵∬牡莉倥″邨ｵ譁・ｭ嶺ｻ倥″縺ｮ陦ｨ險倥↓螟画鋤縺励◆繧ｳ繝斐・繧定ｿ斐☆縲・
    BLANK_LABEL(縲鯉ｼ育ｩｺ逋ｽ・峨・縺ｯ縺昴・縺ｾ縺ｾ縲よ律莉伜・縺ｮ縺ｿ縺悟ｯｾ雎｡縺ｧ縲√後せ繧ｿ繝・ヵ蜷阪榊・縺ｯ
    螟画鋤縺励↑縺・Ｔt.data_editor縺ｮSelectboxColumn縺ｫ貂｡縺呵｡ｨ遉ｺ蟆ら畑縺ｮDataFrame繧・    菴懊ｋ縺溘ａ縺ｫ菴ｿ縺・螳溘ョ繝ｼ繧ｿ繝ｻ菫晏ｭ伜､縺ｫ縺ｯ蠖ｱ髻ｿ縺励↑縺・縲・    """
    date_cols = [c for c in wide_df.columns if c != "繧ｹ繧ｿ繝・ヵ蜷・]
    out = wide_df.copy()
    for col in date_cols:
        out[col] = out[col].map(lambda v: KYUKA_KIND_EMOJI_LABELS.get(v, v))
    return out


def manual_shift_date_labels(dates_df: pd.DataFrame) -> list[str]:
    """謇句虚邱ｨ髮・ユ繝ｼ繝悶Ν縺ｮ蛻苓ｦ句・縺・蝟ｶ讌ｭ譌･縺ｮ縺ｿ)繧定ｿ斐☆縲・""
    business_days_df = dates_df[dates_df["is_business_day"]]
    return [f"{d.month}/{d.day}({wd})" for d, wd in zip(business_days_df["date"], business_days_df["weekday_jp"])]


def build_manual_shift_wide(
    shift_df: pd.DataFrame,
    dates_df: pd.DataFrame,
) -> pd.DataFrame:
    """譛驕ｩ蛹也ｵ先棡(髟ｷ蠖｢蠑・繧偵∵焔蜍慕ｷｨ髮・畑縺ｮ繝ｯ繧､繝牙ｽ｢蠑・蠎苓・ﾃ怜霧讌ｭ譌･縲・繧ｻ繝ｫ1蜷・縺ｫ螟画鋤縺吶ｋ縲・
    螳滄圀縺ｮ蠎苓・蛻･莠ｺ蜩｡荳企剞(STORE_MAX_HEADCOUNT縲・蜷阪∪縺溘・3蜷・縺ｫ髢｢繧上ｉ縺壹∝・7蠎苓・繧・    荳蠕・STORE_SLOT_ROWS(4)譫縺ｧ陦ｨ遉ｺ縺吶ｋ(蠎苓・蛻･譌･蛻･繧ｷ繝輔ヨ陦ｨ繝ｻExcel蜃ｺ蜉帙→蜷後§
    繝ｬ繧､繧｢繧ｦ繝・縲ゅ％繧後↓繧医ｊ縲∫ｹ∝ｿ呎律縺ｮ蠢懈抄繧ｹ繧ｿ繝・ヵ霑ｽ蜉繧・､・焚蠎苓・髢薙・邇臥ｪ√″隱ｿ謨ｴ縺ｮ
    荳譎ら噪縺ｪ蜿励￠逧ｿ縺ｨ縺励※縲・壼ｸｸ縺ｮ莠ｺ蜩｡荳企剞繧定ｶ・∴繧区棧繧ら判髱｢荳翫〒繝励Ν繝繧ｦ繝ｳ邱ｨ髮・〒縺阪ｋ縲・    遨ｺ縺阪そ繝ｫ縺ｯ遨ｺ譁・ｭ怜・縺ｧ陦ｨ迴ｾ縺吶ｋ(陦ｨ遉ｺ譎ゅ↓BLANK_LABEL縺ｸ螟画鋤縺輔ｌ繧・縲・    """
    business_days_df = dates_df[dates_df["is_business_day"]]
    dates = list(business_days_df["date"])
    date_labels = manual_shift_date_labels(dates_df)

    rows = []
    for store in STORES:
        n_slots = STORE_SLOT_ROWS
        per_date_names: dict[dt.date, list[str]] = {}
        for d in dates:
            if shift_df.empty:
                day_rows = pd.DataFrame(columns=["name", "emp_type"])
            else:
                day_rows = shift_df[(shift_df["store"] == store) & (shift_df["date"] == d)]
            names = sorted(
                day_rows["name"].tolist(),
                key=lambda nm: _EMP_TYPE_SORT_ORDER.get(
                    day_rows.loc[day_rows["name"] == nm, "emp_type"].iloc[0], 9
                ),
            )
            per_date_names[d] = names
        for slot in range(n_slots):
            row = {"蠎苓・": store, "譫": slot + 1}
            for d, label in zip(dates, date_labels):
                names = per_date_names[d]
                row[label] = names[slot] if slot < len(names) else ""
            rows.append(row)
    return pd.DataFrame(rows)


def manual_shift_wide_to_long(
    wide_df: pd.DataFrame,
    dates_df: pd.DataFrame,
    staff_df: pd.DataFrame,
) -> pd.DataFrame:
    """謇句虚邱ｨ髮・ｾ後・繝ｯ繧､繝牙ｽ｢蠑上ｒ縲・聞蠖｢蠑・date/store/staff_id/name/emp_type)縺ｸ螟画鋤縺吶ｋ縲・
    遨ｺ谺・遨ｺ譁・ｭ怜・ 縺ｾ縺溘・ BLANK_LABEL)縺ｯ辟｡隕悶☆繧九ょｭ伜惠縺励↑縺・ｰ丞錐縺悟・縺｣縺ｦ縺・◆
    蝣ｴ蜷医ｂ螳牙・蛛ｴ縺ｧ繧ｹ繧ｭ繝・・縺吶ｋ(繧ｯ繝ｩ繝・す繝･縺励↑縺・縲・    """
    date_labels = manual_shift_date_labels(dates_df)
    label_to_date = dict(zip(date_labels, dates_df.loc[dates_df["is_business_day"], "date"]))
    name_to_id = dict(zip(staff_df["name"], staff_df["staff_id"]))
    emp_type_by_name = dict(zip(staff_df["name"], staff_df["emp_type"]))

    rows = []
    for _, r in wide_df.iterrows():
        store = r.get("蠎苓・")
        for label in date_labels:
            name = r.get(label)
            if not name or name == BLANK_LABEL:
                continue
            if name not in name_to_id:
                continue
            rows.append(
                {
                    "date": label_to_date[label],
                    "store": store,
                    "staff_id": name_to_id[name],
                    cand_etype = staff_emp_type.get(csid)
                    if cand_etype == "嘱託":
                        if len(remaining_after_removal) + 1 < min_required:
                            continue
                    elif cand_etype == "パート":
                        cand_role = staff_part_role.get(csid)
                        hypothetical_roles_present = set(part_roles_present)
                        if cand_role:
                            hypothetical_roles_present.add(cand_role)
                        recomputed_min_required = _min_required_employees_for_store(
                            store, hypothetical_roles_present, is_wh
                        )
                        if len(remaining_after_removal) < recomputed_min_required:
                            continue  # このパートを補充しても社員の必要人数を満たせない
                    "name": name,
                    "emp_type": emp_type_by_name[name],
                }
            )
    return pd.DataFrame(rows, columns=["date", "store", "staff_id", "name", "emp_type"])


def _check_store_day_pattern(
    store: str,
    e_count: int,
    part_roles_present: set[str],
    is_weekend_holiday: bool,
) -> tuple[bool, int]:
    """蠎苓・繝ｻ譖懈律蛹ｺ蛻・＃縺ｨ縺ｮ謌千ｫ九ヱ繧ｿ繝ｼ繝ｳ縺ｫ辣ｧ繧峨＠縺ｦ縲∫樟蝨ｨ縺ｮ蝨ｨ邀肴ｧ区・縺梧怏蜉ｹ縺句愛螳壹☆繧九・
    optimizer.py 縺ｮ蠎苓・蛻･繝上・繝牙宛邏・→蜷後§蛻､螳壼渕貅・蟷ｳ譌･/蝨滓律逾昴・驕輔＞繧貞性繧)縲・    謌ｻ繧雁､縺ｯ (謌千ｫ九＠縺ｦ縺・ｋ縺・ 荳崎ｶｳ莠ｺ謨ｰ縺ｮ逶ｮ螳・縲・    """
    if store == TOKUSHIGE_STORE:
        if is_weekend_holiday:
            ok = e_count >= 2
        else:
            ok = e_count >= 2 or (e_count == 1 and {"B", "C"} <= part_roles_present)
    elif store in COMBO_STORE_PART_ROLES and store != "螟ｧ豐ｻ蠎・:
        roles = set(COMBO_STORE_PART_ROLES[store])
        if is_weekend_holiday:
            ok = e_count >= 2
        else:
            ok = e_count >= 2 or (e_count == 1 and bool(part_roles_present & roles))
    elif store == "螟ｧ豐ｻ蠎・:
        ok = (e_count == 2) or (e_count == 1 and len(part_roles_present) >= 1)
    else:
        ok = e_count == 2
    shortfall = 0 if ok else max(0, 2 - e_count)
    return ok, shortfall


def check_manual_shift_alerts(
    shift_df: pd.DataFrame,
    staff_df: pd.DataFrame,
    dates_df: pd.DataFrame,
) -> dict[str, list[dict]]:
    """謇句虚邱ｨ髮・ｾ後・繧ｷ繝輔ヨ(髟ｷ蠖｢蠑・縺ｫ蟇ｾ縺励・㍾隍・・蜍､繝ｻ莠ｺ蜩｡荳崎ｶｳ繝ｻ繧ｹ繧ｭ繝ｫ荳榊惠繧貞・蛻､螳壹☆繧九・
    謌ｻ繧雁､: {"duplicates": [...], "shortages": [...], "skill_issues": [...]}

    shift_df 縺・None/遨ｺ/諠ｳ螳壼､悶・蠖｢(蠢・ｦ√↑蛻励′辟｡縺・ｭ・縺ｮ蝣ｴ蜷医・縲√∪縺譛驕ｩ蛹悶′
    螳溯｡後＆繧後※縺・↑縺・怦蠎ｦ縺ｪ縺ｩ縺ｫ蛻・ｊ譖ｿ縺医◆髫帙↓繧ｯ繝ｩ繝・す繝･縺励↑縺・ｈ縺・∝愛螳壹ｒ
    陦後ｏ縺壹↓遨ｺ縺ｮ繧｢繝ｩ繝ｼ繝・蝠城｡後↑縺玲桶縺・繧定ｿ斐☆縲・    """
    alerts: dict[str, list[dict]] = {"duplicates": [], "shortages": [], "skill_issues": []}

    if shift_df is None or not isinstance(shift_df, pd.DataFrame) or shift_df.empty:
        return alerts
    if "date" not in shift_df.columns or "store" not in shift_df.columns or "staff_id" not in shift_df.columns:
        return alerts

    staff_emp_type = dict(zip(staff_df["staff_id"], staff_df["emp_type"]))
    staff_has_skill = dict(zip(staff_df["staff_id"], staff_df["has_skill"]))
    staff_part_role = dict(zip(staff_df["staff_id"], staff_df["part_role"]))
    staff_name = dict(zip(staff_df["staff_id"], staff_df["name"]))

    for (sid, d), grp in shift_df.groupby(["staff_id", "date"]):
        if len(grp) > 1:
            alerts["duplicates"].append(
                {
                    "date": d,
                    "name": staff_name.get(sid, sid),
                    "stores": grp["store"].unique().tolist(),
                }
            )

    business_days_df = dates_df[dates_df["is_business_day"]]
    for _, day in business_days_df.iterrows():
        d = day["date"]
        is_wh = bool(day["is_weekend_or_holiday"])
        day_shift = shift_df[shift_df["date"] == d]
        for store in STORES:
            store_shift = day_shift[day_shift["store"] == store]
            e_count = 0
            part_roles_present: set[str] = set()
            has_skill_here = False
            staffed = False
            for sid in store_shift["staff_id"]:
                etype = staff_emp_type.get(sid)
                if etype in EMPLOYEE_TYPES:
                    e_count += 1
                    staffed = True
                elif etype == "繝代・繝・:
                    role = staff_part_role.get(sid)
                    if role:
                        part_roles_present.add(role)
                    staffed = True
                if staff_has_skill.get(sid):
                    has_skill_here = True

            ok, shortfall = _check_store_day_pattern(store, e_count, part_roles_present, is_wh)
            if not ok:
                detail = f"遉ｾ蜩｡{e_count}蜷・
                if part_roles_present:
                    detail += f"+繝代・繝・len(part_roles_present)}蜷・
                alerts["shortages"].append({"date": d, "store": store, "荳崎ｶｳ謨ｰ": shortfall, "隧ｳ邏ｰ": detail})
            if staffed and not has_skill_here:
                alerts["skill_issues"].append({"date": d, "store": store})

    return alerts


# ---------------------------------------------------------------------------
# Excel繧ｨ繧ｯ繧ｹ繝昴・繝・# ---------------------------------------------------------------------------


def _generate_pastel_palette(n: int, lightness: float = 0.84, saturation: float = 0.55) -> list[str]:
    """濶ｲ逶ｸ繧貞插遲峨↓蜑ｲ繧頑険縺｣縺溘∬ｦ句・縺代ｄ縺吶＞繝代せ繝・Ν邉ｻHEX繧ｫ繝ｩ繝ｼ繧地濶ｲ逕滓・縺吶ｋ縲・
    STAFF_NAME_COLORS 縺ｫ蟇ｾ蠢懊′縺ｪ縺・ｰ丞錐(=繝・ヵ繧ｩ繝ｫ繝・9蜷堺ｻ･螟・蜷代￠縺ｮ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ逕ｨ縲・    """
    colors = []
    for i in range(n):
        hue = i / max(n, 1)
        r, g, b = colorsys.hls_to_rgb(hue, lightness, saturation)
        colors.append(f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}")
    return colors


# 蜈ｨ19蜷阪・隴伜挨繧ｫ繝ｩ繝ｼ(豌丞錐 -> HEX)縲る團謗･縺励ｄ縺吶＞繝｡繝ｳ繝舌・蜷悟｣ｫ縺ｧ繧ゆｸ逶ｮ縺ｧ隕句・縺代ｉ繧後ｋ繧医≧縲・# 濶ｲ逶ｸ縺悟､ｧ縺阪￥逡ｰ縺ｪ繧九ヱ繧ｹ繝・Ν/隴伜挨繧ｫ繝ｩ繝ｼ繧貞句挨縺ｫ驕ｸ螳壹＠縺ｦ縺・ｋ縲・STAFF_NAME_COLORS: dict[str, str] = {
    "逕滄ｧ・: "BFDFF5",    # 繝代せ繝・Ν繝悶Ν繝ｼ(豌ｴ濶ｲ)
    "霎ｻ譛ｬ": "C2F0DB",    # 繝溘Φ繝医げ繝ｪ繝ｼ繝ｳ(阮・ｷ・
    "蜀・伐": "FFF3B8",    # 繧ｽ繝輔ヨ繧､繧ｨ繝ｭ繝ｼ(阮・ｻ・
    "蟆乗棊": "FFC9CE",    # 繧ｳ繝ｼ繝ｩ繝ｫ繝斐Φ繧ｯ(阮・ｴ・
    "蜉阯､": "E1D3F5",    # 繝ｩ繝吶Φ繝繝ｼ(阮・ｴｫ)
    "逕ｰ荳ｭ": "FFDCB8",    # 繧｢繝励Μ繧ｳ繝・ヨ(阮・ｩ・
    "髟ｷ轢ｬ": "A9E8E8",    # 繧ｹ繧ｫ繧､繧ｷ繧｢繝ｳ(譏弱ｋ縺・搨邱・
    "逵溽伐": "E4F0A0",    # 繝ｬ繝｢繝ｳ繝ｩ繧､繝(鮟・ｷ・
    "闍･譫・: "F7C9DC",    # 繝ｭ繝ｼ繧ｺ繝斐Φ繧ｯ(阮・ヴ繝ｳ繧ｯ)
    "荳ｭ譚・: "A0E4DC",    # 繧ｿ繝ｼ繧ｳ繧､繧ｺ(髱堤ｷ・
    "闍･譚ｾ": "E9DCBB",    # 繧ｵ繝ｳ繝峨・繝ｼ繧ｸ繝･(阮・幻/繝吶・繧ｸ繝･)
    "蜷臥伐": "E2D48A",    # 繧ｪ繝ｪ繝ｼ繝悶ざ繝ｼ繝ｫ繝・阮・が繝ｪ繝ｼ繝・
    "遶ｹ蜀・: "D6B8E6",    # 繝｢繝ｼ繝ｴ(阮・陸濶ｲ)
    "螻ｱ蟯｡": "DCDCDC",    # 繝壹・繝ｫ繧ｰ繝ｬ繝ｼ/繧ｷ繝ｫ繝舌・(譏守・濶ｲ)
    "蟆ｾ貔､": "FFB6A3",    # 繧ｵ繝ｼ繝｢繝ｳ(阮・し繝ｼ繝｢繝ｳ繝斐Φ繧ｯ)
    "荳榊虚驥・: "A0E0BC",  # 繧ｨ繝｡繝ｩ繝ｫ繝・阮・ｷｱ邱・
    "蜑咲伐": "E3AEDD",    # 繧ｪ繝ｼ繧ｭ繝・ラ(襍､邏ｫ)
    "譟ｴ逕ｰ": "DCF3FA",    # 繧｢繧､繧ｹ繝悶Ν繝ｼ(縺斐￥阮・＞髱・
    "驥朱％": "FFDAC0",    # 繝斐・繝・阮・｡・牡)
}


def assign_staff_colors(staff_df: pd.DataFrame) -> dict[str, str]:
    """繧ｹ繧ｿ繝・ヵ1蜷阪＃縺ｨ縺ｫ蝗ｺ譛峨・隴伜挨閭梧勹濶ｲ(HEX荳・譯・繧貞牡繧雁ｽ薙※繧九・
    縺ｾ縺・STAFF_NAME_COLORS 縺ｮ豌丞錐荳閾ｴ繧貞━蜈医＠縲∬ｩｲ蠖薙′縺ｪ縺・せ繧ｿ繝・ヵ(豌丞錐螟画峩譎ゅｄ
    繝・ヵ繧ｩ繝ｫ繝井ｻ･螟悶・霑ｽ蜉繧ｹ繧ｿ繝・ヵ)縺ｫ縺ｯ閾ｪ蜍慕函謌舌＠縺溘ヱ繧ｹ繝・Ν繧ｫ繝ｩ繝ｼ繧貞牡繧雁ｽ薙※繧九・    """
    fallback_palette = _generate_pastel_palette(len(staff_df))
    colors: dict[str, str] = {}
    for i, row in enumerate(staff_df.itertuples()):
        colors[row.staff_id] = STAFF_NAME_COLORS.get(row.name, fallback_palette[i % len(fallback_palette)])
    return colors


def build_export_workbook(
    shift_df: pd.DataFrame,
    dates_df: pd.DataFrame,
    staff_df: pd.DataFrame,
    requests_df: pd.DataFrame | None = None,
) -> bytes:
    """蠎苓・蛻･譌･蛻･繧ｷ繝輔ヨ陦ｨ繝ｻ繧ｹ繧ｿ繝・ヵ蛻･蜃ｺ蜍､荳隕ｧ陦ｨ縺ｮ2繧ｷ繝ｼ繝域ｧ区・Excel繧堤函謌舌☆繧九・
    shift_df: columns = [date, store, staff_id, name, emp_type]  (蜃ｺ蜍､縺檎｢ｺ螳壹＠縺溯｡後・縺ｿ)

    縲悟ｺ苓・蛻･譌･蛻･繧ｷ繝輔ヨ陦ｨ縲阪・迴ｾ蝣ｴ驟榊ｸ・・謗ｲ遉ｺ縺ｫ縺昴・縺ｾ縺ｾ菴ｿ縺医ｋ繧医≧縲・繧ｻ繝ｫ縺ｫ1蜷阪・縺ｿ繧・    驟咲ｽｮ縺吶ｋ邵ｦ蛻・牡繝ｬ繧､繧｢繧ｦ繝医→縺励∽ｸ区ｮｵ縺ｫ莨第律繧ｹ繧ｿ繝・ヵ荳隕ｧ(襍､譁・ｭ・繧剃ｻ倥☆縲・    莨第律繧ｹ繧ｿ繝・ヵ縺ｮ縺・■縲∝ｸ梧悍莨代・邨ｶ蟇ｾ莨代・譛臥ｵｦ逕ｳ隲九・縺・★繧後°繧呈悽莠ｺ縺・    逕ｳ隲九＠縺ｦ縺・◆蝣ｴ蜷医・襍､譁・ｭ暦ｼ句､ｪ蟄励→縺励、I縺瑚・蜍募牡繧雁ｽ薙※縺励◆蜈ｬ莨・騾壼ｸｸ縺ｮ襍､譁・ｭ・
    縺ｨ荳逶ｮ縺ｧ蛹ｺ蛻･縺ｧ縺阪ｋ繧医≧縺ｫ縺吶ｋ縲・    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    dates = list(dates_df["date"])
    date_labels = [
        f"{d.month}/{d.day}({row.weekday_jp})" + ("\n莨第･ｭ" if getattr(row, "is_special_closure", False) else "")
        for d, row in zip(dates, dates_df.itertuples())
    ]
    n_date_cols = len(dates)
    total_cols = 1 + n_date_cols  # 蛻輸=蠎苓・蜷・隕句・縺励∽ｻ･髯阪・譌･莉伜・

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "蠎苓・蛻･譌･蛻･繧ｷ繝輔ヨ陦ｨ"

    # 1陦檎岼(繝倥ャ繝繝ｼ陦・縺ｮ繝輔か繝ｳ繝医し繧､繧ｺ縲ょ､ｪ蟄励・荳ｭ螟ｮ謠・∴縺ｯ蠕捺擂騾壹ｊ邯ｭ謖√☆繧九・    HEADER_FONT_SIZE = 9

    thin = Side(style="thin", color="BFBFBF")
    thick = Side(style="medium", color="404040")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True, size=HEADER_FONT_SIZE)
    store_fills = [PatternFill("solid", fgColor="EDEDED"), PatternFill("solid", fgColor="FFFFFF")]
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    red_font = Font(color="FF0000")
    red_bold_font = Font(color="FF0000", bold=True)
    requested_off_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 蟶梧悍莨・邨ｶ蟇ｾ莨・譛臥ｵｦ逕ｳ隲九・縺・★繧後°繧呈悽莠ｺ縺檎筏隲九＠縺ｦ縺・◆(豌丞錐,譌･莉・縺ｮ髮・粋縲・    # 隧ｲ蠖薙☆繧倶ｼ第律繧ｹ繧ｿ繝・ヵ蜷阪・襍､譁・ｭ暦ｼ句､ｪ蟄励↓縺励※縲、I閾ｪ蜍募牡繧雁ｽ薙※縺ｮ蜈ｬ莨代→蛹ｺ蛻･縺吶ｋ縲・    requested_off_pairs: set[tuple[str, dt.date]] = set()
    if requests_df is not None and not requests_df.empty:
        requested_off_pairs = set(zip(requests_df["name"], requests_df["date"]))

    # 譌･莉倩ｦ句・縺・蝨滓屆/譌･譖懊・逾晄律/蟷ｳ譌･/迚ｹ蛻･莨第･ｭ譌･)縺ｮ蠑ｷ隱ｿ驟崎牡
    SAT_FILL, SAT_FONT = "DCE6F1", "002060"
    SUN_HOLIDAY_FILL, SUN_HOLIDAY_FONT = "FCE4D6", "C00000"
    WEEKDAY_HEADER_FILL, WEEKDAY_HEADER_FONT = "E2EFDA", "375623"
    SPECIAL_CLOSURE_FILL, SPECIAL_CLOSURE_FONT = "BFBFBF", "404040"

    special_closure_by_date = dict(zip(dates_df["date"], dates_df.get("is_special_closure", False)))

    def _date_header_style(d: dt.date) -> tuple[str, str]:
        if special_closure_by_date.get(d):
            return SPECIAL_CLOSURE_FILL, SPECIAL_CLOSURE_FONT
        if jpholiday.is_holiday(d) or d.weekday() == 6:
            return SUN_HOLIDAY_FILL, SUN_HOLIDAY_FONT
        if d.weekday() == 5:
            return SAT_FILL, SAT_FONT
        return WEEKDAY_HEADER_FILL, WEEKDAY_HEADER_FONT

    def _set_bottom_border(row_idx: int, style: Side):
        for col in range(1, total_cols + 1):
            c = ws1.cell(row=row_idx, column=col)
            b = c.border
            c.border = Border(left=b.left, right=b.right, top=b.top, bottom=style)

    # --- 繝倥ャ繝繝ｼ陦・---------------------------------------------------------
    ws1.cell(row=1, column=1, value="蠎苓・").font = header_font
    ws1.cell(row=1, column=1).fill = header_fill
    ws1.cell(row=1, column=1).alignment = center
    ws1.cell(row=1, column=1).border = cell_border
    for j, (label, d) in enumerate(zip(date_labels, dates), start=2):
        fill_hex, font_hex = _date_header_style(d)
        c = ws1.cell(row=1, column=j, value=label)
        c.font = Font(color=font_hex, bold=True, size=HEADER_FONT_SIZE)
        c.fill = PatternFill("solid", fgColor=fill_hex)
        c.alignment = center
        c.border = cell_border

    # --- 繧ｹ繧ｿ繝・ヵ蝗ｺ譛峨・繝代せ繝・Ν閭梧勹濶ｲ -----------------------------------------
    staff_colors = assign_staff_colors(staff_df)
    name_to_id = dict(zip(staff_df["name"], staff_df["staff_id"]))

    def _staff_fill(name: str) -> PatternFill | None:
        sid = name_to_id.get(name)
        color = staff_colors.get(sid)
        return PatternFill("solid", fgColor=color) if color else None

    # --- 蠎苓・縺斐→縺ｮ蜃ｺ蜍､閠・ヶ繝ｭ繝・け(1繧ｻ繝ｫ1蜷阪・邵ｦ蛻・牡) ----------------------------
    current_row = 2
    emp_type_by_staff = dict(zip(shift_df["staff_id"], shift_df["emp_type"])) if not shift_df.empty else {}

    # 縲悟ｺ苓・蛻･譌･蛻･繧ｷ繝輔ヨ陦ｨ縲阪す繝ｼ繝医・縲∝ｮ滄圀縺ｮ蠎苓・蛻･荳企剞莠ｺ謨ｰ(STORE_MAX_HEADCOUNT縲・    # 2蜷阪∪縺溘・3蜷・縺ｫ髢｢繧上ｉ縺壹∝・7蠎苓・繧剃ｸ蠕九檎ｸｦ4陦・1縲・譫)縲阪〒蜃ｺ蜉帙☆繧・    # (蜊ｰ蛻ｷ繝ｻ迴ｾ蝣ｴ驟榊ｸ・凾縺ｮ繝ｬ繧､繧｢繧ｦ繝医ｒ蠎苓・髢薙〒邨ｱ荳縺吶ｋ縺溘ａ縺ｮ陦ｨ遉ｺ荳翫・莉墓ｧ倥〒縺ゅｊ縲・    # 繧ｽ繝ｫ繝舌・蛛ｴ縺ｮ螳滄圀縺ｮ蠎苓・蛻･莠ｺ蜩｡荳企剞繧ｭ繝｣繝・・(2蜷・3蜷・繧貞､画峩縺吶ｋ繧ゅ・縺ｧ縺ｯ縺ｪ縺・縲・    # 螳滉ｺｺ謨ｰ縺・蜷阪↓貅縺溘↑縺・ｺ苓・繝ｻ譌･縺ｯ縲∽ｽ吶▲縺滓棧縺瑚・蜍慕噪縺ｫ遨ｺ谺・↓縺ｪ繧九・    # (逕ｻ髱｢荳翫・謇句虚邱ｨ髮・ユ繝ｼ繝悶Ν縺ｨ蜈ｱ騾壹・螳壽焚 STORE_SLOT_ROWS 繧剃ｽｿ逕ｨ縺励∽ｸ｡閠・・
    # 陦梧焚縺碁｣溘＞驕輔ｏ縺ｪ縺・ｈ縺・↓縺吶ｋ)
    EXCEL_STORE_SLOT_ROWS = STORE_SLOT_ROWS

    for idx, store in enumerate(STORES):
        n_rows = EXCEL_STORE_SLOT_ROWS
        start_row = current_row
        fill = store_fills[idx % 2]

        ws1.merge_cells(start_row=start_row, start_column=1, end_row=start_row + n_rows - 1, end_column=1)
        store_cell = ws1.cell(row=start_row, column=1, value=store)
        store_cell.font = Font(bold=True)
        store_cell.alignment = center
        store_cell.fill = fill
        store_cell.border = cell_border

        for j, d in enumerate(dates, start=2):
            if special_closure_by_date.get(d):
                # 迚ｹ蛻･莨第･ｭ譌･: 蜈ｨ蠎嶺ｼ第･ｭ縺ｮ縺溘ａ縲∝ｺ苓・陦後・繧ｰ繝ｬ繝ｼ繧｢繧ｦ繝医＠縲御ｼ第･ｭ縲阪→陦ｨ遉ｺ縺吶ｋ縲・                for r_offset in range(n_rows):
                    row_idx = start_row + r_offset
                    cell = ws1.cell(row=row_idx, column=j, value="莨第･ｭ")
                    cell.alignment = center
                    cell.font = Font(color=SPECIAL_CLOSURE_FONT)
                    cell.fill = PatternFill("solid", fgColor=SPECIAL_CLOSURE_FILL)
                    cell.border = cell_border
                continue

            if shift_df.empty:
                day_rows = pd.DataFrame(columns=["name", "emp_type"])
            else:
                day_rows = shift_df[(shift_df["store"] == store) & (shift_df["date"] == d)]
            names = sorted(
                day_rows["name"].tolist(),
                key=lambda nm: _EMP_TYPE_SORT_ORDER.get(
                    day_rows.loc[day_rows["name"] == nm, "emp_type"].iloc[0], 9
                ),
            )
            for r_offset in range(n_rows):
                row_idx = start_row + r_offset
                value = names[r_offset] if r_offset < len(names) else ""
                cell = ws1.cell(row=row_idx, column=j, value=value)
                cell.alignment = center
                cell.fill = _staff_fill(value) or fill
                cell.border = cell_border

        for r_offset in range(n_rows):
            ws1.cell(row=start_row + r_offset, column=1).fill = fill

        _set_bottom_border(start_row + n_rows - 1, thick)
        current_row += n_rows

    # --- 莨第律繧ｹ繧ｿ繝・ヵ荳隕ｧ繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ(襍､譁・ｭ・ ------------------------------------
    current_row += 1
    section_row = current_row
    ws1.merge_cells(start_row=section_row, start_column=1, end_row=section_row, end_column=total_cols)
    section_cell = ws1.cell(row=section_row, column=1, value="縲蝉ｼ第律繧ｹ繧ｿ繝・ヵ・亥・莨代・譛我ｼ代・蟶梧悍莨托ｼ峨・)
    section_cell.font = Font(bold=True, color="C00000")
    section_cell.alignment = Alignment(horizontal="left", vertical="center")
    current_row += 1

    all_staff_names = list(staff_df["name"])
    off_lists: list[list[str]] = []
    for d in dates:
        working_names = set(shift_df[shift_df["date"] == d]["name"]) if not shift_df.empty else set()
        off_lists.append([n for n in all_staff_names if n not in working_names])
    max_off = max((len(lst) for lst in off_lists), default=0)
    max_off = max(max_off, 1)

    off_start_row = current_row
    ws1.merge_cells(start_row=off_start_row, start_column=1, end_row=off_start_row + max_off - 1, end_column=1)
    off_label_cell = ws1.cell(row=off_start_row, column=1, value="莨第律繧ｹ繧ｿ繝・ヵ")
    off_label_cell.font = Font(bold=True)
    off_label_cell.alignment = center
    off_label_cell.border = cell_border

    for col_idx, off_names in enumerate(off_lists):
        j = col_idx + 2
        d = dates[col_idx]
        for r_offset in range(max_off):
            row_idx = off_start_row + r_offset
            value = off_names[r_offset] if r_offset < len(off_names) else ""
            cell = ws1.cell(row=row_idx, column=j, value=value)
            cell.alignment = center
            cell.border = cell_border
            is_requested = (value, d) in requested_off_pairs
            if is_requested:
                # 莠句燕逕ｳ隲倶ｼ代∩縺ｯ縲∝倶ｺｺ蝗ｺ譛芽牡繧医ｊ蜆ｪ蜈医＠縺ｦ鮟・牡閭梧勹・玖ｵ､螟ｪ蟄励↓縺吶ｋ縲・                cell.font = red_bold_font
                cell.fill = requested_off_fill
            else:
                cell.font = red_font
                staff_fill = _staff_fill(value)
                if staff_fill:
                    cell.fill = staff_fill

    for r_offset in range(max_off):
        ws1.cell(row=off_start_row + r_offset, column=1).border = cell_border

    # --- 蛻怜ｹ・・陦碁ｫ倥・蜊ｰ蛻ｷ險ｭ螳・---------------------------------------------------
    ws1.column_dimensions[get_column_letter(1)].width = 16
    # B蛻・2蛻礼岼)縲廣G蛻・33蛻礼岼)縺ｯ譌･蛻･縺ｮ蜷・・縺ｨ縺励※荳蠕・43px逶ｸ蠖・蟷・.4)縺ｫ邨ｱ荳縺吶ｋ縲・    # 螳溘ョ繝ｼ繧ｿ縺・3蛻励↓貅縺溘↑縺・怦縺ｧ繧ゅ√ユ繝ｳ繝励Ξ繝ｼ繝医→縺励※AG蛻励∪縺ｧ蟷・ｒ謠・∴縺ｦ縺翫￥縲・    for j in range(2, 34):
        ws1.column_dimensions[get_column_letter(j)].width = 5.4
    for j in range(34, total_cols + 1):
        ws1.column_dimensions[get_column_letter(j)].width = 13

    # 1陦檎岼(繝倥ャ繝繝ｼ)縺ｯ 52px逶ｸ蠖・39pt)縲・陦檎岼縲懷ｺ苓・繝悶Ο繝・け譛邨り｡・2陦檎岼+7蠎苓・ﾃ・    # 4譫-1=29陦檎岼)縺ｯ 29px逶ｸ蠖・21.75pt)縺ｫ邨ｱ荳縺吶ｋ(蠎苓・繝悶Ο繝・け縺ｮ陦梧焚繧貞崋螳壼､縺ｧ
    # 蜀崎ｨ育ｮ励○縺壹∝ｮ滄圀縺ｫ謠冗判縺励◆EXCEL_STORE_SLOT_ROWSﾃ祐TORES謨ｰ縺九ｉ蜍慕噪縺ｫ蟆主・縺吶ｋ)縲・    header_row_end = 1
    store_rows_end = header_row_end + EXCEL_STORE_SLOT_ROWS * len(STORES)  # 2陦檎岼+28陦・29陦檎岼
    ws1.row_dimensions[1].height = 39.0
    for r in range(2, store_rows_end + 1):
        ws1.row_dimensions[r].height = 21.75

    ws1.freeze_panes = "B2"
    ws1.page_setup.orientation = "landscape"
    ws1.page_setup.paperSize = ws1.PAPERSIZE_A3
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.print_title_rows = "1:1"

    # --- 繧ｷ繝ｼ繝・: 繧ｹ繧ｿ繝・ヵ蛻･蜃ｺ蜍､荳隕ｧ陦ｨ -----------------------------------
    ws2 = wb.create_sheet("繧ｹ繧ｿ繝・ヵ蛻･蜃ｺ蜍､荳隕ｧ陦ｨ")
    header2 = ["繧ｹ繧ｿ繝・ヵ", "蛹ｺ蛻・] + date_labels + ["蜃ｺ蜍､譌･謨ｰ"]
    for j, label in enumerate(header2, start=1):
        c = ws2.cell(row=1, column=j, value=label)
        c.alignment = center
        c.border = cell_border
        if 3 <= j <= 2 + n_date_cols:
            fill_hex, font_hex = _date_header_style(dates[j - 3])
            c.font = Font(color=font_hex, bold=True, size=HEADER_FONT_SIZE)
            c.fill = PatternFill("solid", fgColor=fill_hex)
        else:
            c.font = header_font
            c.fill = header_fill

    for i, s in enumerate(staff_df.itertuples(), start=2):
        sid = s.staff_id
        row_fill = PatternFill("solid", fgColor=staff_colors[sid]) if sid in staff_colors else None
        name_cell = ws2.cell(row=i, column=1, value=s.name)
        name_cell.alignment = center
        name_cell.border = cell_border
        emp_cell = ws2.cell(row=i, column=2, value=s.emp_type)
        emp_cell.alignment = center
        emp_cell.border = cell_border
        if row_fill:
            name_cell.fill = row_fill
            emp_cell.fill = row_fill
        work_count = 0
        for j, d in enumerate(dates, start=3):
            match = shift_df[(shift_df["staff_id"] == sid) & (shift_df["date"] == d)] if not shift_df.empty else pd.DataFrame()
            value = match.iloc[0]["store"] if not match.empty else ""
            if not match.empty:
                work_count += 1
            day_cell = ws2.cell(row=i, column=j, value=value)
            day_cell.alignment = center
            day_cell.border = cell_border
            if row_fill:
                day_cell.fill = row_fill
        count_cell = ws2.cell(row=i, column=2 + n_date_cols + 1, value=work_count)
        count_cell.alignment = center
        count_cell.border = cell_border
        if row_fill:
            count_cell.fill = row_fill

    ws2.column_dimensions[get_column_letter(1)].width = 12
    ws2.column_dimensions[get_column_letter(2)].width = 8
    # 譌･蛻･縺ｮ蜷・・(C蛻嶺ｻ･髯阪∝ｺ苓・蛻･譌･蛻･繧ｷ繝輔ヨ陦ｨ縺ｮB縲廣G蛻励↓逶ｸ蠖・縺ｯ蜷後§縺丞ｹ・.4縺ｫ邨ｱ荳縺吶ｋ縲・    for j in range(3, 3 + n_date_cols):
        ws2.column_dimensions[get_column_letter(j)].width = 5.4
    ws2.column_dimensions[get_column_letter(3 + n_date_cols)].width = 10

    # 1陦檎岼(繝倥ャ繝繝ｼ)縲・8陦檎岼縺ｮ陦碁ｫ倥ｒ 52px逶ｸ蠖・39pt)縺ｫ邨ｱ荳縺吶ｋ縲・    for r in range(1, 19):
        ws2.row_dimensions[r].height = 39.0

    ws2.freeze_panes = "C2"
    ws2.page_setup.orientation = "landscape"
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A3
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0
    ws2.sheet_properties.pageSetUpPr.fitToPage = True

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
