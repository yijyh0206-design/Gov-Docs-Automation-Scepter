# -*- coding: utf-8 -*-
"""
gov_docs_batch_part23.py (v1.7.1)
功能：七類公文清單匯入、去重、統計並匯出為單一 Excel 檔。

115.10.30（test 資料夾開發中，尚未升版）：
- 新增「人民申請-建管系統」分頁自動填：讀建管系統下載的「建築相關執照人民申請案件清冊_工作天_*.ods」
  （或直接讀系統下載的 .zip），取 13 欄原樣填入（備註、流程留空給人工），總表建管件數一律算在建設處。
  沒有這個檔時維持舊做法（分頁清空留給人工），不會讓整個轉檔失敗。見 process_BM()／find_bm_file()。

v1.7.1（115.09.23）：
- run_headless() 新增 summary_sink 參數，跑完把結構化收工資料（產出路徑／各類筆數／
  待確認清單）交給呼叫端，給新的單一視窗 GUI（一鍵執行_GUI.py）用，不用解析 log 文字；
  沒有帶這個參數時行為不變，既有呼叫端（CLI --auto、App._run()）不受影響。
- 新增收工摘要（_format_summary）跟輸出檔路徑標記檔（write_output_marker，寫
  _上次輸出.txt），給 一鍵執行.bat／GUI 收工時用。
- ods_report.py 的收工提醒文字拆成「固定提醒」跟「本次列管案件」兩則，避免使用者看到
  「需要人工確認 1 項」誤以為是單一具體問題（詳見 test\開發紀錄.md 115.09.23 補充三）。
- 本檔案的轉檔／去重／H 列管核心邏輯這次沒有變動，是相容性新增，所以升修訂版
  （v1.7.0 → v1.7.1）而不是次版。

v1.7.0（115.09.21）：
- 新增「從資料夾自動帶入」：讀取「報表擷取_前段七類」抓好的 A_～G_*.csv。
- 修正去重：公文文號右邊補空白，導致找得到重複卻刪不掉（v1.6.8 紀錄寫移除 N 筆，實際 0 筆）。
- 新增 H 列管案件（列管案件閘清單＋5050 補充，只收監審議經、限辦日早於抓取日、工作天）；與限期公文重複者從限期刪除。
- 修正陳情案件（C）的 cSCODE_DATA 欄未改名，改為「公文次性質」；轉檔時若還有未改名的代碼欄會在紀錄提醒。
"""

import os
import re
import datetime as dt
from typing import Dict, Tuple, List, Optional

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys

# --- 打包成 exe 後，主控台預設編碼可能是系統內碼（如繁中 Windows 的 cp950），
#     會讓帶有特殊符號（如 ⚠）的訊息輸出時噴 UnicodeEncodeError，這裡統一改成 UTF-8。 ---
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# --- 嘗試載入 chardet 偵測編碼 ---
try:
    import chardet
    HAVE_CHARDET = True
except Exception:
    HAVE_CHARDET = False

def resource_path(rel_path: str) -> str:
    """ 取得打包後 exe 內的資源路徑 """
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, rel_path)

# ================= 共用工具與常數 =================

NAME_MAP = {
    "A": "一般公文",
    "B": "人民申請案公文",
    "C": "陳情案公文",
    "D": "訴願案公文",
    "E": "專案公文",
    "F": "限期公文",
    "G": "特殊性案件公文",
    "H": "列管案件",
    "BM": "人民申請-建管系統",
}
# 收工摘要／GUI 顯示各類筆數的順序（115.10.30 加入建管系統 BM）
SUMMARY_ORDER = list("ABCDEFGH") + ["BM"]

FIRST_TITLES = {
    "A": "宜蘭縣政府及所屬機關一般公文逾期6天以上未結案件",
    "B": "宜蘭縣政府及所屬機關人民申請公文逾期6天以上未結案件",
    "C": "宜蘭縣政府及所屬機關陳情公文逾期6天以上未結案件",
    "D": "宜蘭縣政府及所屬機關訴願公文逾期1天以上未結案件",
    "E": "宜蘭縣政府及所屬機關專案公文逾期1天以上未結案件",
    "F": "宜蘭縣政府及所屬機關限期公文逾期1天以上未結案件",
    "G": "宜蘭縣政府及所屬機關特殊性案件公文逾期1天以上未結案件",
    "H": "宜蘭縣政府及所屬機關列管案件逾期1天以上未結案件",
    "BM": "宜蘭縣政府及所屬機關人民申請-建管系統逾期6天以上未結案案件一覽表",
}

def roc_today_str() -> str:
    d = dt.date.today()
    return f"{d.year - 1911:03d}{d.month:02d}{d.day:02d}"

def second_title_row() -> str:
    d = dt.date.today()
    return f"截至  {d.year - 1911} 年  {d.month:02d} 月  {d.day:02d} 日止未結案"

def detect_encoding(path: str) -> str:
    if not HAVE_CHARDET:
        return "utf-8-sig"
    with open(path, "rb") as f:
        raw = f.read(200000)
    enc = chardet.detect(raw).get("encoding") or "utf-8-sig"
    return enc

def read_csv_safely(path: str) -> pd.DataFrame:
    enc = detect_encoding(path)
    try:
        return pd.read_csv(path, encoding=enc)
    except Exception:
        return pd.read_csv(path, encoding="big5", errors="ignore")

def normalize_str_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()

# ================= 排序與權重 =================

_STROKES: Dict[str, int] = {
    "一": 1, "丨": 1, "丶": 1, "丿": 1, "乙": 1, "亅": 1,
    "人": 2, "入": 2, "八": 2, "十": 2, "丁": 2, "二": 2,
    "工": 3, "山": 3, "川": 3, "己": 3, "已": 3, "土": 3, "士": 3,
    "大": 3, "女": 3, "子": 3, "小": 3, "口": 3,
    "交": 6, "地": 6, "壯": 7, "社": 7, "宜": 8, "長": 8,
    "建": 9, "南": 9, "政": 9, "計": 9, "縣": 9,
    "原": 10, "員": 10, "家": 10, "海": 10, "消": 10, "秘": 10, 
    "動": 11, "教": 11, "勞": 11, "運": 13,
    "農": 13, "殯": 14, "衛": 16, "頭": 16,
    "環": 17, "礁": 18, "蘇": 19, "警": 19, "蘭": 20,
    "水": 4, "民": 5, "文": 4, "戶": 4, "旅": 10, "處": 11, "局": 7, "所": 8, "館": 16, "場": 12,
}

def est_strokes(name: str) -> int:
    if not name: return 0
    ch = name[0]
    return _STROKES.get(ch, 30 + (ord(ch) % 1000))

def unit_level(name: str) -> int:
    """v1.7.0：比照人工週報，先排「處」、再排「局」、最後其他（所、館、場…）。"""
    n = str(name).strip()
    if n.endswith("處"): return 0
    if n.endswith("局"): return 1
    return 2

def sort_key_by_unit(name: str):
    name = str(name).strip()
    return (unit_level(name), est_strokes(name), name)

# ================= 資料處理核心 =================

def drop_columns_by_label_ordered(df: pd.DataFrame, drops: List[str]) -> pd.DataFrame:
    exist = [c for c in drops if c in df.columns]
    return df.drop(columns=exist)

def rename_columns(df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    cols = {c: mapping[c] for c in df.columns if c in mapping}
    return df.rename(columns=cols)

def detect_overdue_column(df: pd.DataFrame) -> Optional[str]:
    for c in df.columns:
        if "逾期天數" in c: return c
    return None

def sort_multi(df: pd.DataFrame) -> pd.DataFrame:
    od = detect_overdue_column(df)
    if od:
        df[od] = pd.to_numeric(df[od], errors="coerce").fillna(0)
    if "承辦單位" in df.columns:
        df["_unit_key1"] = df["承辦單位"].map(lambda x: sort_key_by_unit(str(x)))
        sort_cols = ["_unit_key1"] + ([od] if od else []) + ["公文文號"]
        ascending = [True] + ([False] if od else []) + [True]
        df.sort_values(by=sort_cols, ascending=ascending, inplace=True, kind="mergesort")
        df.drop(columns=["_unit_key1"], inplace=True)
    elif od:
        df.sort_values(by=[od, "公文文號"], ascending=[False, True], inplace=True, kind="mergesort")
    return df

def filter_by_overdue(df: pd.DataFrame, min_days: int, inclusive: bool = False) -> pd.DataFrame:
    col = detect_overdue_column(df)
    if not col: return df
    s = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df[s >= min_days] if inclusive else df[s > min_days]

def reindex_serial(df: pd.DataFrame) -> pd.DataFrame:
    if "序號" in df.columns: df = df.drop(columns=["序號"])
    df.insert(0, "序號", range(1, len(df) + 1))
    return df

def add_three_header_rows(df: pd.DataFrame, type_key: str) -> pd.DataFrame:
    cols = list(df.columns)
    t1_text = FIRST_TITLES.get(type_key, f"逾期案件清單")
    t2_text = second_title_row()
    t1 = pd.DataFrame([[t1_text] + [""] * (len(cols) - 1)], columns=cols)
    t2 = pd.DataFrame([[t2_text] + [""] * (len(cols) - 1)], columns=cols)
    header_row = pd.DataFrame([cols], columns=cols)
    return pd.concat([t1, t2, header_row, df], ignore_index=True)

def strip_unit(df: pd.DataFrame) -> pd.DataFrame:
    # v1.7.0：ODIS 的承辦單位右邊補空白（例：'文化局      '），會影響排序與統計
    if "承辦單位" in df.columns:
        df["承辦單位"] = normalize_str_series(df["承辦單位"])
    return df

def fill_org_if_blank(df: pd.DataFrame) -> pd.DataFrame:
    if "來文機關" in df.columns:
        df["來文機關"] = normalize_str_series(df["來文機關"]).replace("", "創簽稿")
    return df

# --- 各類別處理函式 (A~G) ---
def process_generic(df, key, min_overdue=None):
    drops = ["cUNIT_NAME_PRI", "速別", "主旨", "iGDOC_OVER_DAY", "cGDOC_EXTEND"]
    # 一般公文保留來源速別；其他類別維持既有欄位規則。
    if key != "A":
        drops.append("cGDOC_SPEED")
    df = drop_columns_by_label_ordered(df, drops)
    mapping = {
        "cUNIT_NAME": "承辦單位", "cUSER_NAME_PRI": "承辦人", "cGDOC_SUBJECT": "主旨",
        "cGDOC_ORGNAME": "來文機關", "dGDOC_LIMITDATE_EXT": "計入展期之限辦日期",
        "Textbox76": "逾期天數（扣除展期天數）", "cGDOC_CTRL": "列管字號",
        # v1.7.0：陳情案件（C）版面多一欄次性質，v1.6.8 沒改名，Excel 標題直接顯示 cSCODE_DATA。
        # 名稱「公文次性質」經使用者 115.09.21 核對 ODIS 畫面確認。
        "cSCODE_DATA": "公文次性質",
    }
    if key == "A":
        mapping["cGDOC_SPEED"] = "速別"
    df = rename_columns(df, mapping)
    df = strip_unit(df)
    df = fill_org_if_blank(df)
    df = sort_multi(df)
    if min_overdue is not None:
        df = filter_by_overdue(df, min_overdue, inclusive=False)
    return reindex_serial(df), FIRST_TITLES[key]

PROCESSORS = {
    "A": lambda df: process_generic(df, "A", 5),
    "C": lambda df: process_generic(df, "C", 5),
    "D": lambda df: process_generic(df, "D"),
    "E": lambda df: process_generic(df, "E"),
    "F": lambda df: process_generic(df, "F"),
    "G": lambda df: process_generic(df, "G"),
    "B": lambda df: process_B(df) # B 較特殊
}

def process_B(df):
    df = drop_columns_by_label_ordered(df, ["cUNIT_NAME_PRI", "主旨", "iGDOC_OVER_DAY", "cGDOC_EXTEND", "Textbox6", "Textbox76", "Textbox215"])
    df = rename_columns(df, {
        "cUNIT_NAME": "承辦單位", "cUSER_NAME_PRI": "承辦人", "cGDOC_SUBJECT": "主旨",
        "cSCODE_DATA": "公文次性質", "cGDOC_ORGNAME": "來文機關", "速別": "辦理期限",
        "dGDOC_LIMITDATE_EXT": "計入展期之限辦日期", "cGDOC_SUBT_DATE2": "通知補正日期(補正結束日期)",
        "cGDOC_SUBT_MEMO2": "補正情形", "iGDOC_SUBT_DAY": "補正天數", "Textbox211": "逾期天數 (扣除展期及補正天數)",
    })
    df = strip_unit(df)
    if "公文文號" in df.columns:
        df.insert(df.columns.get_loc("公文文號") + 1, "續辦文號", "")
    df = sort_multi(df)
    df = filter_by_overdue(df, 5, inclusive=False)
    return reindex_serial(df), FIRST_TITLES["B"]

# ================= v1.7.0 新增：H 列管案件（資料來源與 A～G 不同）=================
# 來源：報表擷取抓的兩個檔
#   H_列管案件閘清單.csv    列管案件閘網頁清單（列管編號、公文文號、案由、限辦日期、承辦單位(含科)、承辦人員…）
#   H_補充_單筆查詢.csv     取報表.js 篩出監審議經＋限辦日早於抓取日後，逐筆以收文號查 RCV_QRY_GDOC 補
#                           「收文號、收文日期、來文機關、主辦單位」（115.09.21 起）。
#                           取代先前的「5050 全部性質」與「9050 列管案件一覽表（偵察）」兩個補充檔——
#                           前者查不到已文結案件、後者只做到偵察沒有真的產出資料，見交接手冊「已結案」表。
# 規則（使用者 115.09.21 訂）：
#   1. 只收監察、審計、議會、經費（列管編號第 4 字：監／審／議／經；操作手冊（七））
#   2. 限辦日期早於抓取當天（當天不算）才列入
#   3. 逾限天數＝限辦日期隔天起算到抓取當天的工作天（依 ODIS 行事曆；比照人工週報每週 +5）
#   4. 與限期公文（F）相同公文文號者，從 F 刪除、留在列管
CTRL_FILE = "H_列管案件閘清單.csv"
CTRL_SUPP_FILE = "H_補充_單筆查詢.csv"
CTRL_KINDS = {"監", "審", "議", "經"}
HOLIDAY_REL = os.path.join("範本", "行事曆_非工作日.txt")
PENDING = "待確認"

def parse_roc(s) -> Optional[dt.date]:
    m = re.match(r"^\s*(\d{2,3})/(\d{1,2})/(\d{1,2})", str(s or ""))
    if not m:
        return None
    try:
        return dt.date(int(m.group(1)) + 1911, int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None

def roc_to_date(roc: str) -> dt.date:
    return dt.date(int(roc[:3]) + 1911, int(roc[3:5]), int(roc[5:7]))

def load_holidays(path: Optional[str] = None) -> Tuple[set, set]:
    """讀 ODIS 行事曆非工作日（格式：115/09:5,6,12|…）。回傳（非工作日集合, 有資料的年月集合）。"""
    path = path or next((p for p in (os.path.join(b, HOLIDAY_REL) for b in (_app_dir(), getattr(sys, "_MEIPASS", "")) if b)
                         if os.path.isfile(p)), "")
    days, months = set(), set()
    if not path:
        return days, months
    with open(path, encoding="utf-8-sig") as fh:
        txt = fh.read()
    for part in re.split(r"[|\n]", txt):
        m = re.match(r"^\s*(\d{2,3})/(\d{1,2}):([\d,\s]*)$", part.strip())
        if not m:
            continue
        y, mo = int(m.group(1)) + 1911, int(m.group(2))
        months.add((y, mo))
        for d in m.group(3).split(","):
            if d.strip():
                days.add(dt.date(y, mo, int(d)))
    return days, months

def workdays_after(start: dt.date, end: dt.date, holidays: set, months: set, warn=None) -> int:
    """start 隔天起到 end（含）之間的工作天數。行事曆沒有的月份只扣週末。"""
    n, d = 0, start + dt.timedelta(days=1)
    missing = set()
    while d <= end:
        if (d.year, d.month) in months:
            n += d not in holidays
        else:
            missing.add((d.year - 1911, d.month))
            n += d.weekday() < 5
        d += dt.timedelta(days=1)
    if missing and warn:
        warn("行事曆沒有 " + "、".join(f"{y}/{m:02d}" for y, m in sorted(missing)) + "，這段只扣週末，請更新範本\\行事曆_非工作日.txt")
    return n

def first_level_unit(name: str) -> str:
    """'農業處林務科' → '農業處'；認不出來就原樣。"""
    m = re.match(r"^(.+?(處|局|所|室|館|場|隊|院|中心))", str(name).strip())
    return m.group(1) if m else str(name).strip()

def process_H(ctrl_path: str, supp_path: Optional[str], cutoff: dt.date, log_fn=print,
              holidays: Optional[Tuple[set, set]] = None) -> pd.DataFrame:
    ctrl = read_csv_safely(ctrl_path)
    ctrl.columns = [re.sub(r"\s", "", str(c)) for c in ctrl.columns]
    need = ["列管編號", "公文文號", "案由", "限辦日期", "承辦單位", "承辦人員"]
    lack = [c for c in need if c not in ctrl.columns]
    if lack:
        raise ValueError(f"列管案件閘清單缺少欄位：{'、'.join(lack)}")
    for c in ctrl.columns:
        ctrl[c] = normalize_str_series(ctrl[c])
    total = len(ctrl)
    kind = ctrl["列管編號"].str.extract(r"^\d{2,3}(\D)")[0]
    ctrl = ctrl[kind.isin(CTRL_KINDS)]
    n_kind = len(ctrl)
    lim = ctrl["限辦日期"].map(parse_roc)
    bad = ctrl[lim.isna()]
    if len(bad):
        log_fn(f"  ⚠ 列管：{len(bad)} 筆限辦日期讀不懂，略過：{'、'.join(bad['列管編號'])}")
    ctrl = ctrl[lim.notna() & lim.map(lambda d: d is not None and d < cutoff)].copy()
    log_fn(f"  列管：清單 {total} 筆 → 監審議經 {n_kind} 筆 → 限辦日期早於 {cutoff.year - 1911}/{cutoff.month:02d}/{cutoff.day:02d} 的 {len(ctrl)} 筆")

    # 115.09.21 改版：補充檔改讀取報表.js 逐筆單筆查詢（RCV_QRY_GDOC）的結果，欄位是
    # 收文號／收文日期／來文機關／主辦單位／狀態；狀態不是 OK（查無資料、查詢失敗…）一律當作沒補到。
    supp = {}
    if supp_path and os.path.isfile(supp_path):
        sp = read_csv_safely(supp_path)
        cols = [str(c).strip() for c in sp.columns]
        sp.columns = cols
        if all(c in cols for c in ("收文號", "來文機關", "收文日期", "主辦單位", "狀態")):
            for rec in sp.to_dict("records"):
                supp[str(rec["收文號"]).strip()] = rec
        else:
            log_fn("  ⚠ 列管補充檔版面不對（缺 收文號／來文機關／收文日期／主辦單位／狀態），來文機關、收文日期寫「待確認」")
    else:
        log_fn(f"  ⚠ 找不到 {CTRL_SUPP_FILE}，來文機關、收文日期寫「待確認」")

    hol, months = holidays if holidays is not None else load_holidays()
    warned = []
    rows, pend = [], []
    clean = lambda v: "" if v is None or str(v).strip().lower() == "nan" else str(v).strip()
    for rec in ctrl.to_dict("records"):
        no = rec["公文文號"]
        s = supp.get(no)
        ok = bool(s) and clean(s.get("狀態")) == "OK"
        unit = first_level_unit(clean(s.get("主辦單位"))) if ok and clean(s.get("主辦單位")) else ""
        org = clean(s.get("來文機關")) if ok and clean(s.get("來文機關")) else PENDING
        recv_date = clean(s.get("收文日期")) if ok and clean(s.get("收文日期")) else PENDING
        if not ok:
            狀態 = clean(s.get("狀態")) if s else "沒有補充檔"
            pend.append(f"{no}（{狀態}）")
        rows.append({
            "承辦單位": unit or first_level_unit(rec["承辦單位"]),
            "承辦人": rec["承辦人員"],
            "主旨": rec["案由"],
            "來文機關": org,
            "公文文號": no,
            "列管編號": rec["列管編號"],
            "收創日期": recv_date,
            "計入展期之限辦日期": rec["限辦日期"],
            "逾期天數（工作天）": workdays_after(parse_roc(rec["限辦日期"]), cutoff, hol, months,
                                            warn=(lambda m: warned.append(m)) if not warned else None),
        })
    for m in warned[:1]:
        log_fn("  ⚠ " + m)
    if pend:
        log_fn(f"  ⚠ 列管：{len(pend)} 筆單筆查詢沒補到來文機關／收文日期，寫「待確認」：{'、'.join(pend)}")
    out = pd.DataFrame(rows, columns=["承辦單位", "承辦人", "主旨", "來文機關", "公文文號", "列管編號",
                                      "收創日期", "計入展期之限辦日期", "逾期天數（工作天）"])
    out = sort_multi(out)
    return reindex_serial(out)

def dedup_ctrl(dfs: Dict[str, pd.DataFrame], log_fn) -> Dict[str, pd.DataFrame]:
    """與限期公文（F）相同公文文號者，從 F 刪除、留在列管（H）。"""
    if "H" in dfs and "F" in dfs and len(dfs["H"]):
        h = set(normalize_str_series(dfs["H"]["公文文號"])) - {""}
        f = dfs["F"]
        f["公文文號"] = normalize_str_series(f["公文文號"])
        hit = sorted(set(f["公文文號"]) & h)
        dfs["F"] = reindex_serial(f[~f["公文文號"].isin(h)])
        log_fn(f"去重：從 F 移除 {len(hit)} 筆 (重複於列管 H)" + (f"：{'、'.join(hit)}" if hit else ""))
    return dfs

def ctrl_cutoff(ctrl_path: str) -> dt.date:
    """抓取當天：列管檔所在資料夾是民國日期（例：1150921）就用它，否則用今天。"""
    day = os.path.basename(os.path.dirname(os.path.abspath(ctrl_path)))
    if day.isdigit() and len(day) == 7:
        try:
            return roc_to_date(day)
        except ValueError:
            pass
    return dt.date.today()

# ================= 115.10.30 新增：人民申請-建管系統（BM，資料來源是建管系統，不是 ODIS）=================
# 來源：建管系統「報表列印 → 建築相關執照人民申請案件清單」，條件：掛號日期 起日留空、迄日＝執行當天；
#       未結案；勾「逾限」6 天以上；列印格式選 .ods → 下載一個 zip，內含「工作天」「日曆天」兩個 .ods。
# 規則（使用者 115.10.30 訂）：
#   1. 只用「工作天」那個檔（日曆天不用）。
#   2. 只取 13 欄（使用者截圖標黃的 A～I、K～M、T），值原樣搬：日期保留「115/07/08(三)」、承辦單位照寫。
#      系統欄名「原限辦期限」對範本「原限辦日期」，其餘同名。
#   3. 範本的「備註(請選取)」「流程」人工填，留空。
#   4. 不篩選，系統給幾筆就放幾筆。
#   5. 總表的建管件數一律加到「建設處」那一列（ods_report.BM_SUMMARY_UNIT）。
BM_FILE_KEYWORD = "建築相關執照人民申請案件清冊"
BM_WORKDAY = "工作天"
# 115.11.03：工作天口徑沒有案件時，建管系統只回傳單一「日曆天」.ods（不是 zip）。取報表.js 會在輸出夾留這個
# 標記檔（不含任何案件資料），轉檔看到它就當「建管 0 筆」填分頁，不再提醒「清空留給人工」。
BM_ZERO_MARKER = "建築相關執照人民申請案件清冊_工作天_0筆.txt"
BM_COLUMNS = [   # (範本欄名, 系統檔欄名)
    ("編號", "編號"), ("承辦單位", "承辦單位"), ("承辦人", "承辦人"), ("來文者", "來文者"),
    ("申請項目", "申請項目"), ("收文號", "收文號"), ("收文日期", "收文日期"), ("辦理期限", "辦理期限"),
    ("原限辦日期", "原限辦期限"), ("通知補正日期", "通知補正日期"), ("請申請人限期補正日期", "請申請人限期補正日期"),
    ("補正完成日期", "補正完成日期"), ("逾限天數", "逾限天數"),
]

def _is_bm_workday(name: str) -> bool:
    base = os.path.basename(name)
    return BM_FILE_KEYWORD in base and BM_WORKDAY in base and base.lower().endswith(".ods")

def find_bm_file(folder: str) -> Tuple[Optional[Tuple[str, Optional[str]]], List[str]]:
    """在資料夾裡找建管系統的「工作天」檔：直接放的 .ods，或系統下載的 .zip（裡面的工作天 .ods）。
    回傳（(檔案路徑, zip 內檔名或 None) 或 None, 說明訊息）。有好幾個時用檔名排序最後一個
    （系統檔名結尾是產製時間，例：_11510301619，排最後的就是最新的），並在說明裡寫出用了哪一個。"""
    cands: List[Tuple[str, str, Optional[str]]] = []      # (排序用名稱, 路徑, zip 內檔名)
    notes: List[str] = []
    for n in sorted(os.listdir(folder)):
        p = os.path.join(folder, n)
        if not os.path.isfile(p):
            continue
        if _is_bm_workday(n):
            cands.append((n, p, None))
        elif n.lower().endswith(".zip"):
            try:
                import zipfile
                with zipfile.ZipFile(p) as z:
                    for m in z.namelist():
                        if _is_bm_workday(m):
                            cands.append((os.path.basename(m), p, m))
            except Exception as e:
                notes.append(f"建管：{n} 打不開（{e}），略過")
    if not cands:
        return None, notes
    cands.sort(key=lambda x: x[0])
    if len(cands) > 1:
        notes.append(f"建管：找到 {len(cands)} 個工作天檔，用最新的「{cands[-1][0]}」")
    return (cands[-1][1], cands[-1][2]), notes

def process_BM(path: str, member: Optional[str] = None) -> pd.DataFrame:
    """讀建管系統「工作天」.ods（member 有值時從 zip 裡讀），回傳範本欄名的 13 欄資料表，值一律原樣文字。
    版面不對（找不到表頭、缺欄）就丟 ValueError，由呼叫端決定要不要清空分頁留給人工。"""
    import ods_report
    if member:
        import zipfile
        with zipfile.ZipFile(path) as z:
            rows = ods_report.read_ods_rows(z.read(member))
    else:
        rows = ods_report.read_ods_rows(path)
    norm = lambda v: re.sub(r"\s", "", str(v or ""))
    h = next((i for i, r in enumerate(rows[:20]) if r and norm(r[0]) == "編號" and "收文號" in [norm(c) for c in r]), None)
    if h is None:
        raise ValueError("找不到「編號…收文號」表頭列，可能不是「建築相關執照人民申請案件清冊」")
    header = [norm(c) for c in rows[h]]
    idx = {}
    lack = []
    for tpl, src in BM_COLUMNS:
        if src in header:
            idx[tpl] = header.index(src)
        else:
            lack.append(src)
    if lack:
        raise ValueError(f"系統檔缺少欄位：{'、'.join(lack)}")
    recs = []
    for r in rows[h + 1:]:
        vals = {tpl: (r[i].strip() if i < len(r) else "") for tpl, i in idx.items()}
        if any(vals.values()):
            recs.append(vals)
    return pd.DataFrame(recs, columns=[t for t, _ in BM_COLUMNS])

# ================= 去重與統計 =================

def dedup_part3(dfs: Dict[str, pd.DataFrame], log_fn) -> Dict[str, pd.DataFrame]:
    # v1.7.0 修正：ODIS 匯出的公文文號右邊補空白（例：'1150000001   '）。
    # v1.6.8 找重複時有去空白、刪除時沒有，所以一筆都刪不掉。這裡先統一去空白，找與刪用同一份值。
    for k in dfs:
        if "公文文號" in dfs[k].columns:
            dfs[k]["公文文號"] = normalize_str_series(dfs[k]["公文文號"])

    def get_set(k):
        return set(dfs[k]["公文文號"]) - {""} if k in dfs and "公文文號" in dfs[k].columns else set()

    def drop(k, remove, reason):
        before = len(dfs[k])
        dfs[k] = reindex_serial(dfs[k][~dfs[k]["公文文號"].isin(remove)])
        removed = before - len(dfs[k])   # 紀錄寫「實際刪掉」的筆數，不是「找到」的筆數
        log_fn(f"去重：從 {k} 移除 {removed} 筆 ({reason})" + (f"：{'、'.join(sorted(remove))}" if removed else ""))

    A_set, B_set, F_set, G_set = get_set("A"), get_set("B"), get_set("F"), get_set("G")

    if A_set and (F_set | G_set):
        drop("A", (A_set & F_set) | (A_set & G_set), "重複於 F/G")

    if B_set:
        for k in ["F", "G"]:
            if k in dfs:
                drop(k, get_set(k) & B_set, "重複於 B")
    return dfs

def build_part4_summary(dfs):
    PART4_COLS = [("A", "A_一般"), ("B", "B_人民申請"), ("C", "C_陳情"), ("D", "D_訴願"), ("E", "E_專案"), ("F", "F_限期"), ("G", "G_特殊性"), ("H", "H_列管")]
    units = set()
    for df in dfs.values():
        if "承辦單位" in df.columns: units |= set(df["承辦單位"].dropna())
    
    if not units: return pd.DataFrame()
    
    res = pd.DataFrame({"承辦單位": sorted(list(units), key=sort_key_by_unit)})
    for k, cn in PART4_COLS:
        if k in dfs:
            counts = dfs[k]["承辦單位"].value_counts()
            res[cn] = res["承辦單位"].map(lambda u: int(counts.get(u, 0)))
        else:
            res[cn] = 0
            
    res = res[res.iloc[:, 1:].sum(axis=1) > 0].copy()
    total = res.iloc[:, 1:].sum().to_dict(); total["承辦單位"] = "合計"
    res = pd.concat([res, pd.DataFrame([total])], ignore_index=True)
    
    rename_map = {"承辦單位": "機關(單位)", "A_一般": "一般公文(件數)", "B_人民申請": "人民申請(件數)", "C_陳情": "陳情案件(件數)", "D_訴願": "訴願案件(件數)", "E_專案": "專案管制案件(件數)", "F_限期": "限期公文案件(件數)", "G_特殊性": "特殊性案件(件數)", "H_列管": "列管案件(件數)"}
    return res.rename(columns=rename_map)

# ================= Excel 樣式匯出 =================

def save_to_single_excel(sheets, outdir, roc, log_fn):
    xlsx_path = os.path.join(outdir, f"{roc}_公文清單_彙整.xlsx")
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    side = Side(style="thin", color="000000")
    border = Border(left=side, right=side, top=side, bottom=side)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for k, df in sheets.items():
            sname = k if k != "PART4" else "統計表"
            df.to_excel(writer, sheet_name=sname[:31], index=False, header=False)
            ws = writer.book[sname[:31]]
            max_col = ws.max_column
            
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                    cell.border = border
            
            # 前三列樣式
            for r in range(1, 4):
                if r > ws.max_row: break
                for cell in ws[r]:
                    cell.font = Font(bold=True)
                    cell.fill = header_fill
                if r == 1: 
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
                    ws.cell(1,1).alignment = Alignment(horizontal="center")
                elif r == 2:
                    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)

            for col_idx in range(1, max_col + 1):
                ws.column_dimensions[get_column_letter(col_idx)].width = 20

    log_fn(f"Excel 匯出完成：{xlsx_path}")
    return xlsx_path

# ================= v1.7.0 新增：從資料夾自動帶入 =================

# 每一類一定要有的欄位（沒有就代表不是正確的報表版面）
FETCH_CHECK_COL = {k: "Textbox76" for k in "ACDEFG"}
FETCH_CHECK_COL["B"] = "Textbox211"
FETCH_TOOL_DIR = "報表擷取_前段七類"

def _app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def default_fetch_dir(base: Optional[str] = None) -> str:
    r"""找「報表擷取_前段七類／輸出」裡日期最新的資料夾；找不到回傳空字串。

    v1.7.0 打包成 exe 後才發現的坑：sys.executable 指到的是 exe 檔本身所在的資料夾，
    PyInstaller onefile 預設打包目錄是 dist\，比原始碼 OverdueList_v1.7.0.py 所在的
    v1.7.0_自動帶入\ 深一層；原本只往上找一層（base/..），打包後的 exe 因此永遠找不到
    v1.7.0\報表擷取_前段七類\輸出（用原始碼 python.exe 跑不會踩到，因為 __file__ 就在
    v1.7.0_自動帶入\，往上一層剛好對）。改成同時往上找兩層，兩種執行模式都涵蓋，不用
    分別判斷是否為 frozen。
    """
    base = base or _app_dir()
    candidates = (
        os.path.join(base, FETCH_TOOL_DIR, "輸出"),
        os.path.join(base, "..", FETCH_TOOL_DIR, "輸出"),
        os.path.join(base, "..", "..", FETCH_TOOL_DIR, "輸出"),
    )
    for c in candidates:
        if os.path.isdir(c):
            days = sorted(d for d in os.listdir(c) if d.isdigit() and len(d) == 7 and os.path.isdir(os.path.join(c, d)))
            if days:
                return os.path.normpath(os.path.join(c, days[-1]))
    return ""

def find_category_files(folder: str) -> Tuple[Dict[str, str], List[str]]:
    """在資料夾裡找 A_*.csv～G_*.csv，並檢查版面。回傳（找到的檔案, 問題清單）。"""
    found: Dict[str, str] = {}
    problems: List[str] = []
    names = os.listdir(folder)
    for k in "ABCDEFG":
        hits = sorted(n for n in names if n.upper().startswith(k + "_") and n.lower().endswith(".csv"))
        if not hits:
            problems.append(f"{k} {NAME_MAP[k]}：資料夾裡沒有 {k}_*.csv")
            continue
        if len(hits) > 1:
            problems.append(f"{k} {NAME_MAP[k]}：有 {len(hits)} 個符合的檔案（{'、'.join(hits)}），請只留一個")
            continue
        path = os.path.join(folder, hits[0])
        try:
            cols = [str(c).strip() for c in read_csv_safely(path).columns]
        except Exception as e:
            problems.append(f"{k} {NAME_MAP[k]}：讀檔失敗（{e}）")
            continue
        need = FETCH_CHECK_COL[k]
        if need not in cols:
            problems.append(f"{k} {NAME_MAP[k]}：缺少「{need}」欄，可能不是正確的報表版面")
            continue
        found[k] = path
    return found, problems

def leftover_code_columns(df: pd.DataFrame) -> List[str]:
    """處理後仍是系統代碼的欄名（例：cSCODE_DATA、iGDOC_OVER_DAY、Textbox76），代表對照表漏了。"""
    import re
    return [c for c in df.columns if re.match(r"^(c|i|d)[A-Z][A-Z0-9_]*$|^Textbox\d+$", str(c))]

TEMPLATE_REL = os.path.join("範本", "公文逾限簡報_範本.ods")

def find_template(base: Optional[str] = None) -> str:
    """範本放在程式旁的「範本」資料夾；打包成 exe 時也會帶一份在 exe 裡。"""
    for b in (base or _app_dir(), getattr(sys, "_MEIPASS", "")):
        if b:
            p = os.path.join(b, TEMPLATE_REL)
            if os.path.isfile(p):
                return p
    return ""

def unique_path(path: str) -> str:
    """同名檔已存在（可能正開著）就加 (2)、(3)…，不覆蓋。"""
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{root}({i}){ext}"):
        i += 1
    return f"{root}({i}){ext}"

def folder_date_warning(folder: str, today: Optional[str] = None) -> str:
    """資料夾名稱是民國日期（例：1150921）且不是今天時，回傳提醒文字。"""
    today = today or roc_today_str()
    day = os.path.basename(os.path.normpath(folder))
    if day.isdigit() and len(day) == 7 and day != today:
        return (f"這批報表是 {day} 抓的，但轉出的報表標題會寫「截至今天（{today}）」。\n"
                f"逾期天數會跟著抓取日變，建議當天重抓再轉。")
    return ""

def _dedup_warnings(warnings: List[str]) -> List[str]:
    """把 run_headless() 過程中蒐集到的 ⚠／待確認 訊息去掉開頭的「⚠」符號跟前後空白，
    再依出現順序去重。_format_summary()（文字版摘要框）跟 run_headless() 的 summary_sink
    （給 GUI 用的結構化版）共用同一份去重結果，不要各自重寫一次維護兩套。"""
    return list(dict.fromkeys(w.strip().lstrip("⚠").strip() for w in warnings if w.strip()))


def _format_summary(dfs: Dict[str, pd.DataFrame], out: str, warnings: List[str]) -> List[str]:
    """115.09.23 新增：收工摘要（UI 改善第一優先第 3 點）。
    跑完整段流程的細節已經透過 log_fn 逐行印出、也全部進了 執行紀錄.txt／console，
    但同仁真正想先看的只有兩件事：各類最終筆數、有沒有需要人工確認的項目。這裡把
    run_headless() 過程中蒐集到的 ⚠／待確認 訊息去重後，跟最終筆數一起組成收工時印的摘要框。"""
    lines = ["", "=" * 42, f"✅ 完成：{os.path.basename(out)}"]
    counts = "｜".join(f"{NAME_MAP.get(k, k)} {len(dfs[k])}" for k in SUMMARY_ORDER if k in dfs)
    if counts:
        lines.append("   " + counts)
    uniq_warnings = _dedup_warnings(warnings)
    if uniq_warnings:
        lines.append(f"⚠ 需要人工確認 {len(uniq_warnings)} 項：")
        lines.extend(f"   ・{w}" for w in uniq_warnings)
    else:
        lines.append("（沒有需要人工確認的項目）")
    lines.append("=" * 42)
    return lines


def write_output_marker(out: str) -> None:
    """115.09.23 新增：把最後產出的檔案完整路徑寫到輸出資料夾裡的固定檔名 `_上次輸出.txt`，
    給 一鍵執行.bat 收工時讀出來，直接打開這個 .ods（不用同仁自己在資料夾裡找哪個是最新的，
    尤其 unique_path() 同名時會自動加 (2)(3)…，光看資料夾不一定分得出哪個是這次剛產出的）。
    用 cp950（Windows 繁中內碼頁）存檔，跟 .bat 本身的編碼、雙擊執行時的主控台內碼頁一致——
    這裡如果貪方便用 UTF-8 存，.bat 用 `set /p` 讀回來會依目前主控台內碼頁（950）誤解讀位元組，
    路徑會變亂碼、`start` 打不開檔案，是 CRLF 換行符雷的同一類問題（.bat 這一層的編碼要單獨
    處理，不能假設 Python 端的 UTF-8 直接搬過去就對）。寫檔失敗（例如檔名含 cp950 沒有的字、
    輸出資料夾沒有寫入權限）不影響轉檔本身，只是 .bat 收工時會退回打開資料夾，不是致命錯誤。
    """
    try:
        marker = os.path.join(os.path.dirname(out), "_上次輸出.txt")
        with open(marker, "w", encoding="cp950", errors="replace") as f:
            f.write(out)
    except Exception:
        pass


def run_headless(folder: Optional[str] = None, outdir: Optional[str] = None, log_fn=print,
                  summary_sink: Optional[dict] = None) -> str:
    """115.09.21 新增：命令列一次跑完「從資料夾自動帶入」＋「開始轉檔」，不開視窗。
    給 一鍵執行.bat／CLI --auto 用。邏輯跟 App._auto_load()／App._run() 完全一樣，
    只是資料來源不是 tkinter 元件，是參數／函式回傳值——維持同一套判斷規則，
    避免 GUI 版跟命令列版兩邊各自維護、久了行為兜不起來。
    folder：抓好報表的資料夾，預設用 default_fetch_dir() 抓「報表擷取_前段七類\輸出」裡最新一天。
    outdir：.ods 輸出資料夾，預設放使用者家目錄（跟 GUI 的預設值一致）。
    回傳最後產出的檔案完整路徑；任何一步出錯就丟例外，讓呼叫端（.bat）看得到非 0 結束碼。

    115.09.23 新增（UI 改善第一優先）：過程中所有含「⚠」或「待確認」的訊息會另外收集起來，
    結束前用 log_fn 印一個收工摘要框（各類最終筆數＋需要人工確認的項目去重清單），並呼叫
    write_output_marker() 把產出路徑寫給 .bat 用來自動打開檔案。這兩個都是收工才做的事，
    不影響過程中逐行的 log_fn 輸出，GUI 版（App._run()，走的是自己另一份邏輯，不是呼叫這裡）
    暫時沒有這個摘要，之後如果要讓 GUI 也有，再考慮把兩邊真的合併成同一份呼叫。

    115.09.23 新增（UI 改善第三優先，一鍵執行_GUI.py 用）：summary_sink 給呼叫端一個字典，
    結束前會被填入 {'out', 'counts', 'warnings'}（跟印給 log_fn 的摘要框是同一份資料，只是
    給程式讀的結構化版本，不用去 parse 文字訊息）。預設 None 不影響任何既有呼叫端
    （CLI --auto、既有測試），是純新增的可選參數。
    """
    warnings_collected: List[str] = []

    def _log(msg):
        if isinstance(msg, str) and ("⚠" in msg or "待確認" in msg):
            warnings_collected.append(msg)
        log_fn(msg)

    folder = folder or default_fetch_dir()
    if not folder:
        raise RuntimeError("找不到抓好的報表資料夾（報表擷取_前段七類\輸出 底下沒有日期資料夾），"
                            "請確認報表擷取有先跑成功，或用 --folder 指定資料夾")
    if not os.path.isdir(folder):
        raise RuntimeError(f"資料夾不存在：{folder}")
    outdir = outdir or os.path.expanduser("~")

    found, problems = find_category_files(folder)
    hp = os.path.join(folder, CTRL_FILE)
    if os.path.isfile(hp):
        found["H"] = hp
    else:
        problems.append(f"H 列管案件：資料夾裡沒有 {CTRL_FILE}（列管頁會清空留給人工）")
    if os.path.isfile(hp) and not os.path.isfile(os.path.join(folder, CTRL_SUPP_FILE)):
        problems.append(f"H 列管案件：沒有 {CTRL_SUPP_FILE}，來文機關、收文日期會寫「待確認」")
    bm, bm_notes = find_bm_file(folder)
    bm_zero = bm is None and os.path.isfile(os.path.join(folder, BM_ZERO_MARKER))
    if bm is None and not bm_zero:
        problems.append(f"建管系統：資料夾裡沒有「{BM_FILE_KEYWORD}_{BM_WORKDAY}」.ods（或系統下載的 .zip），"
                        "建管分頁會清空留給人工")
    _log(f"從資料夾帶入：{folder}（找到 {len(found) + (bm is not None)}／9 類，含 H、建管）")
    for p in problems:
        _log("  ⚠ " + p)
    for n in bm_notes:
        _log("  " + ("⚠ " if "打不開" in n else "") + n)
    w = folder_date_warning(folder)
    if w:
        _log("  ⚠ " + w)

    dfs = {}
    for k in "ABCDEFGH":
        if k == "H":
            if "H" in found:
                dfs["H"] = process_H(found["H"], os.path.join(os.path.dirname(found["H"]), CTRL_SUPP_FILE),
                                      ctrl_cutoff(found["H"]), _log)
                _log(f"讀取 H 成功（{len(dfs['H'])} 筆）")
            continue
        if k in found:
            df_raw = read_csv_safely(found[k])
            dfs[k], _ = PROCESSORS[k](df_raw)
            _log(f"讀取 {k} 成功")
            left = leftover_code_columns(dfs[k])
            if left:
                _log(f"  ⚠ {k} 有欄位標題沒改成中文：{'、'.join(left)}（程式對照表漏了，請回報）")

    if bm_zero:
        dfs["BM"] = pd.DataFrame(columns=[t for t, _ in BM_COLUMNS])
        _log("讀取 建管 成功（0 筆）：工作天沒有符合條件的案件（建管系統只回傳日曆天檔）")
    if bm is not None:
        try:
            dfs["BM"] = process_BM(*bm)
            _log(f"讀取 建管 成功（{len(dfs['BM'])} 筆）：{os.path.basename(bm[1] or bm[0])}")
        except Exception as e:
            _log(f"  ⚠ 建管系統：讀檔失敗（{e}），建管分頁會清空留給人工")

    if not dfs:
        raise RuntimeError("沒有任何一類讀到資料，請確認資料夾內容")

    dfs = dedup_part3(dfs, _log)
    dfs = dedup_ctrl(dfs, _log)
    excel_data = {}
    for k, df in dfs.items():
        excel_data[k] = add_three_header_rows(df, k)

    p4 = build_part4_summary(dfs)
    if not p4.empty:
        p4_final = pd.concat([
            pd.DataFrame([["公文逾限一覽表"] + [""] * (len(p4.columns) - 1)]),
            pd.DataFrame([[second_title_row()] + [""] * (len(p4.columns) - 1)]),
            pd.DataFrame([p4.columns]), p4,
        ], ignore_index=True)
        excel_data["PART4"] = p4_final

    roc = roc_today_str()
    template = find_template()
    if template:
        import ods_report
        y, m, d = roc[:3], roc[3:5], roc[5:7]
        out = unique_path(os.path.join(outdir, f"{y}.{m}.{d}_公文逾限簡報_-_未加陳情.ods"))
        ods_report.build_report(template, dfs, out, roc, sort_key_by_unit, _log)
        _log(f"週報匯出完成：{out}")
    else:
        _log(f"⚠ 找不到範本（{TEMPLATE_REL}），改用舊版 Excel 格式輸出")
        out = save_to_single_excel(excel_data, outdir, roc, _log)

    for line in _format_summary(dfs, out, warnings_collected):
        log_fn(line)
    write_output_marker(out)
    if summary_sink is not None:
        summary_sink["out"] = out
        summary_sink["counts"] = {k: len(dfs[k]) for k in SUMMARY_ORDER if k in dfs}
        summary_sink["warnings"] = _dedup_warnings(warnings_collected)
    return out


# ================= GUI 介面 =================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("公文轉檔小助手 v1.7.1")
        self.geometry("1000x600")
        self.var_path = {k: tk.StringVar() for k in "ABCDEFGH"}
        self.var_enable = {k: tk.BooleanVar(value=True) for k in "ABCDEFGH"}
        self.var_outdir = tk.StringVar(value=os.path.expanduser("~"))
        self._setup_ui()

    def _setup_ui(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # 檔案選取區
        list_frame = ttk.LabelFrame(main_frame, text="清單設定", padding=5)
        list_frame.pack(fill="x", pady=5)
        for i, k in enumerate("ABCDEFGH"):
            ttk.Checkbutton(list_frame, variable=self.var_enable[k]).grid(row=i, column=0)
            ttk.Label(list_frame, text=f"{k} {NAME_MAP[k]}").grid(row=i, column=1, padx=5)
            ttk.Entry(list_frame, textvariable=self.var_path[k], width=60).grid(row=i, column=2, padx=5)
            ttk.Button(list_frame, text="瀏覽", command=lambda x=k: self._browse(x)).grid(row=i, column=3)

        # 輸出設定
        out_frame = ttk.Frame(main_frame, padding=5)
        out_frame.pack(fill="x")
        ttk.Label(out_frame, text="輸出資料夾:").pack(side="left")
        ttk.Entry(out_frame, textvariable=self.var_outdir, width=50).pack(side="left", padx=5)
        ttk.Button(out_frame, text="選擇", command=self._browse_dir).pack(side="left")

        # 按鈕與日誌
        btn_frame = ttk.Frame(main_frame, padding=5)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="從資料夾自動帶入", command=self._auto_load).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="開始轉檔", command=self._run).pack(side="left", padx=5)
        
        self.logbox = tk.Text(main_frame, height=10, font=("微軟正黑體", 9))
        self.logbox.pack(fill="both", expand=True, pady=5)

    def log(self, msg):
        self.logbox.insert("end", f"[{dt.datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.logbox.see("end")
        self.update()

    def _browse(self, k):
        p = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if p: self.var_path[k].set(p)

    def _browse_dir(self):
        p = filedialog.askdirectory()
        if p: self.var_outdir.set(p)

    def _auto_load(self):
        init = default_fetch_dir() or self.var_outdir.get()
        folder = filedialog.askdirectory(initialdir=init, title="選擇抓好的報表資料夾（例：輸出\\1150921）")
        if not folder:
            return
        found, problems = find_category_files(folder)
        for k in "ABCDEFG":
            self.var_path[k].set(found.get(k, ""))
            self.var_enable[k].set(k in found)
        hp = os.path.join(folder, CTRL_FILE)
        self.var_path["H"].set(hp if os.path.isfile(hp) else "")
        self.var_enable["H"].set(os.path.isfile(hp))
        if not os.path.isfile(hp):
            problems.append(f"H 列管案件：資料夾裡沒有 {CTRL_FILE}（列管頁會清空留給人工）")
        elif not os.path.isfile(os.path.join(folder, CTRL_SUPP_FILE)):
            problems.append(f"H 列管案件：沒有 {CTRL_SUPP_FILE}，來文機關、收文日期會寫「待確認」")
        self.log(f"從資料夾帶入：{folder}（找到 {len(found)}／7 類）")
        for p in problems:
            self.log("  ⚠ " + p)
        warns = []
        if problems:
            warns.append("有 %d 類沒有帶入：\n" % len(problems) + "\n".join(problems))
        w = folder_date_warning(folder)
        if w:
            warns.append(w)
        if warns:
            messagebox.showwarning("請確認", "\n\n".join(warns))

    def _run(self):
        dfs = {}
        outdir = self.var_outdir.get()
        roc = roc_today_str()
        
        try:
            for k in "ABCDEFGH":
                if k == "H":
                    hp = self.var_path["H"].get()
                    if self.var_enable["H"].get() and hp:
                        dfs["H"] = process_H(hp, os.path.join(os.path.dirname(hp), CTRL_SUPP_FILE), ctrl_cutoff(hp), self.log)
                        self.log(f"讀取 H 成功（{len(dfs['H'])} 筆）")
                    continue
                if self.var_enable[k].get() and self.var_path[k].get():
                    df_raw = read_csv_safely(self.var_path[k].get())
                    dfs[k], _ = PROCESSORS[k](df_raw)
                    self.log(f"讀取 {k} 成功")
                    left = leftover_code_columns(dfs[k])
                    if left:
                        self.log(f"  ⚠ {k} 有欄位標題沒改成中文：{'、'.join(left)}（程式對照表漏了，請回報）")
            
            if not dfs:
                messagebox.showwarning("提示", "未選取任何檔案")
                return

            dfs = dedup_part3(dfs, self.log)
            dfs = dedup_ctrl(dfs, self.log)
            excel_data = {}
            
            for k, df in dfs.items():
                final_df = add_three_header_rows(df, k)
                excel_data[k] = final_df
                
            p4 = build_part4_summary(dfs)
            if not p4.empty:
                # 這裡簡單處理統計表表頭
                p4_final = pd.concat([pd.DataFrame([["公文逾限一覽表"] + [""]*(len(p4.columns)-1)]), pd.DataFrame([[second_title_row()] + [""]*(len(p4.columns)-1)]), pd.DataFrame([p4.columns]), p4], ignore_index=True)
                excel_data["PART4"] = p4_final
                
            template = find_template()
            if template:
                import ods_report
                y, m, d = roc[:3], roc[3:5], roc[5:7]
                out = unique_path(os.path.join(outdir, f"{y}.{m}.{d}_公文逾限簡報_-_未加陳情.ods"))
                ods_report.build_report(template, dfs, out, roc, sort_key_by_unit, self.log)
                self.log(f"週報匯出完成：{out}")
            else:
                self.log(f"⚠ 找不到範本（{TEMPLATE_REL}），改用舊版 Excel 格式輸出")
                save_to_single_excel(excel_data, outdir, roc, self.log)
            messagebox.showinfo("完成", "轉檔已結束！")
            
        except Exception as e:
            self.log(f"發生錯誤: {str(e)}")
            messagebox.showerror("錯誤", str(e))

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="公文轉檔小助手 v1.7.1")
    ap.add_argument("--auto", action="store_true",
                     help="不開視窗，直接從資料夾自動帶入＋轉檔（給 一鍵執行.bat／自動化流程用）")
    ap.add_argument("--folder", default=None, help="抓好報表的資料夾；預設抓最新一天（同 GUI「從資料夾自動帶入」預設值）")
    ap.add_argument("--outdir", default=None, help=".ods 輸出資料夾；預設放使用者家目錄（同 GUI 預設值）")
    args = ap.parse_args()
    if args.auto:
        try:
            out = run_headless(args.folder, args.outdir)
            print(f"轉檔完成：{out}")
            sys.exit(0)
        except Exception as e:
            print(f"轉檔失敗：{e}")
            sys.exit(1)
    else:
        App().mainloop()


