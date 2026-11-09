# -*- coding: utf-8 -*-
"""
ods_report.py — 依「人工週報範本（.ods）」的版面產出前段週報（v1.7.0 新增）

做法：打開範本 .ods，保留封面、樣式、欄寬、隱藏欄、下拉選單、篩選、列印設定與面報說明文字，
只換掉各分頁的資料列、分頁名稱的件數、「截至○年○月○日」與總表件數。
範本的第一筆資料列當「樣板列」，新資料列都複製它的格式。

只用 Python 標準函式庫（zipfile + xml.etree），打包 exe 不需額外套件。
"""
import copy
import math
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Callable, Dict, List, Optional

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
}
T = "{%s}" % NS["table"]
O = "{%s}" % NS["office"]
X = "{%s}" % NS["text"]
S = "{%s}" % NS["style"]
FO = "{%s}" % NS["fo"]
CX = "{%s}" % NS["calcext"]

# 範本欄名 → 小助手處理後的欄名（None＝人工欄，留空）
COLUMN_SOURCE = {
    "編號": "序號", "承辦單位": "承辦單位", "承辦人": "承辦人", "主旨": "主旨",
    "來文機關": "來文機關", "收文號": "公文文號", "收文日期": "收創日期", "速別": "速別",
    "限辦日期": "計入展期之限辦日期", "逾限天數": "__逾期天數__",
    "申請項目": "公文次性質", "續辦文號": "續辦文號", "辦理期限": "辦理期限",
    "通知補正日期(補正結束日期)": "通知補正日期(補正結束日期)", "補正情形": "補正情形",
    "補正天數(補正結束者)": "補正天數", "列管編號": "列管編號",
}
# 115.10.30 新增：建管系統（BM）分頁的欄名跟系統檔「建築相關執照人民申請案件清冊_工作天」一樣，
# 由 OverdueList_v1.7.1.py 的 process_BM() 先把系統欄名換成範本欄名，這裡一律同名對應；
# 「備註(請選取)」「流程」使用者人工填（115.10.30 使用者確認），固定留空。
COLUMN_SOURCE_BY_KEY = {
    "BM": {h: h for h in ("編號", "承辦單位", "承辦人", "來文者", "申請項目", "收文號", "收文日期", "辦理期限",
                          "原限辦日期", "通知補正日期", "請申請人限期補正日期", "補正完成日期", "逾限天數")},
}
COLUMN_SOURCE_BY_KEY["BM"].update({"備註(請選取)": None, "流程": None})
CONSTANT = {"C": {"陳情管道": "公文"}}
NUMERIC = {"編號", "逾限天數", "辦理期限", "補正天數(補正結束者)"}

# 分頁（名稱開頭）→ 類別
SHEET_PREFIX = [("一般公文(", "A"), ("人民申請(", "B"), ("特殊性案件(", "G"), ("限期公文(", "F"), ("陳情案件(", "C"),
                ("列管案件(", "H"), ("人民申請-建管系統(", "BM")]
# 非 ODIS 來源、程式不填的分頁：清空留給人工。115.10.30 起「人民申請-建管系統」改由建管系統的
# 「工作天」報表自動填（上面 SHEET_PREFIX 的 BM），沒有該檔時才清空，跟列管案件（H）沒有資料時一樣。
CLEAR_SHEETS: List[str] = []
OPTIONAL_KEYS = {"H", "BM"}       # 沒有資料時清空該分頁、留給人工，不填 0 筆
# 115.10.30 使用者確認：建管分頁的承辦單位照系統原樣寫（建管科、公會協審…），但總表一律把建管件數
# 加到「建設處」那一列；那週建設處沒有其他公文逾期時，總表就新增一列「建設處」。
BM_SUMMARY_UNIT = "建設處"
NEW_SHEETS = {"D": "訴願案件", "E": "專案案件"}          # 範本沒有，有案件才新增（複製特殊性案件頁）
SUMMARY_PREFIX = "本府及所屬一、二級機關公文逾限一覽表"
SUMMARY_KEYS = [("建管", "BM"), ("列管", "H"), ("一般", "A"), ("人民申請", "B"), ("特殊性", "G"),
                ("限期", "F"), ("陳情", "C"), ("訴願", "D"), ("專案", "E")]

CJK_CM = 0.52     # 14pt 標楷體一個中文字約 0.52cm
LINE_CM = 0.62


def _norm(s) -> str:
    return re.sub(r"\s+", "", str(s or ""))


def cell_text(cell) -> str:
    out = []
    for p in cell.iter(X + "p"):
        parts = []
        def walk(e):
            if e.text: parts.append(e.text)
            for ch in e:
                if ch.tag == X + "s":
                    parts.append(" " * int(ch.get(X + "c", "1")))
                elif ch.tag == X + "line-break":
                    parts.append("\n")
                else:
                    walk(ch)
                if ch.tail: parts.append(ch.tail)
        walk(p)
        out.append("".join(parts))
    return "\n".join(out)


def is_cell(e) -> bool:
    return e.tag in (T + "table-cell", T + "covered-table-cell")


def rep(e, attr=T + "number-columns-repeated") -> int:
    return int(e.get(attr, "1"))


def split_cells(row, upto: int) -> List[ET.Element]:
    """讓前 upto 格都是獨立元素（拆開 number-columns-repeated），回傳這些格。"""
    col, i = 0, 0
    kids = list(row)
    while col < upto and i < len(kids):
        e = kids[i]
        if not is_cell(e):
            i += 1; continue
        r = rep(e)
        if r > 1:
            e.attrib.pop(T + "number-columns-repeated")
            pos = list(row).index(e)
            need = min(r, upto - col) - 1
            for k in range(need):
                row.insert(pos + 1 + k, copy.deepcopy(e))
            if r - 1 - need > 0:
                rest = copy.deepcopy(e)
                if r - 1 - need > 1:
                    rest.set(T + "number-columns-repeated", str(r - 1 - need))
                row.insert(pos + 1 + need, rest)
            kids = list(row)
        col += 1; i += 1
    return [e for e in row if is_cell(e)][:upto]


def set_cell(cell, value, numeric: bool):
    for a in list(cell.attrib):
        if a.startswith(O) or a.startswith(CX) or a == T + "formula":
            cell.attrib.pop(a)
    for ch in list(cell):
        cell.remove(ch)
    if value is None:
        return
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return
    if numeric:
        try:
            f = float(s)
            v = str(int(f)) if f == int(f) else str(f)
            cell.set(O + "value-type", "float"); cell.set(O + "value", v); cell.set(CX + "value-type", "float")
            ET.SubElement(cell, X + "p").text = v
            return
        except ValueError:
            pass
    cell.set(O + "value-type", "string"); cell.set(CX + "value-type", "string")
    for line in s.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        p = ET.SubElement(cell, X + "p")
        _put_text(p, line)


def _put_text(p, line: str):
    """連續空白要寫成 <text:s text:c="n"/>，否則會被吃掉。"""
    parts = re.split(r"( {2,})", line)
    p.text = parts[0]
    last = None
    for k in range(1, len(parts), 2):
        if last is None:
            p.text = (p.text or "") + " "
        else:
            last.tail = (last.tail or "") + " "
        sp = ET.SubElement(p, X + "s")
        if len(parts[k]) - 1 > 1:
            sp.set(X + "c", str(len(parts[k]) - 1))
        sp.tail = parts[k + 1]
        last = sp


def read_ods_rows(src, max_rows: int = 5000) -> List[List[str]]:
    """115.10.30 新增：讀 .ods 第一個工作表，回傳每列的儲存格文字（只用標準函式庫，打包不需額外套件）。
    src 可以是檔案路徑或 bytes。合併儲存格被蓋住的格（covered-table-cell）也算一格，欄位位置才對得上；
    大量重複的空白列／欄（整張表的尾巴）不展開。"""
    import io
    if isinstance(src, (bytes, bytearray)):
        data = src
    else:
        with open(src, "rb") as fh:
            data = fh.read()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        doc = ET.fromstring(z.read("content.xml"))
    table = doc.find(f"{O}body/{O}spreadsheet/{T}table")
    if table is None:
        return []
    out: List[List[str]] = []
    for r in table.iter(T + "table-row"):
        cells: List[str] = []
        for c in r:
            if is_cell(c):
                cells.extend([cell_text(c)] * min(rep(c), 200))
        while cells and not cells[-1].strip():
            cells.pop()
        n = rep(r, T + "number-rows-repeated")
        if not cells and n > 1:
            continue
        out.extend([cells] * min(n, 50))
        if len(out) >= max_rows:
            break
    return out


class OdsReport:
    def __init__(self, template: str, log: Callable[[str], None] = print):
        self.template = template
        self.log = log
        with zipfile.ZipFile(template) as z:
            self.files = {n: z.read(n) for n in z.namelist()}
            self.order = z.namelist()
        raw = self.files["content.xml"].decode("utf-8")
        self.root_tag = re.search(r"<office:document-content[^>]*>", raw).group(0)
        for ev, (pfx, uri) in ET.iterparse(__import__("io").BytesIO(self.files["content.xml"]), events=["start-ns"]):
            ET.register_namespace(pfx, uri)
        self.doc = ET.fromstring(self.files["content.xml"])
        self.body = self.doc.find(f"{O}body/{O}spreadsheet")
        self.auto_styles = self.doc.find(O + "automatic-styles")
        self.col_width_cache: Dict[str, float] = {}
        self.renames: Dict[str, str] = {}
        self.row_styles: Dict[str, str] = {}
        self.data_ranges: Dict[str, tuple] = {}   # 分頁新名稱 → (表頭列1起算, 資料最後一列, 欄數)

    # ---------- 共用 ----------
    def tables(self):
        return self.body.findall(T + "table")

    def find_table(self, prefix: str):
        for t in self.tables():
            if t.get(T + "name", "").startswith(prefix):
                return t
        return None

    def rows(self, table):
        return [r for r in table.iter(T + "table-row")]

    def col_widths(self, table) -> List[float]:
        ws = []
        for c in table.iter(T + "table-column"):
            st = self.auto_styles.find(f"{S}style[@{S}name='{c.get(T + 'style-name')}']")
            w = 2.3
            if st is not None:
                p = st.find(S + "table-column-properties")
                if p is not None and p.get(S + "column-width", "").endswith("cm"):
                    w = float(p.get(S + "column-width")[:-2])
            ws.extend([w] * min(rep(c), 50))
            if len(ws) > 60:
                break
        return ws

    def row_style_for(self, height_cm: float) -> str:
        h = round(height_cm, 1)
        key = f"{h:.1f}"
        if key not in self.row_styles:
            name = f"roGen{len(self.row_styles) + 1}"
            st = ET.SubElement(self.auto_styles, S + "style", {S + "name": name, S + "family": "table-row"})
            ET.SubElement(st, S + "table-row-properties", {S + "row-height": f"{h}cm", FO + "break-before": "auto",
                                                          S + "use-optimal-row-height": "false"})
            self.row_styles[key] = name
        return self.row_styles[key]

    @staticmethod
    def est_lines(text: str, width_cm: float) -> int:
        per = max(1.0, (width_cm - 0.2) / CJK_CM)
        n = 0
        for para in (text or "").split("\n"):
            units = sum(0.5 if ord(ch) < 128 else 1.0 for ch in para)
            n += max(1, math.ceil(units / per))
        return n

    def rename(self, table, new: str):
        old = table.get(T + "name")
        if old != new:
            table.set(T + "name", new)
            self.renames[old] = new

    @staticmethod
    def update_date(text: str, roc_date: str) -> str:
        y, m, d = int(roc_date[:3]), int(roc_date[3:5]), int(roc_date[5:7])
        return re.sub(r"截至\s*\d{2,3}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", f"截至{y}年{m}月{d}日", text)

    def set_title_date(self, table, roc_date: str, extra: Optional[Callable[[str], str]] = None):
        for r in self.rows(table)[:3]:
            for c in r:
                if is_cell(c) and "截至" in cell_text(c):
                    t = self.update_date(cell_text(c), roc_date)
                    if extra: t = extra(t)
                    for ch in list(c): c.remove(ch)
                    p = ET.SubElement(c, X + "p"); _put_text(p, t)
                    return

    def header_index(self, table) -> int:
        for i, r in enumerate(self.rows(table)[:8]):
            cells = [c for c in r if is_cell(c)]
            if cells and _norm(cell_text(cells[0])) == "編號":
                return i
        raise ValueError(f"分頁「{table.get(T + 'name')}」找不到「編號」表頭列")

    def data_rows(self, table, h: int):
        """表頭之後、到大量重複的空白列之前，都算資料區。"""
        out = []
        for r in self.rows(table)[h + 1:]:
            if rep(r, T + "number-rows-repeated") > 100:
                break
            out.append(r)
        return out

    # ---------- 明細頁 ----------
    def fill_detail(self, table, key: str, df, roc_date: str, name_base: str):
        rows = self.rows(table)
        h = self.header_index(table)
        header_cells = [c for c in rows[h] if is_cell(c)]
        headers = []
        for c in header_cells:
            t = _norm(cell_text(c))
            if not t: break
            headers.append(t)
        ncol = len(headers)
        widths = self.col_widths(table)
        old = self.data_rows(table, h)
        if not old:
            raise ValueError(f"分頁「{table.get(T + 'name')}」沒有樣板資料列")
        proto = copy.deepcopy(old[0])
        parent = None
        for p in table.iter():
            if old[0] in list(p): parent = p; break
        pos = list(parent).index(old[0])
        for r in old:
            parent.remove(r)

        od_col = next((c for c in df.columns if "逾期天數" in c), None)
        const = CONSTANT.get(key, {})
        src_map = {**COLUMN_SOURCE, **COLUMN_SOURCE_BY_KEY.get(key, {})}
        missing = [hd for hd in headers if hd not in const and src_map.get(hd) not in (None,) and
                   src_map.get(hd) != "__逾期天數__" and src_map.get(hd) not in df.columns]
        for i, rec in enumerate(df.to_dict("records")):
            row = copy.deepcopy(proto)
            cells = split_cells(row, ncol)
            lines = 1
            for j, hd in enumerate(headers):
                if hd in const:
                    v = const[hd]
                else:
                    src = src_map.get(hd)
                    v = rec.get(od_col) if src == "__逾期天數__" else (rec.get(src) if src else None)
                set_cell(cells[j], v, hd in NUMERIC)
                if v is not None and j < len(widths):
                    lines = max(lines, self.est_lines(str(v), widths[j]))
            for a in (T + "number-rows-repeated",):
                row.attrib.pop(a, None)
            row.set(T + "style-name", self.row_style_for(max(1.2, lines * LINE_CM + 0.45)))
            parent.insert(pos + i, row)
        n = len(df)
        self.rename(table, f"{name_base}({n})")
        self.set_title_date(table, roc_date)
        self.data_ranges[table.get(T + "name")] = (h + 1, h + 1 + n, ncol)
        extra = [c for c in df.columns if c not in {src_map.get(hd) for hd in headers} | {od_col}]
        self.log(f"  {table.get(T + 'name')}：{n} 筆" + (f"（範本沒有對應欄位、未放入：{'、'.join(extra)}）" if extra and n else ""))
        if missing and n:
            self.log(f"  ⚠ {table.get(T + 'name')}：範本欄位找不到資料來源：{'、'.join(missing)}")
        return n

    def clear_detail(self, table, roc_date: str):
        h = self.header_index(table)
        old = self.data_rows(table, h)
        parent = next(p for p in table.iter() if old and old[0] in list(p)) if old else None
        for r in old:
            parent.remove(r)
        base = re.sub(r"\(.*\)$", "", table.get(T + "name"))
        self.rename(table, f"{base}(  )")
        self.set_title_date(table, roc_date)
        self.data_ranges[table.get(T + "name")] = (h + 1, h + 1, None)

    def add_sheet_like(self, like, after, title_word: str):
        new = copy.deepcopy(like)
        new.set(T + "name", "__new__" + title_word)
        idx = list(self.body).index(after)
        self.body.insert(idx + 1, new)
        for r in self.rows(new)[:2]:
            for c in r:
                if is_cell(c) and "特殊性案件" in cell_text(c):
                    t = cell_text(c).replace("特殊性案件", title_word)
                    for ch in list(c): c.remove(ch)
                    _put_text(ET.SubElement(c, X + "p"), t)
        return new

    # ---------- 總表 ----------
    def fill_summary(self, table, counts: Dict[str, Dict[str, int]], units: List[str], roc_date: str):
        rows = self.rows(table)
        hdr_i = next(i for i, r in enumerate(rows[:6]) if _norm(cell_text([c for c in r if is_cell(c)][0])).startswith("機關"))
        hdr = rows[hdr_i]
        hcells = [c for c in hdr if is_cell(c)]
        ncol = 0
        for c in hcells:
            if not _norm(cell_text(c)): break
            ncol += 1
        colkeys = []
        for c in hcells[1:ncol]:
            t = _norm(cell_text(c))
            colkeys.append(next((k for w, k in SUMMARY_KEYS if w in t), None))
        # 訴願／專案有案件卻沒有欄位 → 在最後加欄
        add = [k for k in ("D", "E") if sum(counts.get(k, {}).values()) and k not in colkeys]
        total_i = next(i for i, r in enumerate(rows) if _norm(cell_text([c for c in r if is_cell(c)][0])) == "合計")
        unit_rows = rows[hdr_i + 1:total_i]
        after_rows = [r for r in rows[total_i + 1:] if rep(r, T + "number-rows-repeated") < 100]
        if add:
            label = {"D": ["訴願案件", "(件數)"], "E": ["專案管制案件", "(件數)"]}
            for r in [hdr] + unit_rows + [rows[total_i]]:
                cs = split_cells(r, ncol)
                for k_i, k in enumerate(add):
                    nc = copy.deepcopy(cs[-1])
                    r.insert(list(r).index(cs[-1]) + 1 + k_i, nc)
                    if r is hdr:
                        for ch in list(nc): nc.remove(ch)
                        for line in label[k]:
                            ET.SubElement(nc, X + "p").text = line
            for r in rows[:hdr_i] + after_rows:        # 標題與說明列的合併範圍加寬
                for c in r:
                    if c.tag == T + "table-cell" and c.get(T + "number-columns-spanned"):
                        c.set(T + "number-columns-spanned", str(int(c.get(T + "number-columns-spanned")) + len(add)))
                        idx = list(r).index(c)
                        for _ in add:
                            r.insert(idx + 1, ET.Element(T + "covered-table-cell"))
                        break
            cols = [c for c in table.iter(T + "table-column")]
            # 第 ncol 欄（最後一個件數欄）的欄定義複製一份
            acc = 0
            for c in cols:
                r_ = rep(c)
                if acc + r_ >= ncol:
                    c.set(T + "number-columns-repeated", str(r_ + len(add)))
                    break
                acc += r_
            colkeys += add
            ncol += len(add)
            self.log(f"  總表新增欄位：{'、'.join(label[k][0] for k in add)}")
        proto_first = copy.deepcopy(unit_rows[0])
        proto_last = copy.deepcopy(unit_rows[-1])
        parent = next(p for p in table.iter() if unit_rows[0] in list(p))
        pos = list(parent).index(unit_rows[0])
        for r in unit_rows:
            parent.remove(r)
        for i, u in enumerate(units):
            row = copy.deepcopy(proto_last if i == len(units) - 1 else proto_first)
            cs = split_cells(row, ncol)
            set_cell(cs[0], u, False)
            for j, k in enumerate(colkeys, start=1):
                v = counts.get(k, {}).get(u, 0) if k else 0
                set_cell(cs[j], v if v else None, True)
            parent.insert(pos + i, row)
        # 合計列公式
        first, last = hdr_i + 2, hdr_i + 1 + len(units)
        cs = split_cells(rows[total_i], ncol)
        for j, k in enumerate(colkeys, start=1):
            col = chr(ord("A") + j)
            tot = sum(counts.get(k, {}).values()) if k else 0
            set_cell(cs[j], tot, True)
            cs[j].set(T + "formula", f"of:=SUM([.{col}{first}:.{col}{last}])")
        self.set_title_date(table, roc_date)
        self.data_ranges[table.get(T + "name")] = (hdr_i + 1, last + 1, ncol)

    # ---------- 存檔 ----------
    def _fix_refs(self):
        def fix(v: str) -> str:
            for old, new in self.renames.items():
                v = v.replace(f"'{old}'", f"'{new}'")
            return v
        for e in self.doc.iter():
            for a, v in list(e.attrib.items()):
                if "'" in v:
                    e.set(a, fix(v))
        # 篩選範圍、列印範圍延伸到資料最後一列
        for dr in self.doc.iter(T + "database-range"):
            m = re.match(r"^'?(.+?)'?\.([A-Z]+)(\d+):'?(.+?)'?\.([A-Z]+)(\d+)$", dr.get(T + "target-range-address", ""))
            if m and m.group(1) in self.data_ranges:
                h, end, _ = self.data_ranges[m.group(1)]
                q = f"'{m.group(1)}'"
                dr.set(T + "target-range-address", f"{q}.{m.group(2)}{m.group(3)}:{q}.{m.group(5)}{max(end, int(m.group(3)))}")
        for t in self.tables():
            pr = t.get(T + "print-ranges")
            nm = t.get(T + "name")
            if pr and nm in self.data_ranges:
                h, end, _ = self.data_ranges[nm]
                m = re.match(r"^(.*\.[A-Z]+\d+:.*\.[A-Z]+)(\d+)$", pr)
                if m:
                    t.set(T + "print-ranges", m.group(1) + str(max(end, h + 1)))

    def save(self, out_path: str):
        self._fix_refs()
        body = ET.tostring(self.doc, encoding="unicode")
        body = re.sub(r"^<office:document-content[^>]*>", self.root_tag, body, count=1)
        if re.search(r"<ns\d+:| ns\d+:", body):
            raise RuntimeError("內部錯誤：產生了未宣告的命名空間")
        self.files["content.xml"] = ('<?xml version="1.0" encoding="UTF-8"?>\n' + body).encode("utf-8")
        st = self.files["settings.xml"].decode("utf-8")
        for old, new in self.renames.items():
            st = st.replace(f'config:name="{old}"', f'config:name="{new}"').replace(f">{old}<", f">{new}<")
        self.files["settings.xml"] = st.encode("utf-8")
        # 舊縮圖會顯示上週資料，拿掉
        self.files.pop("Thumbnails/thumbnail.png", None)
        mf = self.files["META-INF/manifest.xml"].decode("utf-8")
        mf = re.sub(r'\s*<manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png"[^>]*/>', "", mf)
        self.files["META-INF/manifest.xml"] = mf.encode("utf-8")
        tmp = out_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(zipfile.ZipInfo("mimetype"), self.files["mimetype"], compress_type=zipfile.ZIP_STORED)
            for n in self.order:
                if n != "mimetype" and n in self.files:
                    z.writestr(n, self.files[n])
        os.replace(tmp, out_path)


def build_report(template: str, dfs: Dict[str, "object"], out_path: str, roc_date: str,
                 unit_key: Callable[[str], object], log: Callable[[str], None] = print) -> Dict[str, int]:
    """dfs：去重後、依類別（A～G）的資料表。回傳各類筆數。"""
    rpt = OdsReport(template, log)
    log(f"套用範本：{os.path.basename(template)}")
    result = {}
    for prefix, key in SHEET_PREFIX:
        t = rpt.find_table(prefix)
        if t is None:
            log(f"  ⚠ 範本沒有「{prefix[:-1]}」分頁，{key} 類略過"); continue
        df = dfs.get(key)
        if df is None and key in OPTIONAL_KEYS:
            rpt.clear_detail(t, roc_date)
            what = "列管" if key == "H" else "建管系統"
            log(f"  {t.get(T + 'name')}：沒有{what}資料，已清空，請人工填寫")
            continue
        if df is None:
            import pandas as pd
            df = pd.DataFrame()
        base = prefix[:-1]
        result[key] = rpt.fill_detail(t, key, df, roc_date, base)
        if key == "C":
            n = result[key]
            rpt.set_title_date(t, roc_date, lambda s: re.sub(r"公文\s*\d*\s*件", f"公文 {n}件", s))
    like = rpt.find_table("特殊性案件(")
    after = rpt.find_table("限期公文(")
    if after is None:
        after = like
    for key, word in NEW_SHEETS.items():
        df = dfs.get(key)
        if df is not None and len(df) and like is not None:
            # 用範本原始的特殊性案件頁再複製一次（不能用已填資料的那頁當樣板列）
            src = OdsReport(template, lambda m: None).find_table("特殊性案件(")
            new = rpt.add_sheet_like(src, after, word)
            result[key] = rpt.fill_detail(new, key, df, roc_date, word)
            after = new
    for prefix in CLEAR_SHEETS:
        t = rpt.find_table(prefix)
        if t is not None:
            rpt.clear_detail(t, roc_date)
            log(f"  {t.get(T + 'name')}：已清空，請人工填寫")
    counts: Dict[str, Dict[str, int]] = {}
    units = set()
    for k, df in dfs.items():
        if k == "BM" and df is not None and len(df):
            counts[k] = {BM_SUMMARY_UNIT: len(df)}     # 總表一律算在建設處（見 BM_SUMMARY_UNIT）
            units.add(BM_SUMMARY_UNIT)
            continue
        if df is not None and len(df) and "承辦單位" in df.columns:
            vc = df["承辦單位"].astype(str).str.strip().value_counts().to_dict()
            counts[k] = vc
            units |= set(vc)
    st = rpt.find_table(SUMMARY_PREFIX)
    if st is not None:
        rpt.fill_summary(st, counts, sorted(units, key=unit_key), roc_date)
        log(f"  總表：{len(units)} 個單位")
    rpt.save(out_path)
    # 115.09.23（使用者反映看到「需要人工確認 1 項」誤以為是這批資料本身有問題，
    # 其實是固定提醒＋列管案件核對兩件不同的事被塞成一句話）：拆成兩則訊息、講清楚
    # 「哪些是每次都會出現的固定提醒（不是資料異常）」、「哪些是這次真的要核對的東西」。
    建管 = ("「人民申請－建管系統」分頁的「備註」「流程」兩欄" if "BM" in result
          else "「人民申請－建管系統」分頁")
    log("  ⚠ 固定提醒（這行每次收工都會出現，不代表本次資料有問題）：總表下方的面報日期、"
        f"{建管}、陳情案件（1999／縣長信箱）程式沒辦法自動填，請自己補上／確認")
    h_count = result.get("H", 0)
    if h_count:
        log(f"  ⚠ 列管案件本次有 {h_count} 筆，請逐筆確認是否已符合解除列管條件")
    return result
