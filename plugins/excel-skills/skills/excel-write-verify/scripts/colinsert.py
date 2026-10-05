#!/usr/bin/env python3
"""Insert columns into a worksheet, at XML level. The sheet may hold an Excel table, also one loaded by Power Query.

  colinsert.py SRC.xlsx DST.xlsx spec.json

spec = {"sheet": "Data", "at": "K", "header_row": 1, "first_row": 2, "last_row": 200,
        "columns": [{"name": "Margin", "formula": "Sales[[#This Row],[Net]]/Sales[[#This Row],[Gross]]", "width": 12}]}

"at" is the letter where the first new column goes; what was there moves right. "formula" has no leading "=";
leave it out for an empty column. The formula text is written unchanged into every row from first_row to
last_row, so use structured references (Table[[#This Row],[Col]]) or references that do not depend on the row.
Optional per column: "header_style" and "data_style" (cell style indexes); by default the new cells take the
styles of the column on the left.

Shifts cells, in-sheet formula refs, shared/array refs, cols, merged cells, hyperlinks, sheet autofilter and
sort refs, CF/DV/x14 sqrefs, comments + VML anchors, table ref/columns/filter column numbers, queryTable
fields, defined names; extends the dimension; drops calcChain. Other zip parts stay byte-identical.
Stops (writes nothing) when
  - another sheet, a conditional format or validation of another sheet, another table, a chart or a pivot
    cache points at columns that would move: those references are not rewritten;
  - the sheet has an autofilter with a saved filter: clear the filter first.
Pictures and charts of the sheet keep their position (a WARNING line says so).
Exit code: 0 done | 1 stopped | 2 usage error. SRC is never modified; SRC and DST must differ.
Needs openpyxl for its formula tokenizer."""
import os, re, html, json, sys, uuid, zipfile
import xml.etree.ElementTree as ET
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
from xlsxpatch import sheet_paths, table_paths, esc, insert_cell, find_cell, F_OPEN, same_file, need

MAXC = 16384


class Book:
    def __init__(self, src):
        z = zipfile.ZipFile(src)
        self.infos = z.infolist()
        self.parts = {n: z.read(n) for n in z.namelist()}
        self.sp = sheet_paths(z)
        self.tp = table_paths(z)
        self.removed = set()

    def get(self, n):
        return self.parts[n].decode("utf8")

    def put(self, n, s):
        self.parts[n] = s.encode("utf8")

    def rel_targets(self, part):
        """{relationship id: part name} of a part; any attribute order, absolute or relative, no external."""
        d, f = part.rsplit("/", 1)
        rp = f"{d}/_rels/{f}.rels"
        if rp not in self.parts:
            return {}
        out = {}
        for r in ET.fromstring(self.parts[rp]):
            t = r.get("Target") or ""
            if r.get("TargetMode") == "External" or not t:
                continue
            if t.startswith("/"):
                out[r.get("Id")] = t[1:]
                continue
            base = d.split("/")
            for seg in t.split("/"):
                if seg == "..":
                    base.pop()
                elif seg != ".":
                    base.append(seg)
            out[r.get("Id")] = "/".join(base)
        return out

    def drop_calc_chain(self):
        if "xl/calcChain.xml" not in self.parts:
            return
        self.removed.add("xl/calcChain.xml")
        r = self.get("xl/_rels/workbook.xml.rels")
        self.put("xl/_rels/workbook.xml.rels", re.sub(r'<Relationship [^>]*Target="calcChain\.xml"[^>]*/>', "", r))
        ct = self.get("[Content_Types].xml")
        self.put("[Content_Types].xml", re.sub(r'<Override [^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", ct))

    def save(self, dst):
        with zipfile.ZipFile(dst, "w") as zout:
            for info in self.infos:
                if info.filename in self.removed:
                    continue
                zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
                zout.writestr(zi, self.parts[info.filename])


def _col(letter, at, n):
    i = ci(letter)
    return gl(min(i + n, MAXC)) if i >= at else letter


def shift_ref(ref, at, n):
    """A1, $A$1, A1:B2, $A:$A, 1:1 -> shifted; None if not a plain reference."""
    out = []
    for p in ref.split(":"):
        m = re.fullmatch(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", p)
        if m:
            out.append(m.group(1) + _col(m.group(2), at, n) + m.group(3) + m.group(4))
            continue
        m = re.fullmatch(r"(\$?)([A-Z]{1,3})", p)
        if m:
            out.append(m.group(1) + _col(m.group(2), at, n))
            continue
        if re.fullmatch(r"\$?\d+", p):
            out.append(p)
            continue
        return None
    return ":".join(out)


def shift_sqref(v, at, n):
    return " ".join(shift_ref(r, at, n) or r for r in v.split())


CAND = re.compile(r"(?<![A-Za-z_\[.])\$?([A-Z]{1,3})\$?\d+|(?<![A-Za-z_\[.])\$?([A-Z]{1,3}):\$?([A-Z]{1,3})\b")


def needs_shift(f, at):
    for m in CAND.finditer(f):
        cols = [c for c in m.groups() if c]
        if any(ci(c) >= at for c in cols if len(c) <= 3):
            return True
    return False


def shift_formula(f, at, n, sheet, unqualified=True):
    """f without leading '='. Refs to `sheet` move (unqualified ones too if unqualified=True); structured refs never move."""
    if not needs_shift(f, at):
        return f
    tok = Tokenizer("=" + f)
    if tok.render() != "=" + f:
        raise SystemExit(f"tokenizer round-trip failed: {f[:120]}")
    for t in tok.items:
        if t.type != "OPERAND" or t.subtype != "RANGE" or "[" in t.value:
            continue
        v = t.value
        if "!" in v:
            sh, r = v.rsplit("!", 1)
            if sh.strip("'").replace("''", "'") != sheet:
                continue
            nr = shift_ref(r, at, n)
            if nr:
                t.value = sh + "!" + nr
        elif unqualified:
            nr = shift_ref(v, at, n)
            if nr:
                t.value = nr
    return tok.render()[1:]


def _sub_ftext(x, at, n, sheet, tag_re=r"(" + F_OPEN + r")(.*?)(</f>)", unqualified=True):
    def rep(m):
        txt = html.unescape(m.group(2))
        new = shift_formula(txt, at, n, sheet, unqualified)
        return m.group(1) + (m.group(2) if new == txt else esc(new)) + m.group(3)
    return re.sub(tag_re, rep, x, flags=re.S)


def _hits(f, at, n, sheet):
    """True when formula text f (no '=') has a reference to a column >= at on `sheet`."""
    if (sheet + "!") not in f and (f"'{sheet}'!") not in f:
        return False
    return shift_formula(f, at, n, sheet, unqualified=False) != f


def _grow_dim(ref, at, n):
    """Dimension after a shift; a column appended past the old end extends it."""
    ref = shift_ref(ref, at, n) or ref
    m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", ref)
    if not m:
        return ref
    c2, r2 = (m.group(3), m.group(4)) if m.group(3) else (m.group(1), m.group(2))
    if ci(c2) < at:
        c2 = gl(at + n - 1)
    return f"{m.group(1)}{m.group(2)}:{c2}{r2}"


def check_guard(book, sheet, at, n, at_letter, own_tables):
    """Stop when a part that colinsert does not rewrite points at columns that would move."""
    def stop(where, f):
        raise SystemExit(f"{where} references {sheet} columns >= {at_letter}: {f[:120]}")
    for name, p in book.sp.items():
        if name == sheet:
            continue        # this sheet's own formulas, CF and DV are shifted by insert_cols
        ox = book.get(p)
        for tg in ("f", "formula", "formula1", "formula2", "xm:f"):
            rx = F_OPEN + "(.*?)</f>" if tg == "f" else r"<%s>(.*?)</%s>" % (tg, tg)
            for m in re.finditer(rx, ox, re.S):
                f = html.unescape(m.group(1))
                if _hits(f, at, n, sheet):
                    stop(f"sheet {name!r}" + ("" if tg == "f" else f" ({tg})"), f)
    for tpart in sorted(k for k in book.parts if k.startswith("xl/tables/") and k.endswith(".xml")):
        if tpart in own_tables:
            continue
        for m in re.finditer(r"<(?:calculatedColumnFormula|totalsRowFormula)\b[^>]*>(.*?)</", book.get(tpart), re.S):
            f = html.unescape(m.group(1))
            if _hits(f, at, n, sheet):
                stop(f"table part {tpart}", f)
    for cpart in sorted(k for k in book.parts if re.match(r"xl/charts/[^/]+\.xml$", k)):
        for m in re.finditer(r"<c:f>(.*?)</c:f>", book.get(cpart), re.S):
            f = html.unescape(m.group(1))
            if _hits(f, at, n, sheet):
                stop(f"chart {cpart}", f)
    for ppart in sorted(k for k in book.parts if re.match(r"xl/pivotCache/pivotCacheDefinition[^/]*\.xml$", k)):
        for m in re.finditer(r"<worksheetSource\b[^>]*>", book.get(ppart)):
            sh, rf = re.search(r'\bsheet="([^"]*)"', m.group(0)), re.search(r'\bref="([^"]*)"', m.group(0))
            if sh and rf and html.unescape(sh.group(1)) == sheet and shift_ref(rf.group(1), at, n) not in (None, rf.group(1)):
                stop(f"pivot cache {ppart}", m.group(0))


def insert_cols(book, sheet, at_letter, new_cols, header_row, first_row, last_row):
    """new_cols: [{name, formula (no '='; None for empty), header_style, data_style, width}] inserted at at_letter."""
    at, n = ci(at_letter), len(new_cols)
    sp = book.sp[sheet]
    x = book.get(sp)
    head, body_tail = x.split("<sheetData", 1)
    sheet_data, tail = body_tail.split("</sheetData>", 1)
    sheet_data = "<sheetData" + sheet_data + "</sheetData>"

    rels = book.rel_targets(sp)
    own_tables = {t for t in rels.values() if "/tables/" in t}
    for m in re.finditer(r"<autoFilter\b[^>]*?(?:/>|>.*?</autoFilter>)", tail, re.S):
        if "<filterColumn" in m.group(0):
            raise SystemExit(f"sheet {sheet!r} has an autofilter with a saved filter: clear the filter first")
    check_guard(book, sheet, at, n, at_letter, own_tables)
    if "<drawing" in x:
        print("WARNING: the sheet has pictures or charts; they keep their position and do not move with the columns")

    # cells + formulas + shared/array refs + spans
    sheet_data = re.sub(r'(<c r=")([A-Z]+)(\d+)"', lambda m: m.group(1) + _col(m.group(2), at, n) + m.group(3) + '"', sheet_data)
    sheet_data = re.sub(r'(<f\b[^>]*\bref=")([^"]+)(")', lambda m: m.group(1) + (shift_ref(m.group(2), at, n) or m.group(2)) + m.group(3), sheet_data)
    sheet_data = _sub_ftext(sheet_data, at, n, sheet)
    sheet_data = re.sub(r'\sspans="[^"]*"', "", sheet_data)

    # head: dimension, views, cols
    head = re.sub(r'(<dimension ref=")([^"]+)(")', lambda m: m.group(1) + _grow_dim(m.group(2), at, n) + m.group(3), head)
    head = re.sub(r'((?:activeCell|topLeftCell)=")([^"]+)(")', lambda m: m.group(1) + (shift_ref(m.group(2), at, n) or m.group(2)) + m.group(3), head)
    head = re.sub(r'(sqref=")([^"]+)(")', lambda m: m.group(1) + shift_sqref(m.group(2), at, n) + m.group(3), head)

    def fix_cols(m):
        cols = re.findall(r"<col [^>]*/>", m.group(0))
        out = []
        for c in cols:
            lo, hi = int(re.search(r'min="(\d+)"', c).group(1)), int(re.search(r'max="(\d+)"', c).group(1))
            if lo >= at:
                lo, hi = min(lo + n, MAXC), min(hi + n, MAXC)
            elif hi >= at:
                hi = min(hi + n, MAXC)
            out.append((lo, re.sub(r'min="\d+"', f'min="{lo}"', re.sub(r'max="\d+"', f'max="{hi}"', c))))
        for k, nc in enumerate(new_cols):
            out.append((at + k, f'<col min="{at + k}" max="{at + k}" width="{nc.get("width", 14)}" customWidth="1"/>'))
        out.sort(key=lambda t: t[0])
        return "<cols>" + "".join(c for _, c in out) + "</cols>"
    head = re.sub(r"<cols>.*?</cols>", fix_cols, head, flags=re.S)

    # tail: merged cells, hyperlinks, sheet autofilter and sort refs, CF / DV / x14 sqrefs and formulas
    tail = re.sub(r'(<(?:mergeCell|hyperlink|autoFilter|sortState|sortCondition)\b[^>]*?\bref=")([^"]+)(")',
                  lambda m: m.group(1) + (shift_ref(m.group(2), at, n) or m.group(2)) + m.group(3), tail)
    tail = re.sub(r'(sqref=")([^"]+)(")', lambda m: m.group(1) + shift_sqref(m.group(2), at, n) + m.group(3), tail)
    tail = re.sub(r"(<xm:sqref>)([^<]+)(</xm:sqref>)", lambda m: m.group(1) + shift_sqref(m.group(2), at, n) + m.group(3), tail)
    for tg in ("formula", "formula1", "formula2", "xm:f"):
        tail = _sub_ftext(tail, at, n, sheet, tag_re=r"(<%s>)(.*?)(</%s>)" % (tg, tg))
    x = head + sheet_data + tail

    # new cells
    for k, nc in enumerate(new_cols):
        col = gl(at + k)
        x = insert_cell(x, f"{col}{header_row}", f'<c r="{col}{header_row}" s="{nc["header_style"]}" t="inlineStr"><is><t>{esc(nc["name"])}</t></is></c>')
        if nc.get("formula"):
            for r in range(first_row, last_row + 1):
                x = insert_cell(x, f"{col}{r}", f'<c r="{col}{r}" s="{nc["data_style"]}"><f>{esc(nc["formula"])}</f></c>')
    book.put(sp, x)

    # table + queryTable: every table of the sheet is shifted; columns are added when `at` is inside the
    # table or right after its last column (not at its first column, not outside it)
    for tpath in sorted(own_tables):
        if tpath not in book.parts:
            continue
        tx = book.get(tpath)
        tref = re.search(r'<table [^>]*\bref="([^"]+)"', tx).group(1)
        c1 = ci(re.match(r"[A-Z]+", tref).group(0))
        c2 = ci(re.match(r"[A-Z]+\d+:([A-Z]+)", tref).group(1))
        if at == c2 + 1:  # appending right after the last column: grow the table
            grow = lambda r: re.sub(r":([A-Z]+)(\d+)$", lambda m: ":" + gl(ci(m.group(1)) + n) + m.group(2), r)
        else:
            grow = lambda r: shift_ref(r, at, n)
        tx = re.sub(r'(<table [^>]*\bref=")([^"]+)(")', lambda m: m.group(1) + grow(m.group(2)) + m.group(3), tx, count=1)
        tx = re.sub(r'(<(?:autoFilter|sortState)\b[^>]*?\bref=")([^"]+)(")', lambda m: m.group(1) + (grow(m.group(2)) or m.group(2)) + m.group(3), tx)
        tx = re.sub(r'(<sortCondition\b[^>]*?\bref=")([^"]+)(")',      # one column: it moves, it never grows
                    lambda m: m.group(1) + (shift_ref(m.group(2), at, n) or m.group(2)) + m.group(3), tx)
        tx = _sub_ftext(tx, at, n, sheet, tag_re=r"(<calculatedColumnFormula(?:\s[^>]*)?>)(.*?)(</calculatedColumnFormula>)")
        cols = list(re.finditer(r"<tableColumn\b[^>]*?(?:/>|>.*?</tableColumn>)", tx, re.S))
        pos = at - c1
        if not (1 <= pos <= len(cols)):
            book.put(tpath, tx)      # shifted only
            continue
        tx = re.sub(r'(<filterColumn\b[^>]*?\bcolId=")(\d+)(")',
                    lambda m: m.group(1) + str(int(m.group(2)) + (n if int(m.group(2)) >= pos else 0)) + m.group(3), tx)
        ids = [int(i) for i in re.findall(r'<tableColumn id="(\d+)"', tx)]
        is_qt = 'tableType="queryTable"' in tx
        qt_path = next((t for t in book.rel_targets(tpath).values() if "/queryTables/" in t), None) if is_qt else None
        qx = book.get(qt_path) if qt_path else None
        qnext = int(re.search(r'nextId="(\d+)"', qx).group(1)) if qx else None
        new_tc, new_qf = [], []
        for k, nc in enumerate(new_cols):
            nid = max(ids) + 1 + k
            attrs = f'id="{nid}"' + (f' xr3:uid="{{{str(uuid.uuid4()).upper()}}}"' if "xmlns:xr3" in tx else "")
            if is_qt:
                attrs += f' uniqueName="{nid}" name="{esc(nc["name"])}" queryTableFieldId="{qnext + k}"'
                new_qf.append(f'<queryTableField id="{qnext + k}" dataBound="0" tableColumnId="{nid}"/>')
            else:
                attrs += f' name="{esc(nc["name"])}"'
            inner = f'<calculatedColumnFormula>{esc(nc["formula"])}</calculatedColumnFormula>' if nc.get("formula") else ""
            new_tc.append(f"<tableColumn {attrs}>{inner}</tableColumn>" if inner else f"<tableColumn {attrs}/>")
        cols = list(re.finditer(r"<tableColumn\b[^>]*?(?:/>|>.*?</tableColumn>)", tx, re.S))
        ins = cols[pos].start() if pos < len(cols) else cols[-1].end()
        tx = tx[:ins] + "".join(new_tc) + tx[ins:]
        tx = re.sub(r'(<tableColumns count=")(\d+)(")', lambda m: m.group(1) + str(int(m.group(2)) + n) + m.group(3), tx, count=1)
        book.put(tpath, tx)
        if qx:
            fields = list(re.finditer(r"<queryTableField\b[^>]*?(?:/>|>.*?</queryTableField>)", qx, re.S))
            qins = fields[pos].start() if pos < len(fields) else fields[-1].end()
            qx = qx[:qins] + "".join(new_qf) + qx[qins:]
            qx = re.sub(r'(<queryTableFields count=")(\d+)(")', lambda m: m.group(1) + str(int(m.group(2)) + n) + m.group(3), qx, count=1)
            qx = re.sub(r'nextId="\d+"', f'nextId="{qnext + n}"', qx, count=1)
            qx = re.sub(r'(unboundColumnsRight=")(\d+)(")', lambda m: m.group(1) + str(int(m.group(2)) + n) + m.group(3), qx, count=1)
            book.put(qt_path, qx)

    # comments + VML
    for t in rels.values():
        if t not in book.parts:
            continue
        if re.search(r"/comments\d*\.xml$", t):
            cx = book.get(t)
            cx = re.sub(r'(<comment ref=")([^"]+)(")', lambda m: m.group(1) + shift_ref(m.group(2), at, n) + m.group(3), cx)
            book.put(t, cx)
        if t.endswith(".vml"):
            vx = book.get(t)
            vx = re.sub(r"(<x:Column>)(\d+)(</x:Column>)", lambda m: m.group(1) + str(int(m.group(2)) + (n if int(m.group(2)) >= at - 1 else 0)) + m.group(3), vx)
            def anc(m):
                v = [s.strip() for s in m.group(2).split(",")]
                for i in (0, 4):
                    if int(v[i]) >= at - 1:
                        v[i] = str(int(v[i]) + n)
                return m.group(1) + ", ".join(v) + m.group(3)
            vx = re.sub(r"(<x:Anchor>)([^<]+)(</x:Anchor>)", anc, vx)
            book.put(t, vx)

    # defined names
    wb = book.get("xl/workbook.xml")
    wb = _sub_ftext(wb, at, n, sheet, tag_re=r"(<definedName\b[^>]*>)(.*?)(</definedName>)", unqualified=False)
    book.put("xl/workbook.xml", wb)
    book.drop_calc_chain()


def main(src, dst, spec):
    if same_file(src, dst):
        raise SystemExit("SRC and DST are the same file; write the result to a new file")
    book = Book(src)
    x = book.get(need(book.sp, spec["sheet"], "sheet"))
    at = ci(spec["at"])
    near = gl(at - 1) if at > 1 else gl(at)

    def style(row):
        m = find_cell(x, f"{near}{row}")
        s = re.search(r'\ss="(\d+)"', m.group(0)) if m else None
        return s.group(1) if s else "0"
    for c in spec["columns"]:
        c.setdefault("header_style", style(spec["header_row"]))
        c.setdefault("data_style", style(spec["first_row"]))
    insert_cols(book, spec["sheet"], spec["at"], spec["columns"], spec["header_row"], spec["first_row"], spec["last_row"])
    wb = book.get("xl/workbook.xml")
    if "fullCalcOnLoad" not in wb:
        book.put("xl/workbook.xml", re.sub(r"<calcPr([^/>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1))
    book.save(dst)
    print(f"inserted {len(spec['columns'])} column(s) at {spec['sheet']}!{spec['at']}; calcChain dropped; fullCalcOnLoad=1")


if __name__ == "__main__":
    if len(sys.argv) == 2 and sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        raise SystemExit(0)
    if len(sys.argv) != 4:
        print(__doc__, file=sys.stderr)
        raise SystemExit(2)
    main(sys.argv[1], sys.argv[2], json.load(open(sys.argv[3])))
