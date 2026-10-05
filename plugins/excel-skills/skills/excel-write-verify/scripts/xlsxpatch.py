#!/usr/bin/env python3
"""Surgical XLSX patcher: edits formula text / constants inside sheet XML, keeps every other
zip part byte-identical. Never round-trips the workbook through a spreadsheet library.

  xlsxpatch.py apply  SRC.xlsx DST.xlsx spec.json      apply a spec, print what was done
  xlsxpatch.py cached SRC.xlsx DST.xlsx values.json    store results in formula cells: {sheet: {ref: value}}
  xlsxpatch.py diff   A.xlsx B.xlsx                    changed zip parts and changed cells

Exit code: 0 done | 1 an operation refused or the workbook is not what the spec expects | 2 usage error.
SRC is never modified; SRC and DST must be different files.
Operations that compare first (they stop on a mismatch): rewrite, formula_replace, table_rewrite,
table_formula_replace, table_add_default. set_value refuses a formula cell. set_formula replaces whatever
the cell holds (it refuses the master of a shared or array group of several cells).
A formula in "old", "new" and "formula" may start with "=": one leading "=" is removed.

spec = {
  "set_value":   [{"sheet": "Data", "cell": "C9", "value": 40}],                 # number, text or true/false; refuses formula cells
  "set_formula": [{"sheet": "Data", "cell": "D9", "formula": "=C9*2"}],           # optional "style": index or a new_styles name
  "rewrite":     [{"sheet": "Data", "cells": "N2:N147", "old": "=A{r}*2", "new": "=A{r}*3"}],   # {r} = row; old must match
  "formula_replace":       [{"sheet": "Data", "cells": "N2:N147", "find": "...", "replace": "..."}],  # find must occur once
  "table_rewrite":         [{"table": "Sales", "column": "Net", "old": "...", "new": "..."}],       # calculated column
                           # table_rewrite and table_formula_replace stop when the column has no formula in the
                           # table part, unless the entry has "allow_missing": true
  "table_formula_replace": [{"table": "Sales", "column": "Net", "find": "...", "replace": "..."}],
  "table_add_default":     [{"table": "Sales", "column": "Net", "formula": "..."}],
  "new_styles":    [{"name": "pct", "base": 12, "format": "0.0%"}],              # clone cell style 12 with a number format
  "set_dimension": [{"sheet": "Data", "ref": "A1:Z200"}],
  "full_calc_on_load": true,      # default: Excel recalculates everything at the next open
  "drop_calc_chain": true         # default: remove the stale calculation chain; Excel rebuilds it
}

rewrite and formula_replace also remove the stored result of the cells they change.
Needs openpyxl only for its reference helpers (range parsing, shared-formula translation).
"""
import os, re, sys, json, zipfile, html
from openpyxl.utils import range_boundaries, get_column_letter, column_index_from_string

ERRORS = {"#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"}
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# An opening <f ...> tag that is not self-closing (<f t="shared" si="1"/> has no text and no </f>).
F_OPEN = r"<f(?:\s[^>]*?)?(?<!/)>"


def lead_eq(s):
    """One leading '=' of a formula is not part of the formula text in the file."""
    return s[1:] if s.startswith("=") else s


def same_file(a, b):
    try:
        return os.path.samefile(a, b)
    except OSError:
        return os.path.abspath(a) == os.path.abspath(b)


def need(mapping, key, what):
    if key not in mapping:
        raise SystemExit(f"unknown {what} {key!r}; the workbook has: {', '.join(map(repr, mapping))}")
    return mapping[key]


def drop_result(cx):
    """Remove the stored result and its type from a cell (<c ... t=".."> ... <v>..</v>)."""
    m = re.match(r"<c\b[^>]*>", cx)
    head = re.sub(r'\st="[^"]*"', "", m.group(0)) if m and not m.group(0).endswith("/>") else m.group(0)
    return head + re.sub(r"<v>.*?</v>|<v/>", "", cx[m.end():], flags=re.S)


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid2t = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid2t.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = {}
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        t = rid2t[rid]
        out[html.unescape(name)] = "xl/" + t.lstrip("/").replace("xl/", "") if not t.startswith("/xl/") else t[1:]
    return out


def table_paths(z):
    out = {}
    for n in z.namelist():
        if n.startswith("xl/tables/table") and n.endswith(".xml"):
            x = z.read(n).decode("utf8")
            m = re.search(r'displayName="([^"]+)"', x)
            out[m.group(1)] = n
    return out


def cells_in(rng):
    c1, r1, c2, r2 = range_boundaries(rng)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            yield f"{get_column_letter(c)}{r}"


CELL_RE = r'<c r="%s"(?:\s[^>]*?)?(?:/>|>(?:(?!<c[\s>]).)*?</c>)'  # never runs into the next cell


def find_cell(x, ref):
    return re.search(CELL_RE % ref, x, re.S)


def insert_cell(x, ref, cell_xml):
    """Insert a new <c> into its row in column order (row must exist)."""
    row = int(re.sub(r"[A-Z]+", "", ref))
    col = column_index_from_string(re.sub(r"\d+", "", ref))
    m = re.search(r'<row r="%d"(?:\s[^>]*?)?(?:/>|>(?:(?!<row[\s>]).)*?</row>)' % row, x, re.S)
    if not m:
        raise SystemExit(f"row {row} missing for {ref}")
    rx = m.group(0)
    if rx.endswith("/>"):
        new = rx[:-2] + ">" + cell_xml + "</row>"
    else:
        pos = None
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"', rx):
            if column_index_from_string(cm.group(1)) > col:
                pos = cm.start()
                break
        body_end = rx.rfind("</row>")
        pos = body_end if pos is None else pos
        new = rx[:pos] + cell_xml + rx[pos:]
    # drop explicit spans so Excel recomputes them
    new = re.sub(r'\sspans="[^"]*"', "", new, count=1)
    return x[:m.start()] + new + x[m.end():]


def style_of(cell_xml):
    m = re.search(r'\ss="(\d+)"', cell_xml)
    return f' s="{m.group(1)}"' if m else ""


def _formula_of(x, ref, shared_masters):
    m = find_cell(x, ref)
    if not m:
        return None, None
    cx = m.group(0)
    fm = re.search(r"<f((?:\s[^>]*?)?)(?<!/)>(.*?)</f>|<f((?:\s[^>]*?)?)/>", cx, re.S)
    if not fm:
        return m, None
    attrs = fm.group(1) if fm.group(1) is not None else fm.group(3)
    text = fm.group(2)
    si = re.search(r'si="(\d+)"', attrs or "")
    if 't="shared"' in (attrs or "") and text is None:
        mref, mtext = shared_masters[si.group(1)]
        from openpyxl.formula.translate import Translator
        text = Translator("=" + html.unescape(mtext), origin=mref).translate_formula(ref)[1:]
        return m, text
    if 't="array"' in (attrs or ""):
        return m, "ARRAY:" + html.unescape(text)
    return m, html.unescape(text)


def rewrite_formulas(x, cells, old_tpl, new_tpl, label):
    """Replace formula of every cell in `cells` (A1 range) whose current formula == old_tpl(row).
    {r} in templates = row number. Shared formulas are un-shared. Fails loudly on any mismatch."""
    shared_masters = {}
    for mm in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>\s*<f t="shared" ref="([^"]+)" si="(\d+)">(.*?)</f>', x, re.S):
        shared_masters[mm.group(3)] = (mm.group(1), mm.group(4))
    targets = list(cells_in(cells))
    tset = set(targets)
    # every shared group touched must be fully inside the target range
    for si, (mref, _) in shared_masters.items():
        gm = re.search(r'<c r="%s"[^>]*>\s*<f t="shared" ref="([^"]+)"' % mref, x)
        group = set(cells_in(gm.group(1)))
        if group & tset and not group <= tset:
            raise SystemExit(f"{label}: shared group {gm.group(1)} extends outside {cells}")
    n = 0
    for ref in targets:
        r = int(re.sub(r"[A-Z]+", "", ref))
        m, cur = _formula_of(x, ref, shared_masters)
        if m is None:
            raise SystemExit(f"{label}: {ref} missing")
        exp = old_tpl.replace("{r}", str(r))
        if cur != exp:
            raise SystemExit(f"{label}: {ref} formula mismatch\n  cur={cur}\n  exp={exp}")
        cx = m.group(0)
        new_f = "<f>" + esc(new_tpl.replace("{r}", str(r))) + "</f>"
        ncx = re.sub(F_OPEN + r".*?</f>|<f(?:\s[^>]*?)?/>", lambda _: new_f, cx, count=1, flags=re.S)
        ncx = drop_result(ncx)
        x = x[:m.start()] + ncx + x[m.end():]
        n += 1
    return x, n


def add_xf(styles, base_index, format_code):
    """Clone cellXfs[base_index] with a number format; return (styles, new_index)."""
    nfs = dict((v, k) for k, v in re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles))
    fc = esc(format_code).replace('"', "&quot;")
    if fc in nfs:
        nid = nfs[fc]
    else:
        ids = [int(i) for i in re.findall(r'<numFmt numFmtId="(\d+)"', styles)] or [163]
        nid = str(max(max(ids), 163) + 1)
        if "<numFmts" in styles:
            styles = re.sub(r'<numFmts count="(\d+)">', lambda m: f'<numFmts count="{int(m.group(1)) + 1}">', styles, count=1)
            styles = styles.replace("</numFmts>", f'<numFmt numFmtId="{nid}" formatCode="{fc}"/></numFmts>', 1)
        else:
            styles = styles.replace("<fonts", f'<numFmts count="1"><numFmt numFmtId="{nid}" formatCode="{fc}"/></numFmts><fonts', 1)
    m = re.search(r"<cellXfs count=\"(\d+)\">(.*?)</cellXfs>", styles, re.S)
    xfs = re.findall(r"<xf [^>]*?(?:/>|>.*?</xf>)", m.group(2), re.S)
    base = xfs[base_index]
    new = re.sub(r'numFmtId="\d+"', f'numFmtId="{nid}"', base, count=1)
    if "applyNumberFormat" not in new:
        new = new.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
    count = int(m.group(1))
    body = m.group(2) + new
    styles = styles[:m.start()] + f'<cellXfs count="{count + 1}">' + body + "</cellXfs>" + styles[m.end():]
    return styles, count


def name_rx(name):
    """Regex for an XML attribute value: & and < are always escaped, > may or may not be."""
    s = name.replace("&", "&amp;").replace("<", "&lt;")
    return re.escape(s).replace(">", "(?:>|&gt;)")


def table_col_inner(x, colname):
    """Return (start, end) of the inner XML of <tableColumn name=colname>, or None if self-closing/absent."""
    m = re.search(r'<tableColumn\b[^>]*\bname="%s"[^>]*>' % name_rx(colname), x)
    if not m:
        raise SystemExit(f"tableColumn {colname!r} not found")
    if m.group(0).endswith("/>"):
        return None
    end = x.index("</tableColumn>", m.end())
    return m.end(), end


def apply(src, dst, spec):
    if same_file(src, dst):
        raise SystemExit("SRC and DST are the same file; write the result to a new file")
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}
    sp = sheet_paths(zin)
    tp = table_paths(zin)
    touched = {}
    log = []

    def get(p):
        return touched.get(p, parts[p].decode("utf8"))


    style_idx = {}
    for e in spec.get("new_styles", []):
        st = get("xl/styles.xml")
        st, idx = add_xf(st, e["base"], e["format"])
        touched["xl/styles.xml"] = st
        style_idx[e["name"]] = idx
        log.append(f"new_style {e['name']} = xf {idx} ({e['format']})")
    for e in spec.get("set_value", []) + spec.get("set_formula", []):
        if isinstance(e.get("style"), str):
            e["style"] = need(style_idx, e["style"], "new_styles name")
    for e in spec.get("rewrite", []):
        p = need(sp, e["sheet"], "sheet")
        x, n = rewrite_formulas(get(p), e["cells"], lead_eq(e["old"]), lead_eq(e["new"]), f"{e['sheet']}!{e['cells']}")
        touched[p] = x
        log.append(f"rewrite {e['sheet']}!{e['cells']}: {n} cells")
    for e in spec.get("table_rewrite", []):
        p = need(tp, e["table"], "table")
        x = get(p)
        span = table_col_inner(x, e["column"])
        if span is None or "<calculatedColumnFormula" not in x[span[0]:span[1]]:
            if not e.get("allow_missing"):
                raise SystemExit(f"table {e['table']}[{e['column']}] has no formula in the table part (set allow_missing to skip)")
            log.append(f"table_rewrite {e['table']}[{e['column']}]: no calculatedColumnFormula (skipped)")
            continue
        inner0 = x[span[0]:span[1]]
        cm = re.search(r"<calculatedColumnFormula>(.*?)</calculatedColumnFormula>", inner0, re.S)
        r0 = str(e.get("row", 2))
        cur = html.unescape(cm.group(1))
        exp = lead_eq(e["old"]).replace("{r}", r0)
        if cur != exp:
            if not e.get("force"):
                raise SystemExit(f"table {e['table']}[{e['column']}] formula mismatch\n  cur={cur}\n  exp={exp}")
            log.append(f"  (forced: table default was {cur[:80]})")
        inner = inner0.replace(cm.group(0), "<calculatedColumnFormula>" + esc(lead_eq(e["new"]).replace("{r}", r0)) + "</calculatedColumnFormula>")
        x = x[:span[0]] + inner + x[span[1]:]
        touched[p] = x
        log.append(f"table_rewrite {e['table']}[{e['column']}]")

    for e in spec.get("table_add_default", []):
        p = need(tp, e["table"], "table")
        x = get(p)
        m = re.search(r'<tableColumn\b[^>]*\bname="%s"[^>]*>' % name_rx(e["column"]), x)
        if not m:
            raise SystemExit(f"tableColumn {e['column']} not found in {e['table']}")
        tag = m.group(0)
        cf = "<calculatedColumnFormula>" + esc(lead_eq(e["formula"])) + "</calculatedColumnFormula>"
        if tag.endswith("/>"):
            new_tag = tag[:-2].rstrip() + ">" + cf + "</tableColumn>"
            x = x[:m.start()] + new_tag + x[m.end():]
        else:
            end = x.index("</tableColumn>", m.end())
            if "<calculatedColumnFormula" in x[m.end():end]:
                raise SystemExit(f"{e['table']}[{e['column']}] already has a default")
            x = x[:m.end()] + cf + x[m.end():]
        touched[p] = x
        log.append(f"table_add_default {e['table']}[{e['column']}]")

    for e in spec.get("formula_replace", []):
        p = need(sp, e["sheet"], "sheet")
        x = get(p)
        find, repl = esc(e["find"]), esc(e["replace"])
        n = 0
        for ref in cells_in(e["cells"]):
            m = find_cell(x, ref)
            if not m:
                if e.get("skip_missing"):
                    continue
                raise SystemExit(f"{e['sheet']}!{ref} missing")
            cx = m.group(0)
            fm = re.search(r"<f>(.*?)</f>", cx, re.S)
            if not fm:
                if e.get("skip_nonformula"):
                    continue
                raise SystemExit(f"{e['sheet']}!{ref} has no plain <f> (shared/array?): {cx[:200]}")
            ftxt = fm.group(1)
            if ftxt.count(find) != 1:
                raise SystemExit(f"{e['sheet']}!{ref}: find occurs {ftxt.count(find)}x in {ftxt[:300]}")
            ncx = drop_result(cx.replace(fm.group(0), "<f>" + ftxt.replace(find, repl) + "</f>"))
            x = x[:m.start()] + ncx + x[m.end():]
            n += 1
        touched[p] = x
        log.append(f"formula_replace {e['sheet']}!{e['cells']}: {n} cells")

    for e in spec.get("table_formula_replace", []):
        p = need(tp, e["table"], "table")
        x = get(p)
        span = table_col_inner(x, e["column"])
        if span is None or "<calculatedColumnFormula" not in x[span[0]:span[1]]:
            if not e.get("allow_missing"):
                raise SystemExit(f"table {e['table']}[{e['column']}] has no formula in the table part (set allow_missing to skip)")
            log.append(f"table_formula_replace {e['table']}[{e['column']}]: no calculatedColumnFormula (skipped)")
            continue
        find, repl = esc(e["find"]), esc(e["replace"])
        inner = x[span[0]:span[1]]
        if inner.count(find) != 1:
            raise SystemExit(f"table {e['table']}[{e['column']}]: find occurs {inner.count(find)}x")
        x = x[:span[0]] + inner.replace(find, repl) + x[span[1]:]
        touched[p] = x
        log.append(f"table_formula_replace {e['table']}[{e['column']}]")

    for e in spec.get("set_value", []):
        p = need(sp, e["sheet"], "sheet")
        x = get(p)
        v = e["value"]
        if v is None:
            raise SystemExit(f"{e['sheet']}!{e['cell']}: value is null; give a number, text or true/false")
        m = find_cell(x, e["cell"])
        s = f' s="{e["style"]}"' if "style" in e else (style_of(m.group(0)) if m else "")
        if isinstance(v, bool):
            cx = f'<c r="{e["cell"]}"{s} t="b"><v>{1 if v else 0}</v></c>'
        elif isinstance(v, (int, float)):
            cx = f'<c r="{e["cell"]}"{s}><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
        else:
            keep = ' xml:space="preserve"' if v != v.strip() else ""
            cx = f'<c r="{e["cell"]}"{s} t="inlineStr"><is><t{keep}>{esc(v)}</t></is></c>'
        if m:
            if re.search(r"<f[ >/]", m.group(0)):
                raise SystemExit(f"{e['sheet']}!{e['cell']} holds a formula; refusing set_value")
            x = x[:m.start()] + cx + x[m.end():]
        else:
            x = insert_cell(x, e["cell"], cx)
        touched[p] = x
        log.append(f"set_value {e['sheet']}!{e['cell']} = {v!r}")

    for e in spec.get("set_formula", []):
        p = need(sp, e["sheet"], "sheet")
        x = get(p)
        f = lead_eq(e["formula"])
        m = find_cell(x, e["cell"])
        if m:
            g = re.search(r'<f\b[^>]*\bt="(?:shared|array)"[^>]*\bref="([^"]+)"', m.group(0))
            if g and ":" in g.group(1) and len(set(g.group(1).split(":"))) > 1:
                raise SystemExit(f"{e['sheet']}!{e['cell']} is the master of a formula group {g.group(1)}; "
                                 "use rewrite on the whole group")
        s = f' s="{e["style"]}"' if "style" in e else (style_of(m.group(0)) if m else "")
        cx = f'<c r="{e["cell"]}"{s}><f>{esc(f)}</f></c>'
        if m:
            x = x[:m.start()] + cx + x[m.end():]
        else:
            x = insert_cell(x, e["cell"], cx)
        touched[p] = x
        log.append(f"set_formula {e['sheet']}!{e['cell']}")

    for e in spec.get("set_dimension", []):
        p = need(sp, e["sheet"], "sheet")
        x = get(p)
        x2 = re.sub(r'<dimension ref="[^"]+"/>', f'<dimension ref="{e["ref"]}"/>', x, count=1)
        touched[p] = x2
        log.append(f"set_dimension {e['sheet']} {e['ref']}")

    removed = set()
    if spec.get("drop_calc_chain", True) and "xl/calcChain.xml" in parts:
        removed.add("xl/calcChain.xml")
        touched["xl/_rels/workbook.xml.rels"] = re.sub(
            r'<Relationship [^>]*Target="calcChain\.xml"[^>]*/>', "", get("xl/_rels/workbook.xml.rels"))
        touched["[Content_Types].xml"] = re.sub(
            r'<Override [^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", get("[Content_Types].xml"))
        log.append("calcChain dropped")

    if spec.get("full_calc_on_load", True):
        wb = get("xl/workbook.xml")
        if "fullCalcOnLoad" not in wb:
            wb2 = re.sub(r"<calcPr([^/>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
            if wb2 == wb:
                raise SystemExit("calcPr not found")
            touched["xl/workbook.xml"] = wb2
        log.append("fullCalcOnLoad=1")

    with zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            if info.filename in removed:
                continue
            data = touched[info.filename].encode("utf8") if info.filename in touched else parts[info.filename]
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    return log


def set_cached_values(src, dst, values):
    """values: {sheet: {ref: number|str|None}} -> rewrite <v> of existing formula cells (numbers only)."""
    if same_file(src, dst):
        raise SystemExit("SRC and DST are the same file; write the result to a new file")
    zin = zipfile.ZipFile(src)
    sp = sheet_paths(zin)
    touched = {}
    n = 0
    for sheet, refs in values.items():
        p = need(sp, sheet, "sheet")
        x = touched.get(p, zin.read(p).decode("utf8"))
        for ref, v in refs.items():
            m = find_cell(x, ref)
            if not m or "<f" not in m.group(0):
                continue
            cx = m.group(0)
            cx2 = re.sub(r'\st="(str|e|b|n)"', "", cx)
            if isinstance(v, bool):
                cx2 = cx2.replace(f'<c r="{ref}"', f'<c r="{ref}" t="b"', 1)
                val = f"<v>{1 if v else 0}</v>"
            elif isinstance(v, (int, float)):
                val = f"<v>{repr(float(v))}</v>"
            elif isinstance(v, str) and v in ERRORS:
                cx2 = cx2.replace(f'<c r="{ref}"', f'<c r="{ref}" t="e"', 1)
                val = f"<v>{esc(v)}</v>"
            elif isinstance(v, str):
                cx2 = cx2.replace(f'<c r="{ref}"', f'<c r="{ref}" t="str"', 1)
                val = f"<v>{esc(v)}</v>"
            else:
                val = ""
            cx2 = re.sub(r"<v>.*?</v>|<v/>", "", cx2, flags=re.S)
            if "</f>" in cx2:
                cx2 = cx2.replace("</f>", "</f>" + val, 1)
            else:  # shared-formula dependent: <f t="shared" si=".."/>
                fm = re.search(r"<f(?:\s[^>]*)?/>", cx2)
                assert fm, f"no formula tag in {ref}"
                cx2 = cx2[:fm.end()] + val + cx2[fm.end():]
            x = x[:m.start()] + cx2 + x[m.end():]
            n += 1
        touched[p] = x
    alldata = {i.filename: (touched[i.filename].encode("utf8") if i.filename in touched else zin.read(i.filename))
               for i in zin.infolist()}                      # read everything before DST is opened
    with zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = alldata[info.filename]
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    return n


def diff(a, b):
    """Cell-level diff of formulas and cached values across all sheets + list of changed zip parts."""
    za, zb = zipfile.ZipFile(a), zipfile.ZipFile(b)
    na, nb = set(za.namelist()), set(zb.namelist())
    changed_parts = sorted(n for n in na & nb if za.read(n) != zb.read(n))
    out = {"added_parts": sorted(nb - na), "removed_parts": sorted(na - nb), "changed_parts": changed_parts, "cells": []}
    spa = sheet_paths(za)
    for name, p in spa.items():
        if p not in changed_parts:
            continue
        xa, xb = za.read(p).decode("utf8"), zb.read(p).decode("utf8")
        cell_rx = r'<c r="([A-Z]+\d+)"(?:\s[^>]*?)?(?:/>|>(?:(?!<c[\s>]).)*?</c>)'      # cannot run into the next cell
        ca = {m.group(1): m.group(0) for m in re.finditer(cell_rx, xa, re.S)}
        cb = {m.group(1): m.group(0) for m in re.finditer(cell_rx, xb, re.S)}
        for ref in sorted(set(ca) | set(cb), key=lambda r: (int(re.sub(r"[A-Z]+", "", r)), len(r), r)):
            if ca.get(ref) != cb.get(ref):
                def parts_of(cx):
                    if cx is None:
                        return (None, None)
                    f = re.search(F_OPEN + "(.*?)</f>", cx, re.S)
                    v = re.search(r"<v>(.*?)</v>", cx, re.S) or re.search(r"<t>(.*?)</t>", cx, re.S)
                    return (html.unescape(f.group(1)) if f else None, html.unescape(v.group(1)) if v else None)
                fa, va = parts_of(ca.get(ref))
                fb, vb = parts_of(cb.get(ref))
                out["cells"].append({"sheet": name, "ref": ref, "f_changed": fa != fb, "v_changed": va != vb,
                                     "f_old": fa, "f_new": fb, "v_old": va, "v_new": vb})
    return out


if __name__ == "__main__":
    if len(sys.argv) == 2 and sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        raise SystemExit(0)
    need_args = {"apply": 5, "cached": 5, "diff": 4}
    if len(sys.argv) < 2 or sys.argv[1] not in need_args or len(sys.argv) != need_args[sys.argv[1]]:
        print(__doc__, file=sys.stderr)
        raise SystemExit(2)
    cmd = sys.argv[1]
    if cmd == "apply":
        print("\n".join(apply(sys.argv[2], sys.argv[3], json.load(open(sys.argv[4])))))
    elif cmd == "cached":
        print(set_cached_values(sys.argv[2], sys.argv[3], json.load(open(sys.argv[4]))), "formula cells updated")
    elif cmd == "diff":
        d = diff(sys.argv[2], sys.argv[3])
        print(json.dumps({k: v for k, v in d.items() if k != "cells"}, indent=1))
        import collections
        c = collections.Counter((x["sheet"], re.sub(r"\d+", "", x["ref"]), x["f_changed"], x["v_changed"]) for x in d["cells"])
        for k, v in sorted(c.items()):
            print(k, v)
