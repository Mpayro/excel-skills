#!/usr/bin/env python3
"""Structural check of an .xlsx / .xlsm package: the defects that make Excel say "we found a problem
with some content" after a file was edited outside Excel.

  xlsx_check.py BOOK.xlsx              check one workbook
  xlsx_check.py BASE.xlsx NEW.xlsx     check NEW, and list the zip parts that differ from BASE

Checks: zip integrity, XML well-formedness, content types (a part needs its own Override; only customXml
items may rely on the xml Default), relationship targets and duplicate relationship ids, r:id references,
row and cell order, cell types and numeric values, style indexes, formula text (no leading "=", not empty),
shared-string indexes, shared-formula groups, the calculation chain (every entry is a formula cell),
tables (header cells equal to column names, ids, ranges, overlap, unique names), merged cells (no overlap),
unique sheet names and defined names, notes against their VML shapes. It also counts formula cells that have
no stored result. It does not check what the numbers mean, charts, pivot tables, conditional formats or
validations, and it cannot promise that Excel opens the file without a message.
Read-only, standard library only.
Exit code: 0 no FAIL line | 1 at least one FAIL line (also a file that is not a readable zip) | 2 usage error.
"""
import html, os, posixpath, re, sys, zipfile
import xml.etree.ElementTree as ET

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
ROW = r'<row\b[^>]*?\br="%d"(?:\s[^>]*?)?(?:/>|>(?:(?!<row[\s>]).)*?</row>)'
CELL = r'<c\b[^>]*?\br="%s"(?:\s[^>]*?)?(?:/>|>(?:(?!<c[\s>]).)*?</c>)'


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def unesc(s):
    """XML entities and Excel's _x000A_ escapes."""
    return re.sub(r"_x([0-9A-Fa-f]{4})_", lambda m: chr(int(m.group(1), 16)), html.unescape(s or ""))


def split_ref(ref):
    a, _, b = (ref or "").partition(":")
    m1, m2 = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", a), re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", b or a)
    if not (m1 and m2):
        raise ValueError(f"not a cell range: {ref!r}")
    return col_num(m1.group(1)), int(m1.group(2)), col_num(m2.group(1)), int(m2.group(2))


def rels_of(z, part):
    """[(id, kind, target part, external)] for a part ('' = package root)."""
    d, b = posixpath.split(part)
    rp = posixpath.join(d, "_rels", b + ".rels")
    out = []
    if rp in z.namelist():
        try:
            rels = list(ET.fromstring(z.read(rp)))
        except ET.ParseError:
            return out                      # reported as "not well-formed" by check()
        for r in rels:
            t = (r.get("Target") or "").split("#")[0]
            ext = r.get("TargetMode") == "External"
            tgt = t if ext else t[1:] if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
            out.append((r.get("Id"), (r.get("Type") or "").rsplit("/", 1)[-1], tgt, ext))
    return out


def shared_strings(z, names):
    if "xl/sharedStrings.xml" not in names:
        return []
    out = []
    for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(M + "si"):
        out.append("".join(t.text or "" for el in si for t in ([el] if el.tag == M + "t" else el.iter(M + "t") if el.tag == M + "r" else [])))
    return out


def cell_text(cx, sst):
    t = re.search(r'\bt="(\w+)"', cx.split(">", 1)[0])
    v = re.search(r"<v>(.*?)</v>", cx, re.S)
    if t and t.group(1) == "s" and v:
        i = int(v.group(1))
        return sst[i] if i < len(sst) else None
    if t and t.group(1) == "inlineStr":
        return unesc("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", cx, re.S)))
    return unesc(v.group(1)) if v else ""


def check(path, base=None):
    out = []
    fail = lambda msg: out.append("FAIL  " + msg)
    note = lambda msg: out.append("note  " + msg)
    try:
        z = zipfile.ZipFile(path)
        names = set(z.namelist())
        bad_member = z.testzip()
    except (zipfile.BadZipFile, EOFError, OSError, NotImplementedError) as e:
        fail(f"zip: not a readable zip package ({e})")
        return out
    if bad_member:
        fail(f"zip: damaged member {bad_member}")
    for need in ("[Content_Types].xml", "xl/workbook.xml"):
        if need not in names:
            fail(f"package: {need} is missing")
    if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
        return out

    for n in sorted(names):
        if n.endswith((".xml", ".rels", ".vml")):
            try:
                ET.fromstring(z.read(n))
            except ET.ParseError as e:
                fail(f"{n}: not well-formed XML ({e})")

    try:
        ct = ET.fromstring(z.read("[Content_Types].xml"))
        ET.fromstring(z.read("xl/workbook.xml"))
    except ET.ParseError:
        return out                          # the well-formedness loop above already printed the FAIL line
    defaults = {e.get("Extension", "").lower() for e in ct if e.tag.endswith("Default")}
    overrides = {e.get("PartName", "").lstrip("/") for e in ct if e.tag.endswith("Override")}
    for n in sorted(names - {"[Content_Types].xml"}):
        if not n.endswith("/") and n not in overrides and n.rsplit(".", 1)[-1].lower() not in defaults:
            fail(f"{n}: no content type")
        elif n.endswith(".xml") and n not in overrides and not re.fullmatch(r"customXml/item\d+\.xml", n):
            fail(f"{n}: no Override in [Content_Types].xml (the xml Default, application/xml, is not its type)")
    for n in sorted(overrides - names):
        fail(f"[Content_Types].xml: Override for a part that does not exist: {n}")

    for rp in sorted(n for n in names if n.endswith(".rels")):
        try:
            ids = [r.get("Id") for r in ET.fromstring(z.read(rp))]
        except ET.ParseError:
            ids = []
        if len(ids) != len(set(ids)):
            fail(f"{rp}: a relationship id is used twice")
        d, b = posixpath.split(rp)
        src = posixpath.join(posixpath.dirname(d), b[:-5])
        for rid, kind, tgt, ext in rels_of(z, src):
            if not ext and tgt not in names:
                fail(f"{rp}: {rid} ({kind}) points to a missing part {tgt}")

    wbx = z.read("xl/workbook.xml").decode("utf8")
    wrels = {rid: (kind, tgt) for rid, kind, tgt, _ in rels_of(z, "xl/workbook.xml")}
    sst = shared_strings(z, names)
    calc = re.search(r"<calcPr[^>]*>", wbx)
    note(f"calculation: {calc.group(0) if calc else 'no calcPr'} | calcChain {'present' if 'xl/calcChain.xml' in names else 'absent'}")

    # unique names: sheets, defined names (per scope), tables
    sheet_els = re.findall(r"<sheet\b[^>]*>", wbx)
    attr = lambda el, k: (re.search(r'\b%s="([^"]*)"' % k, el) or [None, None])[1]
    for label, vals in (("sheet name", [unesc(attr(e, "name") or "").casefold() for e in sheet_els]),
                        ("sheetId", [attr(e, "sheetId") for e in sheet_els]),
                        ("defined name", [((unesc(attr(e, "name") or "")).casefold(), attr(e, "localSheetId"))
                                          for e in re.findall(r"<definedName\b[^>]*>", wbx)])):
        dup = sorted({str(v) for v in vals if vals.count(v) > 1})
        if dup:
            fail(f"workbook.xml: {label} is used twice: {dup[:3]}")
    tnames, tids = [], []
    for tp in sorted(n for n in names if re.fullmatch(r"xl/tables/[^/]+\.xml", n)):
        head = re.search(r"<table\b[^>]*>", z.read(tp).decode("utf8", "replace"))
        if head:
            tnames.append(unesc(attr(head.group(0), "displayName") or attr(head.group(0), "name") or "").casefold())
            tids.append(attr(head.group(0), "id"))
    for label, vals in (("table name", tnames), ("table id", tids)):
        dup = sorted({v for v in vals if vals.count(v) > 1})
        if dup:
            fail(f"tables: {label} is used twice: {dup[:3]}")
    n_xf = None
    if "xl/styles.xml" in names:
        try:
            cx = ET.fromstring(z.read("xl/styles.xml")).find(M + "cellXfs")
            n_xf = len(cx) if cx is not None else None
        except ET.ParseError:
            pass
    sheet_id_of = {}                        # part name -> sheetId
    formula_cells = {}                      # sheetId -> set of cell references that hold a formula

    vml_ids = {}
    tables = 0
    for name, rid in re.findall(r'<sheet\b[^>]*?\bname="([^"]*)"[^>]*?\br:id="([^"]+)"', wbx):
        name = unesc(name)
        kind, part = wrels.get(rid, ("", ""))
        if kind != "worksheet" or part not in names:
            if kind == "worksheet":
                fail(f"sheet {name!r}: part {part} is missing")
            continue
        x = z.read(part).decode("utf8")
        srels = {r: (k, t) for r, k, t, _ in rels_of(z, part)}
        for r in set(re.findall(r'\br:id="([^"]+)"', x)) - set(srels):
            fail(f"sheet {name!r}: r:id {r} is not in its relationships")

        # rows ascending, cells in their own row and in column order
        data = re.search(r"<sheetData\b.*?</sheetData>", x, re.S)
        data = data.group(0) if data else ""
        if len(re.findall(r"<row\b", data)) != len(re.findall(r'<row\b[^>]*?\br="\d+"', data)):
            note(f"sheet {name!r}: rows without r attribute, order not checked")
        else:
            last_row = last_col = 0
            bad = 0
            for m in re.finditer(r'<row\b[^>]*?\br="(\d+)"|<c\b[^>]*?\br="([A-Z]+)(\d+)"', data):
                if m.group(1):
                    r = int(m.group(1))
                    if r <= last_row and bad < 3:
                        bad += 1
                        fail(f"sheet {name!r}: row {r} is out of order or repeated")
                    last_row, last_col = r, 0
                else:
                    c, r = col_num(m.group(2)), int(m.group(3))
                    if (r != last_row or c <= last_col) and bad < 3:
                        bad += 1
                        fail(f"sheet {name!r}: cell {m.group(2)}{r} is out of order, repeated or in the wrong row")
                    last_col = c

        idx = [int(i) for i in re.findall(r'<c\b[^>]*\bt="s"[^>]*>(?:<f\b[^>]*/>|<f\b[^>]*>.*?</f>)?\s*<v>(\d+)</v>', data, re.S)]
        if idx and max(idx) >= len(sst):
            fail(f"sheet {name!r}: shared string index {max(idx)} but the table has {len(sst)} strings")

        # cell types, numeric values, style indexes, formula text
        sid = next((attr(e, "sheetId") for e in sheet_els if attr(e, "name") is not None and attr(e, "r:id") == rid), None)
        fset = formula_cells.setdefault(sid, set())
        bad = 0
        for cm in re.finditer(r"<c\b([^>]*?)(?:/>|>((?:(?!<c[\s>]).)*?)</c>)", data, re.S):
            ca, body = cm.group(1), cm.group(2) or ""
            ref, ty = attr(ca, "r"), attr(ca, "t")
            sty, val = attr(ca, "s"), re.search(r"<v>(.*?)</v>", body, re.S)
            if "<f" in body and ref:
                fset.add(ref)
            msg = None
            if ty is not None and ty not in ("b", "d", "e", "inlineStr", "n", "s", "str"):
                msg = f"unknown cell type t={ty!r}"
            elif ty in (None, "n") and val:
                try:
                    float(val.group(1))
                except ValueError:
                    msg = f"number expected, found {val.group(1)[:20]!r}"
            elif ty == "b" and val and val.group(1) not in ("0", "1"):
                msg = f"boolean expected, found {val.group(1)[:20]!r}"
            if msg is None and sty is not None and (not sty.isdigit() or (n_xf is not None and int(sty) >= n_xf)):
                msg = f"style index s={sty!r} does not exist (the workbook has {n_xf} cell formats)"
            if msg and bad < 3:
                bad += 1
                fail(f"sheet {name!r}: cell {ref}: {msg}")
        if re.search(r"<f\b[^>]*?(?<!/)>\s*=", data):
            fail(f"sheet {name!r}: a formula starts with '=' (the file stores formulas without it)")
        if re.search(r"<f(?:\s[^>]*?)?(?<!/)>\s*</f>|<f\s*/>", data):
            fail(f"sheet {name!r}: a formula element is empty")
        for rm in re.finditer(r"<row\b[^>]*>", data):
            rs = attr(rm.group(0), "s")
            if rs is not None and (not rs.isdigit() or (n_xf is not None and int(rs) >= n_xf)):
                fail(f"sheet {name!r}: row style s={rs!r} does not exist")
                break

        dm = re.search(r'<dimension\b[^>]*?\bref="([^"]*)"', x)
        if dm:
            try:
                split_ref(dm.group(1))
            except ValueError:
                fail(f"sheet {name!r}: dimension {dm.group(1)!r} is not a range")

        # merged cells must not overlap
        merges = []
        for mm in re.findall(r"<mergeCell\b[^>]*?\bref=\"([^\"]+)\"", x):
            try:
                merges.append((mm,) + split_ref(mm))
            except ValueError:
                fail(f"sheet {name!r}: merged range {mm!r} is not a range")
        if len(merges) <= 3000:
            for i, a in enumerate(merges):
                if any(not (a[3] < b[1] or b[3] < a[1] or a[4] < b[2] or b[4] < a[2]) for b in merges[i + 1:]):
                    fail(f"sheet {name!r}: merged ranges overlap ({a[0]})")
                    break

        unstored = sum("<f" in chunk and "<v" not in chunk for chunk in data.split("<c ")[1:])
        if unstored:
            note(f"sheet {name!r}: {unstored} formula cells have no stored result. Power Query and other readers "
                 "see them as empty until Excel opens and saves the file")

        masters, deps = set(), set()
        for attrs in re.findall(r"<f\b([^>]*?)/?>", data):
            if 't="shared"' in attrs:
                si = re.search(r'\bsi="(\d+)"', attrs)
                (masters if "ref=" in attrs else deps).add(si.group(1) if si else "?")
        if deps - masters:
            fail(f"sheet {name!r}: shared formula groups without a master: si {sorted(deps - masters)[:5]}")

        # tables: columns, and header cells equal to column names
        sheet_tables = []
        for r, (k, tpart) in srels.items():
            if k != "table" or tpart not in names:
                continue
            tables += 1
            try:
                t = ET.fromstring(z.read(tpart))
            except ET.ParseError:
                continue
            tname, ref = t.get("displayName") or t.get("name"), t.get("ref")
            cols = list(t.iter(M + "tableColumn"))
            try:
                c1, r1, c2, r2 = split_ref(ref)
            except ValueError:
                fail(f"table {tname}: ref {ref!r} is not a range")
                continue
            for oname, (a1, b1, a2, b2) in sheet_tables:
                if not (a2 < c1 or c2 < a1 or b2 < r1 or r2 < b1):
                    fail(f"table {tname} overlaps table {oname} on sheet {name!r}")
            sheet_tables.append((tname, (c1, r1, c2, r2)))
            tc = t.find(M + "tableColumns")
            if tc is not None and tc.get("count", str(len(cols))) != str(len(cols)):
                fail(f"table {tname}: count={tc.get('count')} but {len(cols)} columns")
            if len(cols) != c2 - c1 + 1:
                fail(f"table {tname}: {len(cols)} columns but ref {ref} is {c2 - c1 + 1} wide")
            if len({c.get("id") for c in cols}) != len(cols):
                fail(f"table {tname}: column ids repeat")
            if len({unesc(c.get("name")).casefold() for c in cols}) != len(cols):
                fail(f"table {tname}: column names repeat")
            af = t.find(M + "autoFilter")
            if af is not None and af.get("ref"):
                a1, b1, a2, b2 = split_ref(af.get("ref"))
                if a1 < c1 or a2 > c2 or b1 < r1 or b2 > r2:
                    fail(f"table {tname}: autoFilter {af.get('ref')} is outside ref {ref}")
            if t.get("headerRowCount") != "0":
                rm = re.search(ROW % r1, data, re.S)
                rowx = rm.group(0) if rm else ""
                for i, c in enumerate(cols[:c2 - c1 + 1]):
                    cm = re.search(CELL % f"{col_letters(c1 + i)}{r1}", rowx, re.S)
                    got = cell_text(cm.group(0), sst) if cm else None
                    if got != unesc(c.get("name")):
                        fail(f"table {tname}: header {col_letters(c1 + i)}{r1} is {got!r} but the column is named {unesc(c.get('name'))!r}")
                        break

        # notes: every comment has its VML shape on the same cell
        cpart = next((t for k, t in srels.values() if k == "comments"), None)
        vpart = next((t for k, t in srels.values() if k == "vmlDrawing"), None)
        if cpart in names:
            refs = sorted(re.findall(r'<comment\b[^>]*?\bref="([A-Z]+\d+)"', z.read(cpart).decode("utf8")))
            shapes = []
            if vpart in names:
                v = z.read(vpart).decode("utf8", "replace")
                for block in re.findall(r'<x:ClientData ObjectType="Note">.*?</x:ClientData>', v, re.S):
                    rr, cc = re.search(r"<x:Row>(\d+)</x:Row>", block), re.search(r"<x:Column>(\d+)</x:Column>", block)
                    if rr and cc:
                        shapes.append(f"{col_letters(int(cc.group(1)) + 1)}{int(rr.group(1)) + 1}")
            if refs != sorted(shapes):
                fail(f"sheet {name!r}: {len(refs)} notes but {len(shapes)} note shapes, or on different cells")
            if "<legacyDrawing" not in x:
                fail(f"sheet {name!r}: has notes but no legacyDrawing element")
        if vpart in names:
            for i in re.findall(r'<v:shape\b[^>]*?\bid="([^"]+)"', z.read(vpart).decode("utf8", "replace")):
                vml_ids.setdefault(i, []).append(vpart)
    if "xl/calcChain.xml" in names:
        # ECMA-376 18.6.1 <c i="..."> is the sheetId; an entry without i belongs to the sheet of the entry before it.
        # Every entry must name a formula cell of that sheet.
        sid, missing = None, []
        for em in re.finditer(r"<c\b([^>]*?)/?>", z.read("xl/calcChain.xml").decode("utf8", "replace")):
            sid = attr(em.group(1), "i") or sid
            ref = attr(em.group(1), "r")
            if sid not in formula_cells:
                missing.append(f"{ref} (unknown sheetId {sid})")
            elif ref not in formula_cells[sid]:
                missing.append(f"{ref} (sheetId {sid})")
        if missing:
            fail(f"calcChain names {len(missing)} cells that hold no formula, for example {missing[:3]}")
    for i, where in vml_ids.items():
        if len(where) != len(set(where)):
            fail(f"VML shape id {i} repeats inside {where[0]}")
        elif len(where) > 1:
            note(f"VML shape id {i} is used in {len(where)} drawings")
    note(f"{len(names)} parts, {tables} tables, {len(sst)} shared strings")

    if base:
        zb = zipfile.ZipFile(base)
        nb = set(zb.namelist())
        note(f"against base: added {sorted(names - nb) or 'none'}, removed {sorted(nb - names) or 'none'}, "
             f"changed {sorted(n for n in names & nb if z.read(n) != zb.read(n)) or 'none'}")
    return out


if __name__ == "__main__":
    if len(sys.argv) == 2 and sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    if len(sys.argv) not in (2, 3):
        print(__doc__, file=sys.stderr)
        sys.exit(2)
    for f in sys.argv[1:]:
        if not os.path.isfile(f):
            print(f"no such file: {f}", file=sys.stderr)
            sys.exit(2)
    lines = check(sys.argv[-1], sys.argv[1] if len(sys.argv) == 3 else None)
    print("\n".join(lines))
    fails = sum(line.startswith("FAIL") for line in lines)
    print(f"{fails} structural problems" if fails else "structure ok")
    sys.exit(1 if fails else 0)
