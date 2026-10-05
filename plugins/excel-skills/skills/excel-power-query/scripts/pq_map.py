#!/usr/bin/env python3
"""Map the Power Query layer of Excel workbooks without opening Excel.

  pq_map.py BOOK.xlsx                 one block per query: load target, sources, refs, last fill
  pq_map.py A.xlsx B.xlsx C.xlsx      same, plus which workbook reads which and the refresh order
  pq_map.py BOOK.xlsx --query NAME    print the M code of one query
  pq_map.py BOOK.xlsx --export DIR    write every query as NN_Name.pq
  pq_map.py BOOK.xlsx --json          the map as JSON

Read-only: workbooks are never modified. Standard library only.
"""
import argparse, base64, io, json, os, posixpath, re, struct, sys, zipfile, zlib
from urllib.parse import unquote
import xml.etree.ElementTree as ET

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
SHARED = re.compile(r'(?m)^shared\s+((?:#"(?:[^"]|"")*"|[^=\s]+))\s*=')
SOURCE = re.compile(r'\b((?:File|Web)\.Contents|Folder\.(?:Files|Contents)|SharePoint\.(?:Files|Contents|Tables)'
                    r'|OData\.Feed|Odbc\.(?:DataSource|Query)|OleDb\.(?:DataSource|Query)|Excel\.CurrentWorkbook'
                    r'|\w+\.Databases?)\s*\(((?:"(?:[^"]|"")*"|[^()"])*)')
TOKEN = re.compile(r'#?"(?:[^"]|"")*"|//[^\n]*|/\*.*?\*/', re.S)   # strings, quoted identifiers, comments
PICK = re.compile(r'\[\s*(?:Item|Name)\s*=\s*"([^"]+)"\s*(?:,\s*Kind\s*=\s*"([^"]+)"\s*)?\]')


def read_mashup(z):
    """Binary DataMashup payload, or None. The part is UTF-16, so a byte search for the tag fails."""
    for n in sorted(z.namelist()):
        if not re.fullmatch(r'customXml/item\d+\.xml', n):
            continue
        raw = z.read(n)
        for enc in ('utf-16', 'utf-8-sig'):
            try:
                text = raw.decode(enc)
            except UnicodeError:
                continue
            m = re.search(r'<DataMashup[^>]*>([^<]+)</DataMashup>', text)
            if m:
                return base64.b64decode(m.group(1))
    return None


def scan_local_headers(data):
    """Fallback for packages without a usable central directory."""
    parts, pos = {}, 0
    while (pos := data.find(b'PK\x03\x04', pos)) >= 0:
        method, = struct.unpack_from('<H', data, pos + 8)
        csize, = struct.unpack_from('<I', data, pos + 18)
        nlen, xlen = struct.unpack_from('<HH', data, pos + 26)
        name = data[pos + 30:pos + 30 + nlen].decode('utf-8', 'replace')
        start = pos + 30 + nlen + xlen
        if method == 8:
            d = zlib.decompressobj(-15)
            parts[name] = d.decompress(data[start:])
            pos = len(data) - len(d.unused_data)
        else:
            parts[name] = data[start:start + csize]
            pos = start + csize
    return parts


def unpack(blob):
    """-> ({package part: bytes}, metadata xml). Layout: version, package, permissions, metadata."""
    plen, = struct.unpack_from('<I', blob, 4)
    try:
        with zipfile.ZipFile(io.BytesIO(blob[8:8 + plen])) as z:
            parts = {n: z.read(n) for n in z.namelist()}
    except zipfile.BadZipFile:
        parts = scan_local_headers(blob)
    meta = ''
    try:
        off = 8 + plen
        off += 4 + struct.unpack_from('<I', blob, off)[0]        # permissions
        xml_len, = struct.unpack_from('<I', blob, off + 8)       # metadata: length, version, xml length
        meta = blob[off + 12:off + 12 + xml_len].decode('utf-8-sig', 'replace')
    except struct.error:
        pass
    return parts, meta


def split_queries(parts):
    out = []
    for name in sorted(parts):
        if re.fullmatch(r'Formulas/Section\d*\.m', name):
            text = parts[name].decode('utf-8-sig')
            ms = list(SHARED.finditer(text))
            for i, m in enumerate(ms):
                end = ms[i + 1].start() if i + 1 < len(ms) else len(text)
                tok = m.group(1)
                qname = tok[2:-1].replace('""', '"') if tok.startswith('#"') else tok
                out.append((qname, text[m.start():end].rstrip()))
    return out


def fill_info(meta_xml):
    """{query: {entry: value}} from the mashup metadata (last fill time, row count, status)."""
    info = {}
    if not meta_xml:
        return info
    try:
        root = ET.fromstring(meta_xml)
    except ET.ParseError:
        return info
    local = lambda e: e.tag.rsplit('}', 1)[-1]
    for item in root.iter():
        if local(item) != 'Item':
            continue
        path = next((e.text for e in item.iter() if local(e) == 'ItemPath'), '') or ''
        m = re.fullmatch(r'Section\d*/([^/]+)', path)
        if not m:
            continue
        entries = {e.get('Type'): (e.get('Value') or '')[1:] for e in item.iter() if local(e) == 'Entry'}
        info.setdefault(unquote(m.group(1)), {}).update(entries)
    return info


def rels(z, part):
    """{rId: (relationship kind, resolved part)}"""
    d, b = posixpath.split(part)
    rp = f'{d}/_rels/{b}.rels'
    out = {}
    if rp in z.namelist():
        for r in ET.fromstring(z.read(rp)):
            t = r.get('Target') or ''
            tgt = t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join(d, t))
            out[r.get('Id')] = ((r.get('Type') or '').rsplit('/', 1)[-1], tgt)
    return out


def load_targets(z):
    """{query: [where it lands]} via connections -> queryTables -> tables -> sheets."""
    conns = {}
    if 'xl/connections.xml' in z.namelist():
        for c in ET.fromstring(z.read('xl/connections.xml')).findall('m:connection', NS):
            db = c.find('m:dbPr', NS)
            m = re.search(r'Location=("[^"]*"|[^;]*)', db.get('connection', '') if db is not None else '')
            if m:
                conns[c.get('id')] = m.group(1).strip('"')
    targets = {}
    wrels = rels(z, 'xl/workbook.xml')
    for sh in ET.fromstring(z.read('xl/workbook.xml')).iterfind('m:sheets/m:sheet', NS):
        kind, part = wrels.get(sh.get(RID), ('', ''))
        if kind != 'worksheet':
            continue
        sheet = (sh.get('name') or '') + (f" [{sh.get('state')}]" if sh.get('state') not in (None, 'visible') else '')
        for kind, tgt in rels(z, part).values():
            if kind == 'table':
                t = ET.fromstring(z.read(tgt))
                col_list = list(t.iter('{%s}tableColumn' % NS['m']))
                cols = {c.get('id'): c for c in col_list}
                formula = '{%s}calculatedColumnFormula' % NS['m']
                filtered = [col_list[int(f.get('colId'))].get('name') for f in t.iterfind('m:autoFilter/m:filterColumn', NS)
                            if (f.get('colId') or '').isdigit() and int(f.get('colId')) < len(col_list)]
                for kind2, qt in rels(z, tgt).values():
                    if kind2 != 'queryTable':
                        continue
                    q = ET.fromstring(z.read(qt))
                    fields = list(q.iter('{%s}queryTableField' % NS['m']))
                    side = [cols[f.get('tableColumnId')] for f in fields
                            if f.get('dataBound') == '0' and f.get('tableColumnId') in cols]
                    excel_side = [c.get('name') for c in side]
                    targets.setdefault(conns.get(q.get('connectionId')), []).append({
                        'sheet': sheet, 'table': t.get('displayName') or t.get('name'), 'range': t.get('ref'),
                        'query_columns': len(fields) - len(excel_side), 'excel_side_columns': excel_side,
                        'typed_columns': [c.get('name') for c in side if c.find(formula) is None],
                        'filtered_on': filtered})
            elif kind == 'queryTable':
                q = ET.fromstring(z.read(tgt))
                targets.setdefault(conns.get(q.get('connectionId')), []).append({
                    'sheet': sheet, 'table': None, 'range': q.get('name'),
                    'query_columns': len(list(q.iter('{%s}queryTableField' % NS['m']))), 'excel_side_columns': [],
                    'typed_columns': [], 'filtered_on': []})
    targets.pop(None, None)
    return targets


def references(body, names, me):
    quoted = set()

    def drop(m):
        if m.group(0).startswith('#"'):
            quoted.add(m.group(0)[2:-1].replace('""', '"'))
        return ' '
    bare = TOKEN.sub(drop, body)
    idents = set(re.findall(r'(?<![\w.\[])[A-Za-z_][\w.]*', bare))     # `[Col]` is a field, not a query
    return [n for n in names if n != me and (n in quoted or n in idents)]


def build_map(path):
    with zipfile.ZipFile(path) as z:
        blob = read_mashup(z)
        if blob is None:
            return {'workbook': path, 'queries': [], 'note': 'no DataMashup part: this workbook has no Power Query'}
        parts, meta = unpack(blob)
        targets = load_targets(z)
    queries, fills = split_queries(parts), fill_info(meta)
    names = [n for n, _ in queries]
    out = []
    for i, (name, body) in enumerate(queries, 1):
        f = fills.get(name, {})
        out.append({
            'n': i, 'name': name, 'm': body,
            'loads_to': targets.get(name, []),
            'sources': [f'{fn}({arg.strip()[:200]})' for fn, arg in SOURCE.findall(body)],
            'picks': [f'{kind or "item"} "{item}"' for item, kind in PICK.findall(body)],
            'uses': references(body, names, name),
            'last_fill': {k: f[k] for k in ('FillLastUpdated', 'FillCount', 'FillStatus', 'FillErrorCount',
                                            'FillErrorCode', 'FillObjectType') if k in f},
        })
    for q in out:
        q['used_by'] = [o['name'] for o in out if q['name'] in o['uses']]
    return {'workbook': path, 'queries': out}


def print_map(wb):
    qs = wb['queries']
    print(f"\n== {wb['workbook']}: {len(qs)} queries, {sum(bool(q['loads_to']) for q in qs)} loaded to a sheet")
    if wb.get('note'):
        print('   ' + wb['note'])
    for q in qs:
        print(f"\n[{q['n']}] {q['name']}")
        for t in q['loads_to']:
            where = f"{t['sheet']}!{t['table']}" if t['table'] else f"{t['sheet']} (plain range {t['range']})"
            print(f"    loads to : {where} {t['range'] if t['table'] else ''} | {t['query_columns']} query columns")
            if t['excel_side_columns']:
                print(f"    excel-side columns (not from the query): {', '.join(t['excel_side_columns'])}")
            if t['typed_columns']:
                print(f"    of these, typed by hand (no column formula): {', '.join(t['typed_columns'])}")
            if t['filtered_on']:
                print(f"    TABLE FILTER saved on: {', '.join(t['filtered_on'])}. Rows can be loaded and still hidden.")
        if not q['loads_to']:
            print('    loads to : nothing on a sheet (connection only, or data model)')
        for s in q['sources']:
            print(f'    source   : {s}')
        if q['picks']:
            print(f"    picks    : {'; '.join(dict.fromkeys(q['picks']))}")
        if q['uses']:
            print(f"    uses     : {', '.join(q['uses'])}")
        if q['used_by']:
            print(f"    used by  : {', '.join(q['used_by'])}")
        if q['last_fill']:
            print('    last fill: ' + ', '.join(f'{k[4:]}={v}' for k, v in q['last_fill'].items()))


def print_chain(maps):
    """Which workbook reads which (by file name inside source expressions) and a safe refresh order."""
    base = {os.path.basename(w['workbook']).lower(): w['workbook'] for w in maps}
    deps = {w['workbook']: set() for w in maps}
    print('\n== Links between the workbooks given')
    for w in maps:
        for q in w['queries']:
            for s in q['sources']:
                hit = [p for b, p in base.items() if b in unquote(s).lower() and p != w['workbook']]
                for p in hit:
                    deps[w['workbook']].add(p)
                    print(f"   {os.path.basename(w['workbook'])} / {q['name']}  reads  {os.path.basename(p)}"
                          f"  ({'; '.join(dict.fromkeys(q['picks'])) or 'see M'})")
                if not hit and s.startswith('File.Contents') and '"' not in s:
                    print(f"   {os.path.basename(w['workbook'])} / {q['name']}: path is an expression, resolve by hand: {s}")
    order, left = [], dict(deps)
    while left:
        ready = sorted(p for p, d in left.items() if not (d & left.keys()))
        if not ready:
            print('   cycle between: ' + ', '.join(map(os.path.basename, left)))
            break
        order += ready
        for p in ready:
            del left[p]
    print('   refresh order (upstream first; save each one before the next): '
          + ' -> '.join(map(os.path.basename, order)))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('workbooks', nargs='+')
    ap.add_argument('--query', help='print the M code of this query')
    ap.add_argument('--export', metavar='DIR', help='write every query as a .pq file')
    ap.add_argument('--json', action='store_true')
    a = ap.parse_args()
    if hasattr(sys.stdout, 'reconfigure'):          # query and column names are often not ASCII
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    maps = [build_map(p) for p in a.workbooks]
    if a.query:
        hits = [(w, q) for w in maps for q in w['queries'] if q['name'].lower() == a.query.lower()]
        if not hits:
            sys.exit(f"no query named {a.query!r}; have: {', '.join(q['name'] for w in maps for q in w['queries'])}")
        print('\n\n'.join((f"// {w['workbook']}\n" if len(maps) > 1 else '') + q['m'] for w, q in hits))
    elif a.export:
        os.makedirs(a.export, exist_ok=True)
        for w in maps:
            for q in w['queries']:
                safe = re.sub(r'\s+', ' ', re.sub(r'[\\/:*?"<>|]', '_', q['name'])).strip()
                body = q['m'] if q['m'].endswith(';') else q['m'] + ';'
                with open(os.path.join(a.export, f"{q['n']:02d}_{safe}.pq"), 'w', encoding='utf-8') as f:
                    f.write('section Section1;\n\n' + body + '\n')
            print(f"{len(w['queries'])} queries from {w['workbook']} -> {a.export}")
    elif a.json:
        json.dump(maps, sys.stdout, ensure_ascii=False, indent=1)
    else:
        for w in maps:
            print_map(w)
        if len(maps) > 1:
            print_chain(maps)


if __name__ == '__main__':
    main()
