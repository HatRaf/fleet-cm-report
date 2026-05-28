"""
read_excel.py — Safe Excel reader using zipfile + ElementTree.
Works on files that break openpyxl due to invalid XML color values.

Usage:
    python3 read_excel.py path/to/file.xlsx

Or import and call read_xlsx(path) which returns {sheet_name: [[row], ...]}
"""
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        # Shared strings
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.parse(z.open('xl/sharedStrings.xml'))
            for si in tree.getroot().findall('.//s:si', NS):
                txt = ''.join(t.text or '' for t in si.findall('.//s:t', NS))
                shared.append(txt)

        # Sheet names + relationship IDs
        wb = ET.parse(z.open('xl/workbook.xml'))
        wb_ns = {'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        sheets = []
        for s in wb.getroot().findall('.//w:sheet', wb_ns):
            rid = (s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                   or s.get('r:id') or '')
            sheets.append((s.get('name', ''), rid))

        # Relationships
        rels = {}
        if 'xl/_rels/workbook.xml.rels' in z.namelist():
            rt = ET.parse(z.open('xl/_rels/workbook.xml.rels'))
            for r in rt.getroot():
                rels[r.get('Id', '')] = r.get('Target', '')

        results = {}
        for sname, sid in sheets:
            target = rels.get(sid, '')
            spath = ('xl/' + target) if not target.startswith('xl/') else target
            if spath not in z.namelist():
                continue
            ws = ET.parse(z.open(spath))
            rows = []
            for row in ws.getroot().findall('.//s:row', NS):
                cells = []
                for c in row.findall('s:c', NS):
                    t = c.get('t', '')
                    v = c.find('s:v', NS)
                    if t == 's' and v is not None:
                        idx = int(v.text)
                        cells.append(shared[idx] if idx < len(shared) else '')
                    elif t == 'inlineStr':
                        cells.append(''.join(x.text or '' for x in c.findall('.//s:t', NS)))
                    else:
                        cells.append(v.text if v is not None else '')
                rows.append(cells)
            results[sname] = rows
        return results


if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        print("Usage: python3 read_excel.py <file.xlsx>")
        sys.exit(1)
    data = read_xlsx(path)
    for sheet, rows in data.items():
        print(f"\nSheet: {sheet} ({len(rows)} rows)")
        for r in rows:
            if any(c for c in r):
                print(f"  {r}")
