"""
report_builder.py  --  Fleet CM Report HTML generator
Approach: Use the reference HTML as a template (preserving all embedded assets,
CSS, and deck-stage JS) and inject data from Excel files.

Usage:
    python3 report_builder.py \
        --summary  "Hat_CM_Summary_EXAMPLE_2026-05.xlsx" \
        --extra    "Hat_CM_Summary_Extra_EXAMPLE_2026-05.xlsx" \
        [--summary-prev "Hat_CM_Summary_EXAMPLE_2026-04.xlsx"] \
        --template "path/to/Example_Fleet_CM_Report_..._Standalone.html" \
        --fleet    "EXAMPLE" \
        --out      "Fleet_CM_Report_EXAMPLE_2026-05.html"

After generating, convert to PDF:
    node scripts/print_pdf.js output.html output.pdf
"""

import argparse
import sys
import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET
import datetime
import html as html_mod

# Report signature block — supplied at runtime via env vars so the public repo
# carries no identifying info. Set HAT_CERTS / HAT_SUPERVISOR / HAT_EMAIL to real values.
SIG_CERTS      = os.environ.get('HAT_CERTS', 'Certifications on file')
SIG_SUPERVISOR = os.environ.get('HAT_SUPERVISOR', 'Report supervisor')
SIG_EMAIL      = os.environ.get('HAT_EMAIL', 'contact@example.com')

# ── Excel reader (zipfile + ElementTree, avoids openpyxl breakage) ────────────
NS  = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
WNS = {'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.parse(z.open('xl/sharedStrings.xml'))
            for si in tree.getroot().findall('.//s:si', NS):
                shared.append(''.join(t.text or '' for t in si.findall('.//s:t', NS)))
        wb = ET.parse(z.open('xl/workbook.xml'))
        sheets = []
        for s in wb.getroot().findall('.//w:sheet', WNS):
            rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') or ''
            sheets.append((s.get('name', ''), rid))
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

def read_rows(path):
    """Return a sheet as a list of row-lists. Supports .csv (plain text, no
    base64 needed — preferred for cloud transfer) and .xlsx (binary)."""
    if path.lower().endswith('.csv'):
        import csv
        with open(path, newline='', encoding='utf-8-sig') as f:
            return list(csv.reader(f))
    sheets = read_xlsx(path)
    return list(sheets.values())[0] if sheets else []

def rows_to_dicts(rows):
    if not rows:
        return []
    headers = [str(h).strip().lower().replace(' ', '_') for h in rows[0]]
    return [dict(zip(headers, row)) for row in rows[1:] if any(c for c in row)]

def safe_int(v):
    try:
        return int(float(str(v).strip()))
    except:
        return None

def safe_pct(v):
    if v is None: return None
    s = str(v).strip()
    has_pct_sign = '%' in s
    # Handle '1% (2 systems)' or '1.5%' or plain '45' or '0.45'
    m = re.match(r'([\d.]+)', s)
    if m:
        try:
            f = float(m.group(1))
            # If the string already contains '%', the number is already a percentage
            if has_pct_sign:
                return f
            return f if f > 1 else f * 100
        except:
            pass
    return None

def months_onboard_from(val):
    """Months between an onboard/first-upload date and now. Handles the several
    encodings that appear across input files: 'YYYY-MM-DD' strings, Unix epoch
    milliseconds (e.g. 1724198400000), Unix epoch seconds, and Excel serial day
    numbers (days since 1899-12-30)."""
    if val is None: return None
    s = str(val).strip()
    if not s: return None
    now = datetime.datetime.now()
    if re.fullmatch(r'\d+(\.\d+)?([eE][+-]?\d+)?', s):   # incl. sci-notation e.g. 1.7242E+12
        f = float(s)
        if f > 1e11:       # epoch milliseconds
            dt = datetime.datetime(1970, 1, 1) + datetime.timedelta(milliseconds=f)
        elif f > 1e8:      # epoch seconds
            dt = datetime.datetime(1970, 1, 1) + datetime.timedelta(seconds=f)
        else:              # Excel serial day number
            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=f)
    else:
        try:
            dt = datetime.datetime.strptime(s[:10], '%Y-%m-%d')
        except:
            return None
    months = (now.year - dt.year) * 12 + (now.month - dt.month)
    return months if months >= 0 else None

# ── Filename parsing helpers ──────────────────────────────────────────────────
def extract_fleet_name(path):
    base = os.path.basename(path)
    parts = base.replace('Hat_CM_Summary_Extra_', '').replace('Hat_CM_Summary_', '').split('_')
    date_idx = next((i for i, p in enumerate(parts) if '-' in p and p[0].isdigit()), len(parts))
    return '_'.join(parts[:date_idx]).replace('.xlsx', '').upper()

def extract_date_ym(path):
    base = os.path.basename(path)
    m = re.search(r'(\d{4}-\d{2})', base)
    return m.group(1) if m else ''

# ── Colour helpers ────────────────────────────────────────────────────────────
def overdue_color(n):
    if n is None: return '#9CA3AF'
    if n == 0:    return '#439B38'
    if n >= 80:   return '#D97706'
    if n >= 120:  return '#D2393C'
    return '#9CA3AF'

def pct_color(pct, thresholds):
    # thresholds: list of (value, color) sorted low to high
    if pct is None: return '#9CA3AF'
    for val, col in sorted(thresholds, reverse=True):
        if pct >= val: return col
    return '#16181A'

# ── HTML helpers ──────────────────────────────────────────────────────────────
def e(text):
    return html_mod.escape(str(text)) if text is not None else ''

def cond_bar(critical, alert, normal):
    """Inline mini condition bar matching reference exactly."""
    c = max(critical or 0, 0)
    a = max(alert or 0, 0)
    n = max(normal or 0, 0)
    u = max(100 - c - a - n, 0)
    parts = []
    if c > 0: parts.append(f'<div style="flex:{c};background:#D2393C;"></div>')
    if a > 0: parts.append(f'<div style="flex:{a};background:#DD7814;"></div>')
    if n > 0: parts.append(f'<div style="flex:{n};background:#439B38;"></div>')
    if u > 0: parts.append(f'<div style="flex:{u};background:#E8E8E6;"></div>')
    return f'<div class="mb">{"".join(parts)}</div>'

def level_badge(level):
    """Priority level badge matching reference exactly."""
    styles = {
        'URGENT':   ('FEE2E2', 'FECACA', 'D2393C'),
        'HIGH':     ('FEF3C7', 'FDE68A', 'DD7814'),
        'MED-HIGH': ('FEF9C3', 'FEF08A', 'B45309'),
        'MEDIUM':   ('CFFAFE', 'A5F3FC', '0E7490'),
        'LOW':      ('DCFCE7', 'BBF7D0', '439B38'),
        'N/A':      ('F5F5F3', 'E8E8E6', '9CA3AF'),
    }
    bg, border, txt = styles.get(level, styles['N/A'])
    return (f'<span style="background:#{bg};border:1px solid #{border};border-radius:999px;'
            f'padding:2px 8px;font-size:9px;font-weight:700;letter-spacing:0.08em;'
            f'color:#{txt};">{e(level)}</span>')

def rank_box(rank, level):
    """Rank number box matching reference exactly."""
    colors = {
        'URGENT':   ('D2393C', 'FFFFFF'),
        'HIGH':     ('DD7814', 'FFFFFF'),
        'MED-HIGH': ('EBB71A', '16181A'),
        'MEDIUM':   ('8DC8CD', '16181A'),
        'LOW':      ('439B38', 'FFFFFF'),
        'N/A':      ('E8E8E6', '9CA3AF'),
    }
    bg, txt = colors.get(level, colors['N/A'])
    fs = '9px' if len(str(rank)) > 1 else '10px'
    return (f'<div style="width:20px;height:20px;border-radius:4px;background:#{bg};'
            f'display:flex;align-items:center;justify-content:center;'
            f'font-weight:800;font-size:{fs};color:#{txt};">{e(str(rank))}</div>')

def spotlight_card(label, vessel, detail, border_color, bg_color, label_color, border_side_color):
    return f'''<div style="background:{bg_color};border:1px solid {border_color};border-left:3px solid {border_side_color};border-radius:0 8px 8px 0;padding:16px 18px;">
        <div style="color:{label_color};font-size:9px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;margin-bottom:7px;">{e(label)}</div>
        <div style="color:#16181A;font-size:15px;font-weight:700;margin-bottom:3px;">{e(vessel)}</div>
        <div style="color:#667085;font-size:12px;">{e(detail)}</div>
      </div>'''

# ── Build vessel data ─────────────────────────────────────────────────────────
def build_vessel_data(sum_rows, ext_rows, prev_rows=None):
    ext_map = {str(r.get('ship_name','')).strip().upper(): r for r in ext_rows}
    prev_map = {str(r.get('ship_name','')).strip().upper(): r for r in (prev_rows or [])}
    vessels = []
    for r in sum_rows:
        name = str(r.get('ship_name', '')).strip().upper()
        if not name: continue
        ex   = ext_map.get(name, {})
        prev = prev_map.get(name, {})
        uploads = safe_int(r.get('all_uploads') or r.get('uploads_since_hat_start_date') or r.get('uploads_since_start'))
        inactive = (uploads is None or uploads == 0)
        months   = safe_int(ex.get('months_onboard'))
        if months is None:
            months = months_onboard_from(r.get('min_upload_date') or r.get('onboard_date'))
        overdue  = safe_int(r.get('overdue'))
        overdue_prev = safe_int(prev.get('overdue')) if prev else None
        c_pct = safe_pct(r.get('critical_pct') or r.get('critical'))
        a_pct = safe_pct(r.get('alert_pct') or r.get('alert'))
        n_pct = safe_pct(r.get('normal_pct') or r.get('normal'))
        low_alarm = str(r.get('low_usage_alarm', '')).strip().lower() in ('1', 'yes', 'true')
        vessels.append({
            'name': name, 'months': months, 'overdue': overdue,
            'overdue_prev': overdue_prev, 'critical_pct': c_pct,
            'alert_pct': a_pct, 'normal_pct': n_pct,
            'low_alarm': low_alarm, 'inactive': inactive,
        })
    return vessels

# ── Page builders ─────────────────────────────────────────────────────────────
def build_page1(fleet_name, vessel_count, date_label):
    return f'''<!-- ══ PAGE 1: COVER ═══════════════════════════════════════════════════════ -->
<section data-screen-label="01 Cover" style="background:#16181A;color:#FFFFFF;">
  <img src="a65ae821-fec6-4d3a-9ffe-6ed1dfcab9e7" alt="" style="position:absolute;width:500px;opacity:0.028;top:42%;left:50%;transform:translate(-50%,-50%);pointer-events:none;">

  <div style="display:flex;justify-content:space-between;align-items:center;padding:44px 56px 0;position:relative;z-index:1;">
    <div style="display:flex;align-items:center;gap:9px;">
      <img src="a65ae821-fec6-4d3a-9ffe-6ed1dfcab9e7" alt="" style="height:26px;display:block;">
      <span style="font-size:21px;font-weight:800;letter-spacing:-0.01em;color:#FFFFFF;"><span style="color:#00AABC;">HAT</span>ANALYTICS</span>
    </div>
    <span style="color:#3a3f44;font-size:11px;font-weight:700;letter-spacing:0.18em;">{e(fleet_name)}</span>
  </div>

  <div style="flex:1;display:flex;flex-direction:column;padding:0 56px;margin-top:88px;position:relative;z-index:1;">
    <div style="color:#00AABC;font-size:11px;font-weight:700;letter-spacing:0.2em;text-transform:uppercase;margin-bottom:22px;">Fleet Condition Monitoring Report</div>
    <div style="color:#FFFFFF;font-size:80px;font-weight:800;line-height:0.93;letter-spacing:-0.03em;margin-bottom:4px;">{e(fleet_name)}</div>
    <div style="color:rgba(255,255,255,0.18);font-size:80px;font-weight:200;line-height:0.93;letter-spacing:-0.02em;font-style:italic;margin-bottom:56px;">Fleet</div>
    <div style="display:flex;gap:8px;align-items:center;margin-bottom:12px;">
      <span style="color:#667085;font-size:15px;font-weight:500;">{e(vessel_count)} Vessels</span>
      <span style="color:#2e3438;font-size:15px;">·</span>
      <span style="color:#667085;font-size:15px;font-weight:500;">{e(date_label)}</span>
    </div>
    <div style="color:#2e3438;font-size:11px;font-weight:500;letter-spacing:0.04em;">System: AdViSe ATEX Handheld Vibration Monitoring Device</div>
    <div style="flex:1;"></div>
    <div style="height:1px;background:#242424;margin-bottom:22px;"></div>
    <div style="display:flex;justify-content:space-between;align-items:flex-start;">
      <div>
        <div style="color:#667085;font-size:10px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;margin-bottom:7px;">HAT services certified by</div>
        <div style="color:#D9D9D9;font-size:12px;line-height:1.85;">{SIG_CERTS}</div>
      </div>
      <div style="text-align:right;">
        <div style="color:#667085;font-size:10px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;margin-bottom:7px;">Report supervised by</div>
        <div style="color:#D9D9D9;font-size:12px;line-height:1.85;">{SIG_SUPERVISOR}</div>
      </div>
    </div>
  </div>

  <div style="background:#00AABC;padding:13px 56px;display:flex;justify-content:space-between;align-items:center;flex-shrink:0;">
    <span style="color:#FFFFFF;font-size:12px;font-weight:700;">HAT Analytics Solutions Ltd</span>
    <span style="color:rgba(255,255,255,0.72);font-size:12px;">{SIG_EMAIL}</span>
  </div>
</section>'''

def page_header(fleet_name, page_num, total_pages, date_label):
    return f'''<div class="ph">
    <div class="ph-logo">
      <img src="a65ae821-fec6-4d3a-9ffe-6ed1dfcab9e7" alt="">
      <span class="ph-logo-text"><span>HAT</span>ANALYTICS</span>
    </div>
    <div class="ph-right">{e(fleet_name)} FLEET · Fleet Condition Monitoring Report<br>Page {page_num} of {total_pages} · {e(date_label)}</div>
  </div>'''

def page_footer(fleet_name, date_label):
    return f'''<div class="pf">
    <span>HAT Analytics Solutions Ltd · Fleet Condition Monitoring Report</span>
    <span>{e(fleet_name)} FLEET · {e(date_label)}</span>
  </div>'''

def build_page2(fleet_name, date_label, exec_paras, spotlights):
    cards_html = ''.join([
        spotlight_card(s['label'], s['vessel'], s['detail'],
                      s['border_color'], s['bg_color'], s['label_color'], s['border_side'])
        for s in spotlights
    ])
    return f'''<!-- ══ PAGE 2: EXECUTIVE SUMMARY ══════════════════════════════════════════ -->
<section data-screen-label="02 Executive Summary">
  {page_header(fleet_name, 2, 4, date_label)}
  <div class="pc">
    <div class="sec">Executive Summary</div>
    {exec_paras}
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
      {cards_html}
    </div>
  </div>
  {page_footer(fleet_name, date_label)}
</section>'''

def build_page3(fleet_name, date_label, vessels):
    rows_html = ''
    for v in vessels:
        name = v['name']
        inactive = v['inactive']
        od = v['overdue']
        c_pct = v['critical_pct']
        a_pct = v['alert_pct']
        n_pct = v['normal_pct']

        if inactive:
            rows_html += f'''<tr>
          <td class="vn" style="color:#9CA3AF;">{e(name)}</td>
          <td class="sl">—</td><td class="sl">—</td>
          <td class="sl">—</td><td class="sl">—</td><td class="sl">—</td>
          <td><span style="font-size:8px;color:#C4C9D0;font-weight:700;letter-spacing:0.07em;">NEWLY ACTIVATED</span></td>
        </tr>'''
            continue

        od_color = overdue_color(od)
        c_color  = pct_color(c_pct, [(5,'#D2393C'),(2,'#D2393C')])
        a_color  = pct_color(a_pct, [(7,'#D2393C'),(5,'#DD7814')])
        n_color  = '#439B38'

        # Name coloring (red if vessel is critical)
        name_style = 'color:#D2393C;' if c_pct and c_pct >= 7 else \
                     'color:#439B38;' if od == 0 else ''

        od_bold = 'font-weight:700;' if od and (od >= 80 or od == 0) else ''
        c_bold  = 'font-weight:700;' if c_pct and c_pct >= 5 else ''
        a_bold  = 'font-weight:700;' if a_pct and a_pct >= 5 else ''

        od_str = str(od) if od is not None else '—'
        c_str  = f'{int(c_pct)}%' if c_pct is not None else '—'
        a_str  = f'{int(a_pct)}%' if a_pct is not None else '—'
        n_str  = f'{int(n_pct)}%' if n_pct is not None else '—'
        mo_str = str(v['months']) if v['months'] is not None else '—'

        bar = cond_bar(c_pct, a_pct, n_pct)
        rows_html += f'''<tr>
          <td class="vn" style="{name_style}">{e(name)}</td>
          <td class="sl">{e(mo_str)}</td>
          <td style="color:{od_color};{od_bold}">{e(od_str)}</td>
          <td style="color:{c_color};{c_bold}">{e(c_str)}</td>
          <td style="color:{a_color};{a_bold}">{e(a_str)}</td>
          <td class="nm">{e(n_str)}</td>
          <td>{bar}</td>
        </tr>'''

    return f'''<!-- ══ PAGE 3: FLEET OVERVIEW ════════════════════════════════════════════ -->
<section data-screen-label="03 Fleet Overview">
  {page_header(fleet_name, 3, 4, date_label)}
  <div class="pc">
    <div class="sec">Fleet Overview</div>
    <table class="ft">
      <thead>
        <tr>
          <th>Vessel</th><th>Months</th><th>Overdue</th>
          <th>Critical %</th><th>Alert %</th><th>Normal %</th><th>Condition</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
    <div style="display:flex;gap:14px;margin-top:14px;align-items:center;">
      <span style="font-size:9px;color:#C4C9D0;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;">Condition bar:</span>
      <div style="display:flex;gap:12px;align-items:center;">
        <span style="display:flex;align-items:center;gap:4px;"><span style="width:7px;height:7px;background:#D2393C;border-radius:1px;display:inline-block;"></span><span style="font-size:10px;color:#9CA3AF;">Critical</span></span>
        <span style="display:flex;align-items:center;gap:4px;"><span style="width:7px;height:7px;background:#DD7814;border-radius:1px;display:inline-block;"></span><span style="font-size:10px;color:#9CA3AF;">Alert</span></span>
        <span style="display:flex;align-items:center;gap:4px;"><span style="width:7px;height:7px;background:#439B38;border-radius:1px;display:inline-block;"></span><span style="font-size:10px;color:#9CA3AF;">Normal</span></span>
        <span style="display:flex;align-items:center;gap:4px;"><span style="width:7px;height:7px;background:#E8E8E6;border-radius:1px;display:inline-block;"></span><span style="font-size:10px;color:#9CA3AF;">Unmonitored</span></span>
      </div>
    </div>
  </div>
  {page_footer(fleet_name, date_label)}
</section>'''

def build_page4(fleet_name, date_label, priorities, recs):
    pri_rows = ''
    for p in priorities:
        rank = p['rank']
        vessel = p['vessel']
        level = p['level']
        issue = p['key_issue']
        dimmed = p.get('dimmed', False)
        v_color = '#C4C9D0' if dimmed else ('#D2393C' if level == 'URGENT' else '#439B38' if level == 'LOW' else '#16181A')
        pri_rows += f'''<tr>
          <td style="padding:7px 8px 7px 12px;border-bottom:1px solid #F0EFED;">{rank_box(rank, level)}</td>
          <td style="padding:7px 10px;border-bottom:1px solid #F0EFED;font-size:11px;font-weight:700;letter-spacing:0.04em;color:{v_color};">{e(vessel)}</td>
          <td style="padding:7px 10px;border-bottom:1px solid #F0EFED;">{level_badge(level)}</td>
          <td style="padding:7px 10px;border-bottom:1px solid #F0EFED;font-size:11px;color:{'#C4C9D0' if dimmed else '#9CA3AF'};">{e(issue)}</td>
        </tr>'''

    rec_colors = {
        'URGENT': ('#FEF2F2', '#FECACA', '#D2393C'),
        'HIGH':   ('#FFF7ED', '#FED7AA', '#DD7814'),
        'MEDIUM': ('#FFFBEB', '#FDE68A', '#EBB71A'),
        'LOW':    ('#F0FDF4', '#BBF7D0', '#439B38'),
        'INFO':   ('#ECFEFF', '#A5F3FC', '#00AABC'),
    }
    recs_html = ''
    for r in recs:
        bg, border, left = rec_colors.get(r.get('level', 'INFO'), rec_colors['INFO'])
        recs_html += f'''<div style="padding:8px 12px;background:{bg};border:1px solid {border};border-left:3px solid {left};border-radius:0 4px 4px 0;margin-bottom:6px;font-size:11.5px;color:#22282B;line-height:1.55;">
      {r['text']}
    </div>'''

    return f'''<!-- ══ PAGE 4: PRIORITY ASSESSMENT & RECOMMENDATIONS ════════════════════ -->
<section data-screen-label="04 Priority &amp; Recommendations">
  {page_header(fleet_name, 4, 4, date_label)}
  <div class="pc">
    <div class="sec">Priority Assessment</div>
    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead>
        <tr>
          <th style="background:#F5F5F3;padding:8px 8px 8px 12px;font-size:9px;font-weight:700;letter-spacing:0.11em;text-transform:uppercase;color:#667085;text-align:left;border-radius:6px 0 0 6px;width:34px;">#</th>
          <th style="background:#F5F5F3;padding:8px 10px;font-size:9px;font-weight:700;letter-spacing:0.11em;text-transform:uppercase;color:#667085;text-align:left;width:168px;">Vessel</th>
          <th style="background:#F5F5F3;padding:8px 10px;font-size:9px;font-weight:700;letter-spacing:0.11em;text-transform:uppercase;color:#667085;text-align:left;width:116px;">Level</th>
          <th style="background:#F5F5F3;padding:8px 10px;font-size:9px;font-weight:700;letter-spacing:0.11em;text-transform:uppercase;color:#667085;text-align:left;border-radius:0 6px 6px 0;">Key Issue</th>
        </tr>
      </thead>
      <tbody>{pri_rows}</tbody>
    </table>
    <div style="height:1px;background:#F0EFED;margin:12px 0 14px;"></div>
    <div class="sec" style="margin-bottom:10px;">Recommendations</div>
    {recs_html}
    <div style="flex:1;"></div>
    <div style="height:1px;background:#E8E8E6;margin-bottom:18px;margin-top:14px;"></div>
    <div style="display:flex;justify-content:space-between;align-items:flex-start;">
      <div>
        <div style="color:#667085;font-size:10px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;margin-bottom:7px;">HAT services certified by</div>
        <div style="color:#22282B;font-size:12px;line-height:1.85;">{SIG_CERTS}</div>
      </div>
      <div style="text-align:right;">
        <div style="color:#667085;font-size:10px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;margin-bottom:7px;">Report supervised by</div>
        <div style="color:#22282B;font-size:12px;line-height:1.85;">{SIG_SUPERVISOR}</div>
      </div>
    </div>
  </div>
  {page_footer(fleet_name, date_label)}
</section>'''

# ── Deck-stage bundle handling ────────────────────────────────────────────────
# The reference standalone HTML is a "bundle": the entire report HTML is stored
# as a JSON-encoded string inside <script type="__bundler/template">, and a
# loader script JSON.parses it at runtime. Editing the raw file with regex would
# splice literal newlines/quotes into that JSON string and break JSON.parse
# ("Bad control character in string literal in JSON"). So we must decode the
# packed template, edit the real HTML, then re-encode it.
BUNDLE_RE = re.compile(r'(<script type="__bundler/template">)(.*?)(</script>)', re.DOTALL)

def split_bundle(outer_html):
    """Return (prefix, inner_html, suffix). If not a bundle, prefix/suffix are
    None and inner_html is the original text (treated as raw editable HTML)."""
    m = BUNDLE_RE.search(outer_html)
    if not m:
        return None, outer_html, None
    inner = json.loads(m.group(2).strip())
    return outer_html[:m.start(2)], inner, outer_html[m.end(2):]

def repack_bundle(prefix, inner_html, suffix):
    """Re-encode edited inner HTML back into the bundle's template <script>."""
    if prefix is None:
        return inner_html
    packed = json.dumps(inner_html)          # ensure_ascii escapes all non-ASCII
    packed = packed.replace('</', '<\\/')    # keep </script> from closing the tag
    return prefix + '\n  ' + packed + '\n  ' + suffix

# ── Assemble full document from reference skeleton ────────────────────────────
def build_report(template_path, vessels, fleet_name, date_label, vessel_count,
                 exec_paras, spotlights, priorities, recs):
    with open(template_path, 'r', encoding='utf-8') as f:
        outer = f.read()
    prefix, inner, suffix = split_bundle(outer)

    # Build the 4 pages
    p1 = build_page1(fleet_name, vessel_count, date_label)
    p2 = build_page2(fleet_name, date_label, exec_paras, spotlights)
    p3 = build_page3(fleet_name, date_label, vessels)
    p4 = build_page4(fleet_name, date_label, priorities, recs)

    new_body = f'''<deck-stage width="794" height="1123">

{p1}

{p2}

{p3}

{p4}

</deck-stage>'''

    # Replace the deck-stage content in the decoded inner HTML
    inner = re.sub(
        r'<deck-stage[^>]*>.*?</deck-stage>',
        lambda _: new_body,
        inner,
        flags=re.DOTALL
    )
    # Update title
    new_title = f'<title>{e(fleet_name)} Fleet — Fleet Condition Monitoring Report — {e(date_label)}</title>'
    inner = re.sub(r'<title>[^<]*</title>', lambda _: new_title, inner)

    return repack_bundle(prefix, inner, suffix)

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--summary',      required=True)
    ap.add_argument('--extra',        required=False, default=None)
    ap.add_argument('--summary-prev', required=False)
    ap.add_argument('--template',     required=True, help='Path to reference standalone HTML')
    ap.add_argument('--out',          required=True)
    ap.add_argument('--fleet',        required=False)
    args = ap.parse_args()

    fleet_name = (args.fleet or extract_fleet_name(args.summary)).upper()

    # Date is always the current month at runtime
    now        = datetime.datetime.now()
    date_ym    = now.strftime('%Y-%m')
    date_label = now.strftime('%B %Y')

    # Read Excel data
    sum_rows  = rows_to_dicts(read_rows(args.summary))
    ext_rows  = rows_to_dicts(read_rows(args.extra)) if args.extra else []
    prev_rows = rows_to_dicts(read_rows(args.summary_prev)) if args.summary_prev else []

    vessels = build_vessel_data(sum_rows, ext_rows, prev_rows)
    active  = [v for v in vessels if not v['inactive']]
    vessel_count = str(len(active))

    # Claude fills these in during the skill run — placeholders show structure
    # The skill instructs Claude to: 1) read vessel data printed below,
    # 2) write exec_paras, spotlights, priorities, recs into the HTML

    # Print extracted vessel data for Claude to analyse
    print(f"\n=== EXTRACTED VESSEL DATA ({fleet_name} / {date_label}) ===")
    for v in vessels:
        status = "INACTIVE" if v['inactive'] else "active"
        delta = ''
        if v['overdue_prev'] is not None:
            diff = (v['overdue'] or 0) - v['overdue_prev']
            arrow = '+' if diff > 0 else ''
            delta = f", overdue_prev={v['overdue_prev']} ({arrow}{diff})"
        print(f"  {v['name']}: months={v['months']}, overdue={v['overdue']}{delta}, "
              f"crit={v['critical_pct']}%, alert={v['alert_pct']}%, "
              f"normal={v['normal_pct']}%, low_alarm={v['low_alarm']}, {status}")
    print(f"\nActive vessels: {vessel_count}")

    # Default placeholders — Claude replaces these in the skill workflow
    exec_paras = f'<p style="font-size:13.5px;line-height:1.78;color:#22282B;margin-bottom:20px;">Fleet condition monitoring report for {e(fleet_name)}, {e(date_label)}. [Claude: replace with executive summary paragraphs]</p>'

    spotlights = [
        {'label':'MOST CRITICAL',      'vessel':'[VESSEL]', 'detail':'[detail]',
         'bg_color':'#FEF2F2','border_color':'#FECACA','label_color':'#D2393C','border_side':'#D2393C'},
        {'label':'STALLED COMPLIANCE', 'vessel':'[VESSEL]', 'detail':'[detail]',
         'bg_color':'#FFFBEB','border_color':'#FDE68A','label_color':'#D97706','border_side':'#EBB71A'},
        {'label':'MOST IMPROVED',      'vessel':'[VESSEL]', 'detail':'[detail]',
         'bg_color':'#F0FDF4','border_color':'#BBF7D0','label_color':'#439B38','border_side':'#439B38'},
        {'label':'BEST COMPLIANCE',    'vessel':'[VESSEL]', 'detail':'[detail]',
         'bg_color':'#ECFEFF','border_color':'#A5F3FC','label_color':'#00AABC','border_side':'#00AABC'},
    ]

    priorities = [
        {'rank': i+1, 'vessel': v['name'],
         'level': 'N/A' if v['inactive'] else 'MEDIUM',
         'key_issue': 'System newly activated — insufficient data' if v['inactive'] else '[Claude: assess and rank]',
         'dimmed': v['inactive']}
        for i, v in enumerate(vessels)
    ]

    recs = [
        {'level': 'INFO', 'text': '[Claude: write recommendations based on vessel data above]'}
    ]

    html = build_report(args.template, vessels, fleet_name, date_label,
                        vessel_count, exec_paras, spotlights, priorities, recs)

    with open(args.out, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\nHTML written -> {args.out}")
    print("Next: edit the HTML to fill in exec summary, spotlights, priorities, and recommendations")
    print("Then: node scripts/print_pdf.js output.html output.pdf")

if __name__ == '__main__':
    main()
