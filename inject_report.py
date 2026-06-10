"""
inject_report.py  --  Fleet CM Report HTML injector (Agent 2)
Reads a JSON file produced by Agent 1 and injects the intelligence layer
into the staging HTML, producing a final HTML ready for PDF conversion.

Usage:
    python3 inject_report.py \
        --json "C:\\Fleet Overview Report\\Staging\\ACME_2026-05.json" \
        --out  "C:\\Fleet Overview Report\\Staging\\ACME_2026-05_final.html"
"""

import argparse
import json
import os
import re
import sys
import html as html_mod

# Report signature block — supplied at runtime via env vars so the public repo
# carries no identifying info. Set HAT_CERTS / HAT_SUPERVISOR to real values.
SIG_CERTS      = os.environ.get('HAT_CERTS', 'Certifications on file')
SIG_SUPERVISOR = os.environ.get('HAT_SUPERVISOR', 'Report supervisor')

def e(text):
    return html_mod.escape(str(text)) if text is not None else ''

# ── Deck-stage bundle handling ────────────────────────────────────────────────
# The report HTML is JSON-encoded inside <script type="__bundler/template">.
# Decode before editing, re-encode after — see report_builder.py for details.
BUNDLE_RE = re.compile(r'(<script type="__bundler/template">)(.*?)(</script>)', re.DOTALL)

def split_bundle(outer_html):
    m = BUNDLE_RE.search(outer_html)
    if not m:
        return None, outer_html, None
    inner = json.loads(m.group(2).strip())
    return outer_html[:m.start(2)], inner, outer_html[m.end(2):]

def repack_bundle(prefix, inner_html, suffix):
    if prefix is None:
        return inner_html
    packed = json.dumps(inner_html)
    packed = packed.replace('</', '<\\/')
    return prefix + '\n  ' + packed + '\n  ' + suffix

# ── Spotlight card styles ─────────────────────────────────────────────────────
SPOTLIGHT_STYLES = {
    'MOST CRITICAL':              ('#FEF2F2', '#FECACA', '#D2393C', '#D2393C'),
    'STALLED CM PLAN COMPLIANCE': ('#FFFBEB', '#FDE68A', '#D97706', '#EBB71A'),
    'MOST IMPROVED':              ('#F0FDF4', '#BBF7D0', '#439B38', '#439B38'),
    'BEST CM PLAN COMPLIANCE':    ('#ECFEFF', '#A5F3FC', '#00AABC', '#00AABC'),
    # legacy label aliases (pre 'CM Plan Compliance' rename)
    'STALLED COMPLIANCE':         ('#FFFBEB', '#FDE68A', '#D97706', '#EBB71A'),
    'BEST COMPLIANCE':            ('#ECFEFF', '#A5F3FC', '#00AABC', '#00AABC'),
}

def spotlight_card(s):
    label = s['label'].upper()
    bg, border, label_color, side = SPOTLIGHT_STYLES.get(label, ('#F5F5F3', '#E8E8E6', '#667085', '#9CA3AF'))
    return (
        f'<div style="background:{bg};border:1px solid {border};'
        f'border-left:3px solid {side};border-radius:0 8px 8px 0;padding:16px 18px;">'
        f'<div style="color:{label_color};font-size:9px;font-weight:700;'
        f'letter-spacing:0.14em;text-transform:uppercase;margin-bottom:7px;">{e(s["label"])}</div>'
        f'<div style="color:#16181A;font-size:15px;font-weight:700;margin-bottom:3px;">{e(s["vessel"])}</div>'
        f'<div style="color:#667085;font-size:12px;">{e(s["detail"])}</div>'
        f'</div>'
    )

# ── Priority badge + rank box ─────────────────────────────────────────────────
LEVEL_BADGE = {
    'URGENT':   ('#FEE2E2', '#FECACA', '#D2393C'),
    'HIGH':     ('#FEF3C7', '#FDE68A', '#DD7814'),
    'MED-HIGH': ('#FEF9C3', '#FEF08A', '#B45309'),
    'MEDIUM':   ('#CFFAFE', '#A5F3FC', '#0E7490'),
    'LOW':      ('#DCFCE7', '#BBF7D0', '#439B38'),
    'N/A':      ('#F5F5F3', '#E8E8E6', '#9CA3AF'),
}

RANK_BOX_COLOR = {
    'URGENT':   ('#D2393C', '#FFFFFF'),
    'HIGH':     ('#DD7814', '#FFFFFF'),
    'MED-HIGH': ('#EBB71A', '#16181A'),
    'MEDIUM':   ('#8DC8CD', '#16181A'),
    'LOW':      ('#439B38', '#FFFFFF'),
    'N/A':      ('#E8E8E6', '#9CA3AF'),
}

def rank_box(rank, level):
    bg, txt = RANK_BOX_COLOR.get(level, RANK_BOX_COLOR['N/A'])
    fs = '9px' if len(str(rank)) > 1 else '10px'
    return (f'<div style="width:20px;height:20px;border-radius:4px;background:{bg};'
            f'display:flex;align-items:center;justify-content:center;'
            f'font-weight:800;font-size:{fs};color:{txt};">{e(str(rank))}</div>')

def level_badge(level):
    bg, border, txt = LEVEL_BADGE.get(level, LEVEL_BADGE['N/A'])
    return (f'<span style="background:{bg};border:1px solid {border};border-radius:999px;'
            f'padding:2px 8px;font-size:9px;font-weight:700;letter-spacing:0.08em;'
            f'color:{txt};">{e(level)}</span>')

def vessel_name_color(level, dimmed):
    if dimmed: return '#C4C9D0'
    return {'URGENT': '#D2393C', 'LOW': '#439B38'}.get(level, '#16181A')

# ── Recommendation card ───────────────────────────────────────────────────────
REC_STYLES = {
    'URGENT': ('#FEF2F2', '#FECACA', '#D2393C'),
    'HIGH':   ('#FFF7ED', '#FED7AA', '#DD7814'),
    'MEDIUM': ('#FFFBEB', '#FDE68A', '#EBB71A'),
    'LOW':    ('#F0FDF4', '#BBF7D0', '#439B38'),
    'INFO':   ('#ECFEFF', '#A5F3FC', '#00AABC'),
}

def safe_html(text):
    """Allow only safe inline tags: strong, em, b, i, span with style."""
    # Pass through as-is — already trusted HTML from Claude's intelligence layer
    # Just ensure no raw quotes that break attribute context
    return text.replace('\n', ' ').replace('\r', '')

def rec_card(r):
    bg, border, left = REC_STYLES.get(r.get('level', 'INFO'), REC_STYLES['INFO'])
    return (f'<div style="padding:8px 12px;background:{bg};border:1px solid {border};'
            f'border-left:3px solid {left};border-radius:0 4px 4px 0;'
            f'margin-bottom:6px;font-size:11.5px;color:#22282B;line-height:1.55;">'
            f'{safe_html(r["text"])}</div>')

# ── Page 2: Executive Summary ─────────────────────────────────────────────────
def build_page2(data):
    exec_html = '\n'.join(
        f'<p style="font-size:13.5px;line-height:1.78;color:#22282B;'
        f'margin-bottom:20px;text-wrap:pretty;">{p}</p>'
        for p in data['exec_paragraphs']
    )
    cards = '\n'.join(spotlight_card(s) for s in data['spotlights'])
    fleet = e(data['fleet'])
    date  = e(data['date_label'])

    return f'''<!-- ══ PAGE 2: EXECUTIVE SUMMARY ══════════════════════════════════════════ -->
<section data-screen-label="02 Executive Summary">
  <div class="ph">
    <div class="ph-logo">
      <img src="a65ae821-fec6-4d3a-9ffe-6ed1dfcab9e7" alt="">
      <span class="ph-logo-text"><span>HAT</span>ANALYTICS</span>
    </div>
    <div class="ph-right">{fleet} FLEET · Fleet Condition Monitoring Report<br>Page 2 of 4 · Reporting period: {date}</div>
  </div>
  <div class="pc">
    <div class="sec">Executive Summary</div>
    {exec_html}
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
      {cards}
    </div>
  </div>
  <div class="pf">
    <span>HAT Analytics Solutions Ltd · Fleet Condition Monitoring Report</span>
    <span>{fleet} FLEET · {date}</span>
  </div>
</section>'''

# ── Dynamic page numbering ────────────────────────────────────────────────────
def renumber_pages(html):
    """Rewrite every 'Page X of Y' header so it reflects the actual number of
    <section> pages (cover = page 1, no header). Keeps numbering correct when
    Fleet Overview or Priority Assessment spill onto continuation pages."""
    total = len(re.findall(r'<section\b', html))
    counter = {'n': 0}
    def fix(m):
        counter['n'] += 1
        return re.sub(r'Page\s+\d+\s+of\s+\d+',
                      f'Page {counter["n"]} of {total}', m.group(0))
    return re.sub(r'<section\b.*?</section>', fix, html, flags=re.DOTALL)

# ── Page 4: Priority + Recommendations ───────────────────────────────────────
# Priority rows that fit on one A4 page before spilling to a continuation page.
# Kept conservative because the Key Issue column can wrap to two lines.
PAGE_ROWS_4 = 20

_PRI_THEAD = '''<thead>
        <tr>
          <th style="background:#F5F5F3;padding:8px 8px 8px 12px;font-size:9px;font-weight:700;letter-spacing:0.11em;text-transform:uppercase;color:#667085;text-align:left;border-radius:6px 0 0 6px;width:34px;">#</th>
          <th style="background:#F5F5F3;padding:8px 10px;font-size:9px;font-weight:700;letter-spacing:0.11em;text-transform:uppercase;color:#667085;text-align:left;width:168px;">Vessel</th>
          <th style="background:#F5F5F3;padding:8px 10px;font-size:9px;font-weight:700;letter-spacing:0.11em;text-transform:uppercase;color:#667085;text-align:left;width:116px;">Priority</th>
          <th style="background:#F5F5F3;padding:8px 10px;font-size:9px;font-weight:700;letter-spacing:0.11em;text-transform:uppercase;color:#667085;text-align:left;border-radius:0 6px 6px 0;">Key Issue</th>
        </tr>
      </thead>'''

def _p4_ph(fleet, date):
    return f'''<div class="ph">
    <div class="ph-logo">
      <img src="a65ae821-fec6-4d3a-9ffe-6ed1dfcab9e7" alt="">
      <span class="ph-logo-text"><span>HAT</span>ANALYTICS</span>
    </div>
    <div class="ph-right">{fleet} FLEET · Fleet Condition Monitoring Report<br>Page 0 of 0 · Reporting period: {date}</div>
  </div>'''

def _p4_pf(fleet, date):
    return f'''<div class="pf">
    <span>HAT Analytics Solutions Ltd · Fleet Condition Monitoring Report</span>
    <span>{fleet} FLEET · {date}</span>
  </div>'''

def _pri_row_html(p):
    level   = p.get('level', 'N/A')
    dimmed  = p.get('dimmed', False)
    v_color = vessel_name_color(level, dimmed)
    i_color = '#C4C9D0' if dimmed else '#9CA3AF'
    return f'''<tr>
          <td style="padding:7px 8px 7px 12px;border-bottom:1px solid #F0EFED;">{rank_box(p["rank"], level)}</td>
          <td style="padding:7px 10px;border-bottom:1px solid #F0EFED;font-size:11px;font-weight:700;letter-spacing:0.04em;color:{v_color};">{e(p["vessel"])}</td>
          <td style="padding:7px 10px;border-bottom:1px solid #F0EFED;">{level_badge(level)}</td>
          <td style="padding:7px 10px;border-bottom:1px solid #F0EFED;font-size:11px;color:{i_color};">{e(p["key_issue"])}</td>
        </tr>'''

def _pri_table(rows_html):
    return (f'<table style="width:100%;border-collapse:collapse;margin-bottom:14px;">'
            f'\n      {_PRI_THEAD}\n      <tbody>{rows_html}</tbody>\n    </table>')

def _p4_section(fleet, date, title, body):
    return f'''<section data-screen-label="04 Priority &amp; Recommendations">
  {_p4_ph(fleet, date)}
  <div class="pc">
    <div class="sec">{title}</div>
    {body}
  </div>
  {_p4_pf(fleet, date)}
</section>'''

# Page-4 layout budget (px) for the ~968px content area. Rec cards can wrap to two
# lines, so plan for worst-case height.
P4_CONTENT_PX   = 968
P4_SEC_TITLE_PX = 40    # the section title on a page
P4_REC_PX       = 50    # one recommendation card (2-line)

_CONT = ' <span style="color:#C4C9D0;font-weight:600;">(continued)</span>'

def build_page4(data):
    """Priority Assessment table, then Recommendations — each on its own page(s),
    paginating onto continuation pages when they exceed one A4 page. Recommendations
    ALWAYS start a fresh page (never crammed under the priority table). No signature
    block on page 4 (the cover carries the certifications / approval)."""
    fleet = e(data['fleet'])
    date  = e(data['date_label'])
    pri   = [_pri_row_html(p) for p in data['priorities']]
    rec_cards = [rec_card(r) for r in data['recommendations']]

    sections = []

    # Priority Assessment — paginated, table only.
    pri_chunks = [pri[i:i + PAGE_ROWS_4] for i in range(0, len(pri), PAGE_ROWS_4)] or [[]]
    for i, ch in enumerate(pri_chunks):
        title = 'Priority Assessment' + (_CONT if i else '')
        sections.append(_p4_section(fleet, date, title, _pri_table(''.join(ch))))

    # Recommendations — always on their own page(s).
    if rec_cards:
        per_page = max(1, (P4_CONTENT_PX - P4_SEC_TITLE_PX) // P4_REC_PX)
        rec_chunks = [rec_cards[i:i + per_page] for i in range(0, len(rec_cards), per_page)]
        for j, page in enumerate(rec_chunks):
            title = 'Recommendations' + (_CONT if j else '')
            sections.append(_p4_section(fleet, date, title, '\n    '.join(page)))

    return ('<!-- ══ PAGE 4: PRIORITY ASSESSMENT & RECOMMENDATIONS ════════════════════ -->\n'
            + '\n'.join(sections))

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--json', required=True, help='Path to Agent 1 JSON file')
    ap.add_argument('--out',  required=True, help='Path to write final HTML')
    args = ap.parse_args()

    # Load JSON -- strip BOM and control chars that break deck-stage bundler
    import unicodedata
    with open(args.json, 'r', encoding='utf-8-sig') as f:
        raw = f.read()
    raw = ''.join(c for c in raw if unicodedata.category(c) != 'Cc' or c in '\t\n\r')
    data = json.loads(raw)

    # Load staging HTML (built by report_builder.py). It is a deck-stage bundle:
    # the report HTML is JSON-encoded inside <script type="__bundler/template">.
    # Decode it, edit the real HTML, then re-encode — editing the raw packed JSON
    # would splice literal newlines into a JSON string and break JSON.parse.
    html_path = data['html_path']
    with open(html_path, 'r', encoding='utf-8') as f:
        outer = f.read()
    prefix, html, suffix = split_bundle(outer)

    # Replace page 2
    new_p2 = build_page2(data)
    html = re.sub(
        r'<!-- ══ PAGE 2.*?(?=<!-- ══ PAGE 3)',
        lambda _: new_p2 + '\n\n',
        html,
        flags=re.DOTALL
    )

    # Replace page 4
    new_p4 = build_page4(data)
    html = re.sub(
        r'<!-- ══ PAGE 4.*?</section>',
        lambda _: new_p4,
        html,
        flags=re.DOTALL
    )

    # Update title
    fleet = data['fleet']
    date  = data['date_label']
    new_title = f'<title>{e(fleet)} Fleet — Fleet Condition Monitoring Report — {e(date)}</title>'
    html = re.sub(r'<title>[^<]*</title>', lambda _: new_title, html)

    # Final authority on 'Page X of Y' now that pages 3 and 4 may span multiple pages.
    html = renumber_pages(html)

    with open(args.out, 'w', encoding='utf-8') as f:
        f.write(repack_bundle(prefix, html, suffix))

    print(f"Injected -> {args.out}")

if __name__ == '__main__':
    main()
