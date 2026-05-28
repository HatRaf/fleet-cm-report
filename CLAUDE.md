# Fleet CM Report Automation
# HAT Analytics Solutions Ltd
# Claude Code Routine

## Your job
Generate Fleet Condition Monitoring PDF reports from Excel input files.
One report per company per month. Run end to end without human intervention.

## Working directory
C:\Fleet Overview Reporting\

## Folder structure
Input files\
  MAY 2026\   <- one subfolder per month (uppercase month name + year)
    Hat_CM_Summary_EXAMPLE.xlsx
    Hat_CM_Summary_SAMPLE FLEET.xlsx
  APRIL 2026\ <- prior month (used for trend comparison)
    Hat_CM_Summary_EXAMPLE.xlsx
Reports\        <- Output PDFs, one subfolder per company
logs\           <- Run logs
Skills\         <- All scripts and template
Skills\Fleet_CM_Report_Template.html  <- reference design, NEVER modify this file

## Skills folder contents
Skills\SKILL.md                       <- read this first, every run
Skills\report_builder.py              <- reads Excel, builds HTML
Skills\inject_report.py               <- injects intelligence into HTML
Skills\print_pdf.js                   <- HTML to PDF via Chrome
Skills\read_excel.py                  <- fallback Excel reader
Skills\Fleet_CM_Report_Template.html  <- reference design, never modify

---

## Run instructions

### Step 1 — Read SKILL.md
Always read Skills\SKILL.md before doing anything else, every single run.

### Step 2 — Determine the report date
Date is always the current month at runtime.
  date_ym         = current year-month (e.g. "2026-05")
  date_label      = month and year (e.g. "May 2026")
  month_folder    = uppercase month name + year (e.g. "MAY 2026")
  prev_month_folder = previous calendar month, same format (e.g. "APRIL 2026")
Get these from the system date. Never hardcode them.

### Step 3 — Determine which companies to process

If the invocation prompt names a specific company (e.g. "Run the report for EXAMPLE"),
process ONLY that company. Do not touch any other files.

If no company is named, scan "C:\Fleet Overview Reporting\Input files\[month_folder]\"
for all files matching Hat_CM_Summary_[COMPANY].xlsx and process each one.

Extract company name from filename:
  Hat_CM_Summary_EXAMPLE.xlsx → company = "EXAMPLE"

Skip any company where the output PDF already exists:
  C:\Fleet Overview Reporting\Reports\[COMPANY]\Fleet_CM_Report_[COMPANY]_[YYYY-MM].pdf

### Step 4 — For each company file found:

#### 4a. Run the data extractor
  Prior month file (optional):
    prev_summary = "C:\Fleet Overview Reporting\Input files\[prev_month_folder]\Hat_CM_Summary_[COMPANY].xlsx"
    Check if this file exists. Include --summary-prev only if it does.
    If it does not exist → first run for this company → omit --summary-prev.

  python3 "C:\Fleet Overview Reporting\Skills\report_builder.py" \
    --summary      "C:\Fleet Overview Reporting\Input files\[month_folder]\Hat_CM_Summary_[COMPANY].xlsx" \
    --summary-prev "[prev_summary]" \                    ← omit entire line if file not found
    --template     "C:\Fleet Overview Reporting\Skills\Fleet_CM_Report_Template.html" \
    --fleet        "[COMPANY]" \
    --out          "C:\Fleet Overview Reporting\Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_staging.html"
  Create Reports\[COMPANY]\ if it does not exist.

#### 4b. Read the script output — single source of truth
  The script prints all extracted vessel data.
  Use ONLY these numbers for all analysis and writing.
  Never re-read the Excel file directly.
  Never invent or estimate numbers.
  If a number cannot be traced to the script output, do not use it.

#### 4c. Write the intelligence layer

  EXECUTIVE SUMMARY — strict rules:
  - Maximum 2 short paragraphs, max 8 sentences total
  - Paragraph 1 (2-3 sentences, fixed structure):
      Sentence 1: "This report covers [N] [FLEET] vessels enrolled in the HAT AdViSe
                   condition monitoring program."
      Sentence 2: "Overall fleet health is [good/moderate/concerning]: across [N] active vessels
                   the majority sit in Normal condition, and [N] vessels carry zero Critical or
                   Alert findings."
      Sentence 3 (optional): One sentence flagging the number of areas that warrant attention
                   this period, without naming vessels (e.g. "Two areas warrant attention this period."
                   or "Three vessels require urgent follow-up this period.").
      Keep this paragraph factual and brief — all detail goes in paragraph 2.
  - Paragraph 2 (3-4 sentences):
      Top concern with exact numbers, best performer with exact numbers,
      immediate priorities.
      If prior month data available: include deltas for key vessels
      (e.g. "overdue increased from 28 to 41").
      If no prior month data: note this is the first reporting cycle.
  - NO third paragraph — spotlight cards must fit on the same page
  - Use colored HTML for vessel names:
      critical  → <strong style="color:#D2393C;">VESSEL</strong>
      improved  → <strong style="color:#439B38;">VESSEL</strong>
      excluded  → <strong style="color:#667085;">VESSEL</strong>
  - Every number must exactly match the script output

  SPOTLIGHTS — one vessel per category, script data only:
  - MOST CRITICAL:      highest critical% or worst overall condition
  - STALLED COMPLIANCE: overdue_prev available and overdue >= overdue_prev (no improvement or worse),
                        especially if low_alarm active.
                        If no prior month data → pick vessel with most overdue + low_alarm.
  - MOST IMPROVED:      only if prior month data available → vessel with largest (overdue_prev - overdue).
                        If no prior month data → replace this card with a second BEST COMPLIANCE.
  - BEST COMPLIANCE:    lowest overdue, best condition profile
  - Detail: 1 short sentence, exact numbers from script only.
            If trend available: include delta (e.g. "overdue dropped from 41 → 34").

  PRIORITY RANKING — all vessels, 1 to N:
  - Rank strictly by severity using script data only
  - Levels: URGENT / HIGH / MED-HIGH / MEDIUM / LOW / N/A
  - key_issue: 1 sentence, factual, numbers from script only
  - Newly activated with no data → rank last, level N/A, dimmed: true
  - Ranking MUST be consistent with exec summary narrative
    (whoever is rank 1 must be the top concern in exec summary)

  | Level    | Criteria                                              |
  |----------|-------------------------------------------------------|
  | URGENT   | Zero uploads + low usage alarm active                 |
  | HIGH     | 40+ overdue, 8+ months onboard                        |
  | MED-HIGH | Newly deployed with critical equipment OR 80+ overdue |
  | MEDIUM   | 1%+ critical or 5%+ alert, moderate overdue           |
  | LOW      | Improving trend, low overdue                          |
  | N/A      | months=0 at runtime — enrolled this reporting period. |
  |          | Always rank last, dimmed: true, regardless of overdue |
  |          | count or alarm status. Do not use for any spotlight.  |

  RECOMMENDATIONS — 3-6 items, ordered by urgency:
  - Reference specific vessel names and exact numbers from script output
  - Levels: URGENT / HIGH / MEDIUM / LOW / INFO

#### 4d. Cross-check before proceeding
  Verify all of the following before moving on:
  [ ] Every number in exec paragraphs matches script output exactly
  [ ] Rank 1 vessel matches top concern in exec summary
  [ ] Spotlight MOST CRITICAL matches rank 1 vessel
  [ ] date_ym and date_label reflect the current month at runtime
  [ ] exec_paragraphs has exactly 2 entries, max 8 sentences total
  [ ] No numbers appear that cannot be found in script output

#### 4e. Build the intelligence JSON and save it
  Save as: C:\Fleet Overview Reporting\Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_intelligence.json

  {
    "fleet": "EXAMPLE",
    "date_ym": "2026-05",
    "date_label": "May 2026",
    "vessel_count": "10",
    "html_path": "C:\\Fleet Overview Reporting\\Reports\\EXAMPLE\\EXAMPLE_2026-05_staging.html",
    "exec_paragraphs": ["paragraph 1 HTML", "paragraph 2 HTML"],
    "spotlights": [
      {"label": "MOST CRITICAL",      "vessel": "...", "detail": "..."},
      {"label": "STALLED COMPLIANCE", "vessel": "...", "detail": "..."},
      {"label": "MOST IMPROVED",      "vessel": "...", "detail": "..."},
      {"label": "BEST COMPLIANCE",    "vessel": "...", "detail": "..."}
    ],
    "priorities": [
      {"rank": 1, "vessel": "...", "level": "URGENT", "key_issue": "...", "dimmed": false}
    ],
    "recommendations": [
      {"level": "URGENT", "text": "..."}
    ]
  }

#### 4f. Inject intelligence into HTML
  python3 "C:\Fleet Overview Reporting\Skills\inject_report.py" \
    --json "C:\Fleet Overview Reporting\Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_intelligence.json" \
    --out  "C:\Fleet Overview Reporting\Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_final.html"

#### 4g. Convert to PDF
  python3 "C:\Fleet Overview Reporting\Skills\print_pdf.py" \
    "C:\Fleet Overview Reporting\Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_final.html" \
    "C:\Fleet Overview Reporting\Reports\[COMPANY]\Fleet_CM_Report_[COMPANY]_[YYYY-MM].pdf"

#### 4h. Clean up
  Delete: Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_staging.html
  Delete: Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_intelligence.json
  Delete: Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_final.html
  Keep only the final PDF.
  If PDF generation fails, keep intermediate files for debugging.

#### 4i. Log the result
  Append to logs\runner.log:
  "[timestamp] DONE: [COMPANY] [YYYY-MM] → Reports\[COMPANY]\Fleet_CM_Report_[COMPANY]_[YYYY-MM].pdf"

### Step 5 — Print run summary
  Total files found: N
  Generated: N
  Skipped (already exist): N
  Errors: N (list any)

---

## Rules
- Always read Skills\SKILL.md before starting
- Never include api_pct anywhere in the report
- Never overwrite an existing PDF — skip and log instead
- Date is always the current month at runtime — never hardcode it
- Use script output as the only source of truth — never re-read Excel
- Always clean up intermediate files after successful PDF generation
- If a filename does not match Hat_CM_Summary_[COMPANY].xlsx — skip and log it