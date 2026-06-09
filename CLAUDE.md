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
    ACME.json          <- one JSON per company (cron_fetch.py output; carries tags)
    BETA_TANKERS.json
  APRIL 2026\ <- prior month (used for trend comparison)
    ACME.json
Reports\        <- Output PDFs, one subfolder per company (per fleet for split companies)
logs\           <- Run logs
Skills\         <- All scripts and template
Skills\template.html                  <- reference design, NEVER modify this file

## Skills folder contents
Skills\SKILL.md                       <- read this first, every run
Skills\report_builder.py              <- reads the company JSON, builds HTML
Skills\inject_report.py               <- injects intelligence into HTML
Skills\print_pdf.py                   <- HTML to PDF via Playwright (Chromium)
Skills\read_excel.py                  <- fallback reader for legacy xlsx/csv input
Skills\template.html                  <- reference design, never modify

---

## Run instructions

### Step 1 — Read SKILL.md
Always read Skills\SKILL.md before doing anything else, every single run.

### Step 2 — Determine the report date
The reporting month comes from the DATA, not the clock. A run in June over the
MAY 2026 folder is still a May 2026 report.
  month_folder    = the most recent "Input files\<MONTH YYYY>\" folder that actually
                    contains input files (e.g. "MAY 2026") — NOT the system month
  date_label      = that month as "Month YYYY"  (e.g. "May 2026")
  date_ym         = that month as "YYYY-MM"     (e.g. "2026-05")
  prev_month_folder = the calendar month before month_folder (e.g. "APRIL 2026"),
                    used for trend comparison
report_builder.py derives date_ym / date_label itself from the --summary path's month
folder and prints them ("Reporting month (from data): ..."). Use exactly those values
for the output filename and the skip check. Never use the run date; never hardcode.
(To force a specific month, pass --date-ym YYYY-MM.)

### Step 3 — Determine which companies to process

If the invocation prompt names a specific company (e.g. "Run the report for ACME"),
process ONLY that company. Do not touch any other files.

If no company is named, scan "C:\Fleet Overview Reporting\Input files\[month_folder]\"
for all [COMPANY].json files and process each one.

Company name comes from the JSON's "company_name" field (clean, with spaces):
  ACME.json → company = "ACME"   BETA_TANKERS.json → company = "BETA TANKERS"

Skip any company where the output PDF already exists (for split companies this check
is per-fleet — see Step 3b):
  C:\Fleet Overview Reporting\Reports\[COMPANY]\Fleet_CM_Report_[COMPANY]_[YYYY-MM].pdf

### Step 3b — Multi-fleet companies (per-fleet sub-reports)

Some companies group their vessels into fleets via the `tags` field in the input
JSON (one tag per vessel, e.g. "Fleet1", "GroupA"). For these, produce ONE report per
fleet instead of a single combined company report.

Detect automatically: if a company's input JSON has more than one distinct non-blank
`tags` value, treat it as multi-fleet and produce one report per tag.

For each distinct tag, run the pipeline exactly as for a normal company (Step 4: data
extractor → intelligence layer → inject → PDF), scoped to that fleet, with these
additions to the Step 4a report_builder.py call:
    --tag          "[TAG]"           keep only that fleet's vessels
    --fleet-label  "[Display]"       cover label; optional — auto-derived from the tag
                                     (Fleet1 → "Fleet 1", GroupA → "Group A")
The cover's italic line then shows the fleet label; everything else is unchanged.

Per-fleet output (one subfolder per fleet), used for both the --out staging file and
the final PDF / skip check:
  Reports\[COMPANY]\[TAG]\Fleet_CM_Report_[COMPANY]_[TAG]_[YYYY-MM].pdf
Skip, clean-up and logging are per-fleet (skip a fleet only if its own PDF exists).

If a multi-fleet company has blank-tag vessels that belong to a known fleet, add
`--include-untagged` to that fleet's run so they are included rather than dropped; the
dispatcher routine specifies any such cases.

### Step 4 — For each company file found:

#### 4a. Run the data extractor
  Prior month file (optional):
    prev_summary = "C:\Fleet Overview Reporting\Input files\[prev_month_folder]\[COMPANY_FILE].json"
    Check if this file exists. Include --summary-prev only if it does.
    If it does not exist → first run for this company → omit --summary-prev.

  python3 "C:\Fleet Overview Reporting\Skills\report_builder.py" \
    --summary      "C:\Fleet Overview Reporting\Input files\[month_folder]\[COMPANY_FILE].json" \
    --summary-prev "[prev_summary]" \                    ← omit entire line if file not found
    --template     "C:\Fleet Overview Reporting\Skills\template.html" \
    --fleet        "[COMPANY]" \
    --out          "C:\Fleet Overview Reporting\Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_staging.html"
  For a split company add --tag "[TAG]" (and --fleet-label, --include-untagged as in Step 3b),
  and write to the per-fleet subfolder. Create the output folder if it does not exist.
  report_builder.py derives the reporting month from the input's month folder and prints
  it ("Reporting month (from data): ...") — use that date_ym for filenames.

#### 4b. Read the script output — single source of truth
  The script prints all extracted vessel data.
  Use ONLY these numbers for all analysis and writing.
  Never re-read the input JSON directly.
  Never invent or estimate numbers.
  If a number cannot be traced to the script output, do not use it.

#### 4c. Write the intelligence layer

  EXECUTIVE SUMMARY — strict rules:
  - Maximum 2 short paragraphs, max 8 sentences total
  - Paragraph 1 (2-3 sentences, fixed structure):
      Sentence 1: "This report covers [N] [FLEET] vessels enrolled in the HAT AdViSe
                   condition monitoring program."
      Sentence 2: "Across [N] active vessels, the majority are in Normal condition, and
                   [N] vessels carry zero Critical or Alert findings."
                   (Do NOT open with "Overall fleet health is …". Do NOT use the word "sit".)
      Sentence 3 (optional): One sentence flagging the number of areas that warrant attention
                   this period, without naming vessels (e.g. "Two areas warrant attention this period."
                   or "Three vessels require urgent follow-up this period.").
      Keep this paragraph factual and brief — all detail goes in paragraph 2.
  - Paragraph 2 (3-4 sentences):
      Top concern with exact numbers, best performer with exact numbers,
      immediate priorities.
      If prior month data available: state the change "compared to the previous month"
      (e.g. "overdue fell to 34, compared to 41 the previous month").
      If no prior month data: note this is the first reporting cycle.
  - NO third paragraph — spotlight cards must fit on the same page
  - Wording rules:
      * Never use "only" before a number or percentage (no "only 65%") — state the figure plainly.
      * Say "CM Plan compliance", never bare "compliance".
      * Any month-over-month change is phrased "compared to the previous month".
  - Use colored HTML for vessel names:
      critical  → <strong style="color:#D2393C;">VESSEL</strong>
      improved  → <strong style="color:#439B38;">VESSEL</strong>
      excluded  → <strong style="color:#667085;">VESSEL</strong>
  - Every number must exactly match the script output

  SPOTLIGHTS — one vessel per category, script data only:
  - MOST CRITICAL:               highest critical% or worst overall condition
  - STALLED CM PLAN COMPLIANCE:  overdue_prev available and overdue >= overdue_prev (no improvement
                                 or worse), especially if low_alarm active.
                                 If no prior month data → pick vessel with most overdue + low_alarm.
  - MOST IMPROVED:               only if prior month data available → vessel with largest
                                 (overdue_prev - overdue). If no prior month data → replace this card
                                 with a second BEST CM PLAN COMPLIANCE.
  - BEST CM PLAN COMPLIANCE:     lowest overdue, best condition profile
  - Detail: 1 short sentence, exact numbers from script only.
            If trend available: state it "compared to the previous month"
            (e.g. "overdue fell to 34, compared to 41 the previous month").

  PRIORITY RANKING — all vessels, 1 to N:
  - Rank strictly by severity using script data only
  - Levels: URGENT / HIGH / MED-HIGH / MEDIUM / LOW / N/A
  - key_issue: 1 sentence, factual, numbers from script only
  - Newly activated with no data → rank last, level N/A, dimmed: true
  - Ranking MUST be consistent with exec summary narrative
    (whoever is rank 1 must be the top concern in exec summary)
  - The rank-1 / top-concern vessel is NEVER below HIGH. High condition severity is
    urgent even when uploads are current: Critical >= 5% is URGENT.

  | Level    | Criteria                                                        |
  |----------|-----------------------------------------------------------------|
  | URGENT   | Critical >= 5%, OR (zero uploads + low usage alarm active)       |
  | HIGH     | 40+ overdue with 8+ months onboard, OR rank-1 top concern        |
  | MED-HIGH | Newly deployed with critical equipment OR 80+ overdue           |
  | MEDIUM   | 1%+ critical or 5%+ alert, moderate overdue                     |
  | LOW      | Improving trend, low overdue                                    |
  | N/A      | months=0 at runtime — enrolled this reporting period.           |
  |          | Always rank last, dimmed: true, regardless of overdue           |
  |          | count or alarm status. Do not use for any spotlight.            |

  RECOMMENDATIONS — 3-6 items, ordered by urgency:
  - Reference specific vessel names and exact numbers from script output
  - Levels: URGENT / HIGH / MEDIUM / LOW / INFO

#### 4d. Cross-check before proceeding
  Verify all of the following before moving on:
  [ ] Every number in exec paragraphs matches script output exactly
  [ ] Rank 1 vessel matches top concern in exec summary
  [ ] Spotlight MOST CRITICAL matches rank 1 vessel
  [ ] date_ym and date_label reflect the reporting month from the data (the input month folder)
  [ ] exec_paragraphs has exactly 2 entries, max 8 sentences total
  [ ] No numbers appear that cannot be found in script output

#### 4e. Build the intelligence JSON and save it
  Save as: C:\Fleet Overview Reporting\Reports\[COMPANY]\[COMPANY]_[YYYY-MM]_intelligence.json

  {
    "fleet": "ACME",
    "date_ym": "2026-05",
    "date_label": "May 2026",
    "vessel_count": "10",
    "html_path": "C:\\Fleet Overview Reporting\\Reports\\ACME\\ACME_2026-05_staging.html",
    "exec_paragraphs": ["paragraph 1 HTML", "paragraph 2 HTML"],
    "spotlights": [
      {"label": "MOST CRITICAL",              "vessel": "...", "detail": "..."},
      {"label": "STALLED CM PLAN COMPLIANCE", "vessel": "...", "detail": "..."},
      {"label": "MOST IMPROVED",              "vessel": "...", "detail": "..."},
      {"label": "BEST CM PLAN COMPLIANCE",    "vessel": "...", "detail": "..."}
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
- The reporting month comes from the data (the input month folder), never the run
  date and never hardcoded — report_builder.py derives and prints it; --date-ym overrides
- Use script output as the only source of truth — never re-read the input JSON
- Always clean up intermediate files after successful PDF generation
- Process only [COMPANY].json inputs — skip and log anything else
- Multi-fleet companies (more than one distinct `tags` value) must be split into one
  report per tag — never a single combined report (see Step 3b)