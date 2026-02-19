# Progress Log

---

## [2026-02-19] — Auto Insurance Dashboard: Data Generation + Enhanced Dashboard

**Task:** Generate 10K mock auto insurance data and build an interactive dashboard with showcase features for a manager demo.
**Status:** Done

### What was done
- Generated 10K customer records with realistic correlations (age/gender/car year affect premiums and losses)
- Fixed critical bug in original dashboard that only embedded 200 records (top claims) instead of all 10K
- Built enhanced dashboard with all 10K records, proper KPIs (70.4% loss ratio, 35.4% claims freq)
- Added Risk Heatmap: Age Group × Car Era matrix with color-coded loss ratios
- Added Premium & Risk Predictor: interactive what-if calculator with predicted premium, risk level, and claim probability
- Added animated KPI counters (smooth count-up on load and filter changes)
- Added smooth chart transitions (800ms animations)

### Files created/modified
- `generate_mock_data.py` — Python script generating 10K records with realistic correlations, outputs CSV + Excel + JSON
- `auto_insurance_data.csv` — 10K records in CSV format
- `auto_insurance_data.xlsx` — 10K records in Excel format
- `dashboard_data.json` — Compact JSON (294KB) for dashboard embedding
- `auto_insurance_dashboard.html` — Single-file interactive dashboard with Chart.js, heatmap, predictor, animated KPIs

### Decisions made
- Used compact array format [id, gender_int, age, year, premium, loss] to keep JSON under 300KB
- Embedded all 10K records client-side rather than pre-aggregating (enables proper filtering)
- Single self-contained HTML file for easy sharing and demo

### Issues / Notes
- Original dashboard only had top 200 claims embedded, causing 910% loss ratio and 100% claims frequency (misleading)
- Fixed by embedding full 10K dataset and computing aggregations client-side

---

## [2026-02-19] — Auto Insurance Presentation (PPTX)

**Task:** Create a professional PowerPoint presentation from dashboard data to demo AI capabilities
**Status:** Done

### What was done
- Created 9-slide PPTX presentation with dark theme matching the dashboard
- Slides: Title, Executive Summary (KPIs), Loss Ratio by Age (chart), Claims Freq + Avg Premium (dual chart), Risk Heatmap (color-coded table), Gender Analysis (grouped bar chart), AI Predictor Showcase, What AI Agents Can Do, Thank You
- All charts use real aggregated data from the 10K dataset
- Risk heatmap uses color-coded cells (green to red based on loss ratio)
- Added "What AI Agents Can Do" slide showing 4 capabilities with time savings

### Files created/modified
- `create_presentation.py` — Python script generating the PPTX using python-pptx
- `Auto_Insurance_AI_Demo.pptx` — Final 9-slide presentation (70KB)

### Decisions made
- Used python-pptx instead of html2pptx because Node.js was not installed
- Dark theme (#0F1117 background) to match the dashboard aesthetic
- Widescreen 16:9 format for modern projectors

### Issues / Notes
- LibreOffice not installed, so thumbnail verification was not possible
- User should open in PowerPoint to verify visual quality

---

## [2026-02-19] — Auto Insurance PDF Actuarial Report

**Task:** Create a professional 6-page PDF actuarial report with matplotlib charts and reportlab layout
**Status:** Done

### What was done
- Created a comprehensive 6-page PDF report analyzing the 10K auto insurance portfolio
- Page 1: Cover page with dark navy background, title, subtitle, accent stripes
- Page 2: Executive summary with KPI table (5 metrics), narrative overview, 5 key findings bullets
- Page 3: Age group analysis with two matplotlib charts (loss ratio + claims frequency by age)
- Page 4: Risk segmentation with heatmap (Age Group x Car Era) + line chart (LR by car era)
- Page 5: Gender & premium analysis with grouped bar chart, gender comparison table, premium histogram
- Page 6: Six actionable recommendations with projected impact table and disclaimer
- All 6 charts use dark theme (navy backgrounds, white text, colored bars/lines)
- Charts include data labels, reference lines (portfolio average), color-coded risk levels
- AI-generated narrative text throughout (not placeholder) with real computed statistics
- Page numbers in footer (pages 2-6), report title in footer

### Files created/modified
- `create_pdf_report.py` — Python script generating the PDF using matplotlib + reportlab (1145 lines)
- `Auto_Insurance_Report.pdf` — Final 6-page PDF report (491 KB)

### Decisions made
- Used matplotlib for chart generation (saved as 180 DPI PNGs, then embedded in PDF via reportlab)
- Used BaseDocTemplate with separate page templates for cover vs content pages
- NumberedCanvas subclass for page numbering without numbering the cover page
- Temporary chart directory (`_report_charts/`) auto-cleaned after PDF generation
- Chart colors: green/amber/red for risk levels, indigo/cyan/pink for categories
- Letter page size (8.5x11) with 0.65in margins

### Issues / Notes
- Initial run had FormatStrFormatter error with `$%,.0f` (old-style % formatting doesn't support comma separator); fixed with FuncFormatter
- First layout had 7 pages due to content overflow on gender/premium page; fixed by reducing chart heights and tightening spacing
- PDF rendering cannot be visually verified in this environment (no pdftoppm); user should open in a PDF reader to confirm visual quality

---

## [2026-02-19] — PDF Report + Fraud Detection + NL Chatbot

**Task:** Build 3 additional showcase demos: PDF actuarial report, fraud detection script, and NL data Q&A chatbot
**Status:** Done

### What was done

**1. PDF Actuarial Report (6 pages)**
- Cover page with dark theme
- Executive summary with KPI table + key findings
- Age group analysis with matplotlib charts (loss ratio + claims frequency)
- Risk heatmap (age x car era) + car era trend line chart
- Gender & premium analysis with grouped bars + distribution histogram
- Recommendations page with 6 data-driven strategies

**2. Fraud Detection Script**
- Score-based system with 5 detection rules
- 95 flagged customers out of 3,539 claims (2.7% flag rate)
- Rules: extreme loss ratio, statistical outlier, new car high loss, young driver extreme claim, premium-loss mismatch
- Outputs CSV of flagged claims + dark-themed HTML summary report

**3. NL Data Q&A Chatbot**
- Flask app with Claude API integration
- Dark-themed chat interface with suggestion chips
- API key input in browser (user's key never stored)
- Pre-computed data context sent to Claude for accurate answers

### Files created/modified
- `create_pdf_report.py` — PDF generation script
- `Auto_Insurance_Report.pdf` — 6-page actuarial report (491 KB)
- `fraud_detection.py` — Anomaly detection script
- `flagged_claims.csv` — 95 flagged suspicious claims
- `fraud_detection_summary.html` — Interactive fraud report with charts
- `data_chatbot.py` — Flask chatbot app (run with `python data_chatbot.py`)

### Decisions made
- Used reportlab + matplotlib for PDF (no external PDF tools available)
- Score-based fraud detection (simple, explainable, demo-friendly)
- Flask backend for chatbot (avoids CORS issues with direct API calls)
- Claude Sonnet for chatbot responses (fast + cost-effective for Q&A)

### Issues / Notes
- Chatbot requires user to enter their own Anthropic API key in the browser UI
- PDF uses pdftoppm-less rendering approach since LibreOffice is not installed

---

## 2026-02-19 — V2 Presentation with New Feature Slides

**Task:** Extend presentation to include fraud detection, chatbot, and PDF report slides as a V2 version
**Status:** Done

### What was done
- Added Slide 8: Fraud Detection (5 KPI cards, detection rules table, top flagged case, score distribution chart)
- Added Slide 9: NL Data Chatbot (chat mockup with 3 Q&A pairs, "How It Works" steps, example questions)
- Added Slide 10: PDF Report Showcase (6 page cards showing report structure)
- Updated Slide 11: "What AI Can Do" expanded from 4 to 7 capability cards
- Updated Slide 12: Thank You slide text references all deliverables
- Changed output filename to Auto_Insurance_AI_Demo_V2.pptx (V1 preserved)

### Files created/modified
- `create_presentation.py` — Added 3 new slides (fraud, chatbot, PDF report), expanded capabilities slide, updated thank you, changed output to V2
- `Auto_Insurance_AI_Demo_V2.pptx` — Generated V2 presentation (12 slides, 85KB)

### Decisions made
- Kept V1 untouched per user request, output to separate V2 filename
- Used compact card layouts to fit rich content per slide

### Issues / Notes
- Session was interrupted by connection loss mid-edit; resumed and completed successfully
- LibreOffice not available for thumbnail verification, verified via file size and slide count
