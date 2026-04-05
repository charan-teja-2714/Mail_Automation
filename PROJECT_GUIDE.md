# Weekly Sanity Check Report — Project Guide

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Project Structure](#2-project-structure)
3. [How to Run](#3-how-to-run)
4. [File-by-File Reference](#4-file-by-file-reference)
   - [main.py](#mainpy--entry-point)
   - [json_loader.py](#json_loaderpy--data-loading--validation)
   - [utils.py](#utilspy--shared-constants--helpers)
   - [html_generator.py](#html_generatorpy--email-body-builder)
   - [doc_generator.py](#doc_generatorpy--docx-report-builder)
   - [mail_sender.py](#mail_senderpy--outlook-email-sender)
   - [data.json](#datajson--input-data)
5. [Customisation Reference](#5-customisation-reference)
6. [Data JSON Field Reference](#6-data-json-field-reference)
7. [Status Values Reference](#7-status-values-reference)
8. [Dependencies](#8-dependencies)
9. [Troubleshooting](#9-troubleshooting)

---

## 1. Project Overview

This project automates the generation and delivery of a **Weekly Sanity Check Report**. Each week, you update `data.json` with the current status of monitored activities, run `python main.py`, and the script:

1. Reads and validates `data.json`
2. Builds a colour-coded HTML table for the email body
3. Generates a formatted Word document (DOCX) with descriptions and screenshots
4. Sends the email via your locally installed Microsoft Outlook (no passwords needed)

**Platform:** Windows only (uses Outlook COM automation via `pywin32`)

---

## 2. Project Structure

```
Mail Automation/
├── main.py              # Entry point — orchestrates all steps
├── json_loader.py       # Reads and validates data.json
├── utils.py             # Shared colours, status maps, path helpers
├── html_generator.py    # Builds the HTML email body
├── doc_generator.py     # Creates the DOCX Word report
├── mail_sender.py       # Sends email via Outlook (win32com)
├── data.json            # Your weekly activity data (edit this weekly)
├── requirements.txt     # Python package dependencies
├── images/              # Folder for screenshot PNGs (create manually)
│   ├── st22.png
│   ├── sm37.png
│   └── ...
└── Weekly_Sanity_Report.docx   # Auto-generated output (do not edit)
```

---

## 3. How to Run

### First-time setup
```cmd
pip install -r requirements.txt
```

### Every week
1. Update `data.json` with new statuses and descriptions
2. Drop new screenshots into the `images/` folder
3. Run:
```cmd
python main.py
```

### Expected console output
```
============================================================
  Weekly Sanity Check Report – Automation Started
============================================================

[1/4] Loading activity data from JSON...
[2/4] Generating HTML email body...
[3/4] Generating DOCX report...
[4/4] Sending email via local Outlook application...

============================================================
  All steps completed successfully!
============================================================
```

---

## 4. File-by-File Reference

---

### `main.py` — Entry Point

**What it does:**
Ties all modules together. Reads config (recipients, file paths), calls each module in sequence, and exits with a clear error message if any step fails.

**Execution flow:**
```
load_activities()  →  generate_html_body()  →  generate_docx()  →  send_report()
```

**Where to make changes:**

| What to change | Location |
|---|---|
| TO email recipients | Lines 24–26 — edit the `TO_RECIPIENTS` list |
| CC email recipients | Lines 28–30 — edit the `CC_RECIPIENTS` list |
| Email subject line | Line 33 — change `EMAIL_SUBJECT` |
| Input JSON filename | Line 18 — change `JSON_FILE` |
| Output DOCX filename | Line 21 — change `DOCX_OUTPUT` |

**Example — adding more recipients:**
```python
TO_RECIPIENTS = [
    "charantej2714@outlook.com",
    "colleague@company.com",       # add more here
]

CC_RECIPIENTS = [
    "charan.mamidi@capgemini.com",
    "another.manager@capgemini.com",
]
```

---

### `json_loader.py` — Data Loading & Validation

**What it does:**
Opens `data.json`, parses it, and validates every record. Skips malformed records with a warning instead of crashing. Normalises the `status` field (e.g. `"ok"` → `"Done"`) and fills in default values for optional fields.

**Required fields per record:** `sno`, `activity`, `status`
**Optional fields:** `doc_title`, `doc_description`, `image`

**Key behaviour:**
- If `doc_title` is missing → falls back to the value of `activity`
- If `doc_description` is missing → uses `"No description provided."`
- If `image` is missing or `null` → image section is skipped in DOCX
- If `sno` is not a number → auto-assigned from record position

**Where to make changes:**

| What to change | Location |
|---|---|
| Add a new required field | Line 10 — add to `REQUIRED_FIELDS` set |
| Add a new optional field | Line 13 — add to `OPTIONAL_FIELDS` set |
| Change default description text | Line 74 — `setdefault("doc_description", ...)` |

---

### `utils.py` — Shared Constants & Helpers

**What it does:**
The single source of truth for status colours and mappings, used by both `html_generator.py` and `doc_generator.py`. Also provides path resolution and image existence checks.

**Defined constants:**

| Constant | Used by | Purpose |
|---|---|---|
| `STATUS_VARIANTS` | `json_loader.py` | Maps raw strings to canonical status |
| `STATUS_COLORS_HTML` | `html_generator.py` | Text/badge foreground hex colour |
| `STATUS_BG_HTML` | `html_generator.py` | Badge background hex colour |
| `STATUS_COLORS_DOCX` | `doc_generator.py` | RGB tuple for Word document text |

**Where to make changes:**

| What to change | Location |
|---|---|
| HTML badge text colour (Done/Warning/Failed) | Lines 33–38 — `STATUS_COLORS_HTML` |
| HTML badge background colour | Lines 41–46 — `STATUS_BG_HTML` |
| DOCX text colour for status | Lines 49–54 — `STATUS_COLORS_DOCX` |
| Accept new status alias (e.g. `"complete"` → Done) | Lines 17–30 — add to `STATUS_VARIANTS` |

**Example — change Warning colour to dark red instead of orange:**
```python
STATUS_COLORS_HTML = {
    "Done":    "#1a7a1a",
    "Warning": "#8B0000",   # changed to dark red
    "Failed":  "#cc0000",
    "N/A":     "#555555",
}
```

---

### `html_generator.py` — Email Body Builder

**What it does:**
Generates the complete HTML string sent as the email body. Contains:
- Greeting paragraph (`Dear Team,`)
- Colour-coded summary table (S.No / Activity / Status)
- Closing paragraph with Regards sign-off

**Table structure:**

| Column | Width | Style |
|---|---|---|
| S.No | 50px | Centred, grey text |
| Activity | Auto | Left-aligned |
| Status | 100px | Coloured badge (rounded pill) |

**Where to make changes:**

| What to change | Location |
|---|---|
| Table header background colour | Line 21 — `background-color:#003366` in `_TH_STYLE` |
| Table header text colour | Line 22 — `color:#ffffff` in `_TH_STYLE` |
| Table font and size | Lines 15–17 — `_TABLE_STYLE` |
| Cell padding | Line 29 — `padding:9px 14px` in `_TD_STYLE_BASE` |
| Alternating row colours | Lines 33–34 — `_ROW_EVEN_BG` / `_ROW_ODD_BG` |
| Status badge border-radius (pill shape) | Line 117 — `border-radius:4px` |
| Status badge font size | Line 119 — `font-size:12px` |
| Greeting text | Line 54 — `Dear Team,` |
| Opening paragraph | Lines 57–60 |
| Closing paragraph | Lines 65–69 |
| Sign-off name | Line 73 — `Sanity Check Automation` |
| Max table width | Line 15 — `max-width:650px` |
| Add a new column (e.g. Remarks) | `_build_table()` from line 84 — add `<th>` in header and `<td>` in row loop |

**Example — change header to dark green:**
```python
_TH_STYLE = (
    "background-color:#1a5c1a;"   # was #003366
    "color:#ffffff;"
    "padding:10px 14px;"
    "text-align:left;"
    "letter-spacing:0.5px;"
)
```

---

### `doc_generator.py` — DOCX Report Builder

**What it does:**
Creates a formatted Word document using `python-docx`. The document contains:
1. A title page section: **"Weekly Sanity Check Report"**
2. An **Executive Summary** table (same columns as the email)
3. A page break
4. One detailed section per activity containing:
   - Section heading (`doc_title`)
   - Status with colour (green / orange / red)
   - Description paragraph (`doc_description`)
   - Screenshot image (if found in `images/` folder)
   - Horizontal divider line

**Page layout:**
- Page size: A4 equivalent margins (1.0" top/bottom, 1.1" left/right)
- Image max width: 5.5 inches (auto-capped to fit margins)

**Where to make changes:**

| What to change | Location |
|---|---|
| Image max width in DOCX | Line 18 — `IMAGE_MAX_WIDTH = Inches(5.5)` |
| Spacing between sections | Line 19 — `SECTION_SPACING = Pt(8)` |
| Document title text | Line 43 — `"Weekly Sanity Check Report"` |
| Title colour (navy) | Line 45 — `RGBColor(0, 51, 102)` |
| Subtitle text | Line 49 — `"Automated System Health Overview"` |
| Summary table heading | Line 56 — `"Executive Summary"` |
| Summary table columns | Lines 62–65 — change column headers |
| Section heading colour | Line 99 — `RGBColor(0, 51, 102)` |
| Status label prefix | Line 103 — `"Status: "` |
| Divider line colour | Line 128 — `w:color` value `"AAAAAA"` |
| Divider line thickness | Line 126 — `w:sz` value `"4"` (in eighths of a point) |

**Example — change image max width to 4 inches:**
```python
IMAGE_MAX_WIDTH = Inches(4.0)   # line 18
```

---

### `mail_sender.py` — Outlook Email Sender

**What it does:**
Uses Windows COM automation (`win32com.client`) to connect to the locally running Microsoft Outlook application and send the composed email. No passwords or SMTP configuration needed — it uses whatever account is already signed in to Outlook.

**How it works:**
1. Validates recipients list and attachment path
2. Calls `win32com.client.Dispatch("Outlook.Application")`
3. Creates a new mail item (`olMailItem = 0`)
4. Sets TO, CC, Subject, HTMLBody, and Attachment
5. Calls `mail.Send()` — email appears in Outlook Sent Items

**Key behaviours:**
- Tries `Dispatch` first; falls back to `EnsureDispatch` (cache rebuild) if it fails
- Sent emails appear in your **Outlook Sent Items** folder
- Works with personal Outlook.com accounts and corporate Microsoft 365 accounts
- Outlook must be open or at minimum installed and configured

**Where to make changes:**

| What to change | Location |
|---|---|
| Default email subject | Line 33 — `DEFAULT_SUBJECT` |
| Add BCC recipients | After line 90 — add `mail.BCC = "; ".join(bcc_list)` |
| Set email importance (High) | After line 94 — add `mail.Importance = 2` (0=Low, 1=Normal, 2=High) |
| Request read receipt | After line 94 — add `mail.ReadReceiptRequested = True` |

**Example — set high importance:**
```python
mail.Subject  = subject
mail.HTMLBody = html_body
mail.Importance = 2          # add this line for High importance
```

---

### `data.json` — Input Data

**What it does:**
The only file you need to edit weekly. Contains the list of sanity check activities with their current status, descriptions, and optional screenshot paths.

**Full field reference:**

| Field | Type | Required | Description |
|---|---|---|---|
| `sno` | integer | Yes | Serial number (used for ordering in table) |
| `activity` | string | Yes | Short activity name shown in email table |
| `status` | string | Yes | One of: Done / Warning / Failed (see aliases below) |
| `doc_title` | string | No | Heading in the DOCX report (defaults to `activity` if omitted) |
| `doc_description` | string | No | Detailed description paragraph in DOCX |
| `image` | string or null | No | Relative path to screenshot PNG (e.g. `"images/st22.png"`) |

**Minimal valid record (only required fields):**
```json
{
  "sno": 1,
  "activity": "Check for short dumps in ST22",
  "status": "Done"
}
```

**Full record:**
```json
{
  "sno": 1,
  "activity": "Check for short dumps in ST22",
  "status": "Done",
  "doc_title": "Short Dump Monitoring (ST22)",
  "doc_description": "Checked system for runtime errors using ST22. No critical dumps found.",
  "image": "images/st22.png"
}
```

---

## 5. Customisation Reference

### Change email recipients → `main.py` lines 24–30
```python
TO_RECIPIENTS = ["recipient@company.com"]
CC_RECIPIENTS = ["manager@company.com"]
```

### Change email subject → `main.py` line 33
```python
EMAIL_SUBJECT = "SAP Weekly Health Check – April 2026"
```

### Change table header colour → `html_generator.py` line 21
```python
"background-color:#003366;"   # any hex colour
```

### Change status badge colours → `utils.py` lines 33–46
```python
STATUS_COLORS_HTML = { "Done": "#1a7a1a", ... }
STATUS_BG_HTML     = { "Done": "#d4edda", ... }
```

### Change DOCX image size → `doc_generator.py` line 18
```python
IMAGE_MAX_WIDTH = Inches(5.0)
```

### Add a new status type (e.g. "Pending") → `utils.py`
```python
# Line 17 — STATUS_VARIANTS
"pending": "Pending",

# Line 33 — STATUS_COLORS_HTML
"Pending": "#0056b3",   # blue text

# Line 41 — STATUS_BG_HTML
"Pending": "#cce5ff",   # light blue background

# Line 49 — STATUS_COLORS_DOCX
"Pending": (0, 86, 179),
```

### Add a Remarks column to the email table → `html_generator.py`
In `_build_table()` (line 84):
1. Add `<th>` for Remarks in the header block (around line 93)
2. Add a `remarks_td` variable in the row loop using `item.get("remarks", "-")`
3. Append it to the `<tr>` string

Then add `"remarks"` field to each record in `data.json`.

### Add BCC to emails → `mail_sender.py`
In `send_report()`, after line 90 add:
```python
if bcc_list:
    mail.BCC = "; ".join(bcc_list)
```
And update the function signature to accept `bcc_recipients`.

---

## 6. Data JSON Field Reference

```
data.json
│
├── sno              → Row number in email table and DOCX summary
├── activity         → Short label shown in email table
├── status           → Drives colour coding everywhere
├── doc_title        → H2 heading in DOCX per-activity section
├── doc_description  → Body paragraph in DOCX per-activity section
└── image            → Path to PNG screenshot embedded in DOCX
                       Use null or omit to skip image
```

---

## 7. Status Values Reference

The `status` field in `data.json` is case-insensitive. Accepted values:

| You write | Normalised to | Email badge | DOCX text |
|---|---|---|---|
| `done`, `ok`, `pass`, `passed` | **Done** | Green | Green |
| `warning`, `warn`, `caution` | **Warning** | Orange | Orange |
| `failed`, `fail`, `error` | **Failed** | Red | Red |
| `na`, `n/a` | **N/A** | Grey | Grey |

---

## 8. Dependencies

| Package | Version | Purpose |
|---|---|---|
| `python-docx` | >=1.1.0 | Create and format the `.docx` report |
| `pywin32` | >=306 | Connect to Microsoft Outlook via Windows COM |

Install all:
```cmd
pip install -r requirements.txt
```

---

## 9. Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `Invalid class string` | Classic Outlook not installed | Install Microsoft Office / Outlook desktop app |
| `Image not found, skipping` | PNG file missing from `images/` folder | Add the image file or set `"image": null` in JSON |
| `JSON file not found` | `data.json` missing or wrong path | Check `JSON_FILE` in `main.py` line 18 |
| `No valid activities found` | All records failed validation | Check that `sno`, `activity`, `status` exist in every record |
| `Attachment not found` | DOCX wasn't generated | Check Step 3 output for DOCX errors |
| Outlook security prompt | Windows COM security policy | Click **Allow** when prompted |
