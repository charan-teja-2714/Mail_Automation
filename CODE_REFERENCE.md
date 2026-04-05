# Code Reference — Function-by-Function Breakdown

This document explains every function in every file of the Weekly Sanity Check Report project.
Use this alongside `PROJECT_GUIDE.md` to understand not just *where* to make changes, but *why* the code is written the way it is.

---

## Table of Contents

1. [main.py](#1-mainpy)
2. [utils.py](#2-utilspy)
3. [json_loader.py](#3-json_loaderpy)
4. [html_generator.py](#4-html_generatorpy)
5. [doc_generator.py](#5-doc_generatorpy)
6. [mail_sender.py](#6-mail_senderpy)

---

## 1. `main.py`

This file has no reusable functions — it contains one top-level function `main()` that acts as the controller for the entire automation pipeline.

---

### `main()` — Lines 39–100

```python
def main():
```

**Purpose:**
The entry point and orchestrator. When you run `python main.py`, Python calls this function. It runs all four steps in sequence and stops immediately with a clear error message if any step fails, so you always know which step broke and why.

**What it does step by step:**

```
Step 1 → json_loader.load_activities()    reads and validates data.json
Step 2 → html_generator.generate_html_body()  builds the email HTML
Step 3 → doc_generator.generate_docx()   creates the Word document
Step 4 → mail_sender.send_report()       sends via Outlook
```

**Key design decisions:**
- Each step is wrapped in its own `try/except` block so a failure in step 2 doesn't silently affect step 3
- `sys.exit(1)` is called on any fatal error — this signals to Windows/CI that the script failed (exit code 1 = error)
- File paths are resolved using `os.path.abspath(__file__)` so the script works correctly regardless of which folder you run it from

**Flow diagram:**
```
python main.py
      │
      ▼
 resolve paths (base_dir, json_path, docx_path)
      │
      ▼
 load_activities(json_path)
      │  ← FileNotFoundError or ValueError → print FATAL + exit(1)
      ▼
 generate_html_body(activities)
      │  ← any Exception → print FATAL + exit(1)
      ▼
 generate_docx(activities, docx_path)
      │  ← RuntimeError → print FATAL + exit(1)
      ▼
 send_report(html_body, docx_file, TO, CC, subject)
      │  ← RuntimeError → print FATAL + exit(1)
      ▼
 "All steps completed successfully!"
```

---

## 2. `utils.py`

Contains no class definitions. Provides shared constants and three utility functions used across the project. All other modules import from here.

---

### Module-level constants

#### `STATUS_VARIANTS` — Line 17
```python
STATUS_VARIANTS = { "done": "Done", "ok": "Done", ... }
```
A dictionary that maps every acceptable raw string a user might type in `data.json` to one of the four canonical statuses: `Done`, `Warning`, `Failed`, `N/A`.
Used by `normalize_status()`. The lookup is always done in lowercase so it is case-insensitive.

---

#### `STATUS_COLORS_HTML` — Line 33
```python
STATUS_COLORS_HTML = { "Done": "#1a7a1a", "Warning": "#b35900", ... }
```
Maps each canonical status to a hex colour string used as the **text/foreground colour** of the status badge in the HTML email table. Imported by `html_generator.py`.

---

#### `STATUS_BG_HTML` — Line 41
```python
STATUS_BG_HTML = { "Done": "#d4edda", "Warning": "#fff3cd", ... }
```
Maps each status to a hex colour used as the **background colour** of the status badge in the HTML email. Paired with `STATUS_COLORS_HTML` so text and background always have sufficient contrast.

---

#### `STATUS_COLORS_DOCX` — Line 49
```python
STATUS_COLORS_DOCX = { "Done": (0, 128, 0), "Warning": (204, 102, 0), ... }
```
Maps each status to an RGB tuple (Red, Green, Blue values 0–255) used to colour the status text inside the Word document. `python-docx` uses `RGBColor(r, g, b)` objects, so tuples are used here and unpacked with `*` when passed to `RGBColor`.

---

### `normalize_status(raw_status)` — Line 57

```python
def normalize_status(raw_status: str) -> str:
```

**Purpose:**
Accepts whatever the user typed in the `"status"` field of `data.json` and returns the correct canonical form.

**How it works:**
1. Returns `"N/A"` immediately if the input is not a string at all (e.g. `null` in JSON becomes `None` in Python)
2. Strips whitespace and converts to lowercase
3. Looks up the cleaned string in `STATUS_VARIANTS`
4. If found, returns the canonical value (e.g. `"pass"` → `"Done"`)
5. If not found, returns the original stripped string as-is (so unknown statuses pass through instead of crashing)

**Called by:** `json_loader._validate_record()`

**Examples:**
```
"done"    → "Done"
"PASS"    → "Done"
"warn"    → "Warning"
"FAILED"  → "Failed"
"ok"      → "Done"
"xyz"     → "xyz"   (unknown, passed through as-is)
None      → "N/A"
```

---

### `resolve_path(path, base_dir)` — Line 68

```python
def resolve_path(path: str, base_dir: str = None) -> str:
```

**Purpose:**
Converts a relative file path into an absolute one, anchored to the project root directory. Prevents the "file not found" errors that happen when you run the script from a different working directory.

**How it works:**
- If the given path is already absolute (starts with a drive letter or `/`), it is returned unchanged
- Otherwise, it is joined with `base_dir` (defaults to the folder containing `utils.py`)

**Called by:** `image_exists()`

---

### `image_exists(image_path, base_dir)` — Line 80

```python
def image_exists(image_path: str, base_dir: str = None) -> str | None:
```

**Purpose:**
Checks whether a screenshot file actually exists on disk before trying to insert it into the Word document. Returns the absolute path if found, or `None` if missing — along with a warning log message so you know which image was skipped.

**How it works:**
1. Calls `resolve_path()` to get the absolute path
2. Uses `os.path.isfile()` to check existence
3. Returns the absolute path on success, `None` on failure

**Called by:** `doc_generator._add_activity_section()`

**Why this matters:**
Without this check, `python-docx` would raise an unhandled exception when it tries to open a missing image file, crashing the entire DOCX generation step.

---

## 3. `json_loader.py`

Responsible for reading `data.json` from disk, parsing it, and returning a clean list of validated activity dictionaries.

---

### `load_activities(json_path)` — Line 16

```python
def load_activities(json_path: str) -> list[dict]:
```

**Purpose:**
The public-facing function called by `main.py`. Opens the JSON file, parses it, and delegates validation of each individual record to `_validate_record()`.

**How it works:**
1. Checks the file exists — raises `FileNotFoundError` if not
2. Opens and parses the file with `json.load()` — raises `ValueError` on invalid JSON syntax
3. Confirms the top-level structure is a list — raises `ValueError` if it's a dict or anything else
4. Loops over every record, calls `_validate_record()` on each
5. Collects only the records that passed validation (non-None results)
6. Returns the final clean list

**Returns:** `list[dict]` — each dict is a validated, normalised activity record

**Raises:**
- `FileNotFoundError` — file path does not exist
- `ValueError` — file contains invalid JSON, or JSON root is not an array

---

### `_validate_record(record, position)` — Line 51

```python
def _validate_record(record: dict, position: int) -> dict | None:
```

**Purpose:**
Validates and cleans a single record from the JSON array. The leading underscore `_` is a Python convention meaning this is a private/internal function — it is only called by `load_activities()`, not from outside the module.

**How it works step by step:**

| Step | What it checks | On failure |
|---|---|---|
| 1 | Is the record a dict (object)? | Logs warning, returns `None` (skip) |
| 2 | Are all required fields present? | Logs warning with missing fields, returns `None` (skip) |
| 3 | Normalise `status` via `normalize_status()` | Always succeeds (unknown values pass through) |
| 4 | Fill `doc_title` default | Uses `activity` value if missing |
| 5 | Fill `doc_description` default | Uses `"No description provided."` if missing |
| 6 | Fill `image` default | Uses `None` if missing |
| 7 | Convert `sno` to integer | Falls back to `position` if conversion fails |

**Returns:** The cleaned `dict`, or `None` if the record should be skipped

**Why it returns None instead of raising:**
So that one bad record in the JSON doesn't abort processing of all the valid records. The warning in the log tells you exactly which record failed.

---

## 4. `html_generator.py`

Builds the full HTML string that becomes the Outlook email body. Uses only Python's built-in string formatting — no external HTML library required.

---

### Module-level style constants — Lines 12–34

```python
_TABLE_STYLE   = "border-collapse:collapse; ..."
_TH_STYLE      = "background-color:#003366; ..."
_TD_STYLE_BASE = "padding:9px 14px; ..."
_ROW_EVEN_BG   = "#f8f9fa"
_ROW_ODD_BG    = "#ffffff"
```

**Purpose:**
All CSS is defined as module-level string constants at the top so you can change the visual design without hunting through the function bodies. These are **inline CSS** strings (not a stylesheet) because Outlook strips `<style>` tags and external stylesheets.

**Why inline CSS:**
Microsoft Outlook's email renderer (based on Word's rendering engine) ignores `<style>` blocks and `class` attributes. Every style must be applied directly on the HTML element via the `style=""` attribute for it to render correctly.

---

### `generate_html_body(activities)` — Line 37

```python
def generate_html_body(activities: list[dict]) -> str:
```

**Purpose:**
The public function called by `main.py`. Assembles the complete HTML document string for the email body, including the greeting, the table, and the closing.

**How it works:**
1. Calls `_build_table()` to generate the `<table>` HTML string
2. Wraps it in a full `<html><body>...</body></html>` structure using an f-string
3. Embeds the greeting, table, and sign-off paragraphs
4. Returns the complete HTML string

**Returns:** A complete HTML string ready to assign to `mail.HTMLBody`

---

### `_build_table(activities)` — Line 84

```python
def _build_table(activities: list[dict]) -> str:
```

**Purpose:**
Builds the `<table>...</table>` HTML block with a header row and one data row per activity. Private function called only by `generate_html_body()`.

**How it works:**
1. Constructs the `<thead>` block with three `<th>` cells: S.No, Activity, Status
2. Loops over the activities list with `enumerate()` to track even/odd index for row striping
3. For each row:
   - Picks background colour based on even/odd index
   - Gets the status and looks up its foreground + background badge colour from `utils.py`
   - Builds three `<td>` strings: S.No (centred), Activity (plain), Status (coloured badge)
   - Assembles them into a `<tr>` string
4. Joins all rows and wraps with `<tbody>...</tbody>`
5. Returns the full table string

**Status badge structure per cell:**
```html
<td style="padding:9px 14px; border-bottom:1px solid #dee2e6; background-color:#f8f9fa;">
  <span style="display:inline-block; padding:3px 10px; border-radius:4px;
               background-color:#d4edda; color:#1a7a1a; font-weight:bold; font-size:12px;">
    Done
  </span>
</td>
```

---

### `_escape(text)` — Line 130

```python
def _escape(text: str) -> str:
```

**Purpose:**
Prevents HTML injection. If an activity name in `data.json` contains characters like `<`, `>`, `&`, or `"`, they would break the HTML structure or be interpreted as tags. This function replaces them with their safe HTML entity equivalents.

**Replacements performed:**

| Character | Replaced with |
|---|---|
| `&` | `&amp;` |
| `<` | `&lt;` |
| `>` | `&gt;` |
| `"` | `&quot;` |

**Called by:** `_build_table()` for every activity name and status string

---

## 5. `doc_generator.py`

Generates the formatted Word document using the `python-docx` library. The document has two parts: a summary table and a detailed per-activity section.

---

### Module-level layout constants — Lines 18–19

```python
IMAGE_MAX_WIDTH = Inches(5.5)
SECTION_SPACING = Pt(8)
```

**Purpose:**
`IMAGE_MAX_WIDTH` caps how wide an inserted screenshot can be. Without this, a large image would overflow the page margins. `SECTION_SPACING` controls the vertical gap after each activity's description paragraph.

---

### `generate_docx(activities, output_path)` — Line 22

```python
def generate_docx(activities: list[dict], output_path: str) -> str:
```

**Purpose:**
The public function called by `main.py`. Creates a new `Document` object and calls the private helper functions in order to build the full document, then saves it to disk.

**How it works:**
1. Creates a blank `Document()` object
2. Calls `_set_page_margins()` — sets A4-friendly margins on all sections
3. Calls `_add_document_title()` — adds the title and subtitle
4. Calls `_add_summary_table()` — adds the Executive Summary table
5. Calls `_add_page_break()` — separates summary from details
6. Loops over activities and calls `_add_activity_section()` for each
7. Saves the document with `doc.save(output_path)`
8. Returns the absolute path of the saved file

**Returns:** Absolute path string of the saved `.docx` file

**Raises:** `RuntimeError` wrapping any `python-docx` exception, so `main.py` can catch it cleanly

---

### `_set_page_margins(doc)` — Line 49

```python
def _set_page_margins(doc: Document):
```

**Purpose:**
Sets comfortable page margins on every section of the document. `python-docx` creates documents with very narrow default margins; this makes the output look more like a professional report.

**Margins applied:**
- Top / Bottom: 1.0 inch
- Left / Right: 1.1 inch

**Why it loops over sections:**
A Word document can have multiple sections (e.g. after a page break with a section break). The loop ensures margins are consistent throughout even if sections are added later.

---

### `_add_document_title(doc)` — Line 57

```python
def _add_document_title(doc: Document):
```

**Purpose:**
Adds the main report title ("Weekly Sanity Check Report") styled as a centred Heading 0 in dark navy, followed by a centred italic subtitle ("Automated System Health Overview") in grey.

**How it works:**
1. `doc.add_heading(text, level=0)` — level 0 is the Title style in Word
2. Iterates over the heading's `runs` to apply navy `RGBColor`
3. Adds a subtitle paragraph with smaller, italic, grey text
4. Adds a blank paragraph for breathing room

---

### `_add_summary_table(doc, activities)` — Line 70

```python
def _add_summary_table(doc: Document, activities: list[dict]):
```

**Purpose:**
Inserts an "Executive Summary" table at the top of the document showing S.No, Activity, and Status for all activities at a glance. The Status column uses coloured text (not badges) since Word tables handle colour differently from HTML.

**How it works:**
1. Adds a "Executive Summary" Heading 1
2. Creates a `doc.add_table(rows=1, cols=3)` with the built-in "Light List Accent 1" Word style
3. Populates the header row with bold text labels
4. Loops over activities and adds one row per activity
5. For the Status cell: clears default text, adds a `run`, sets `run.bold = True`, and applies `RGBColor` from `STATUS_COLORS_DOCX`

---

### `_add_activity_section(doc, item)` — Line 91

```python
def _add_activity_section(doc: Document, item: dict):
```

**Purpose:**
Adds one full detailed section to the document for a single activity. This is the most complex function in the project — it handles heading, status, description, image insertion, and the divider line.

**How it works:**

| Sub-step | What it adds |
|---|---|
| 1 | `doc.add_heading(doc_title, level=2)` — navy coloured H2 heading |
| 2 | Status paragraph: bold label `"Status: "` + bold coloured status value |
| 3 | Description paragraph from `doc_description` with `SECTION_SPACING` after |
| 4 | If `image` field is set: calls `image_exists()` to verify file exists |
| 5 | If image found: inserts centred picture with `IMAGE_MAX_WIDTH` cap |
| 6 | If image missing: skips silently (already logged by `image_exists()`) |
| 7 | Calls `_add_horizontal_line()` for a visual divider |
| 8 | Adds a blank paragraph for spacing |

**Image insertion detail:**
```python
run = pic_para.add_run()
run.add_picture(abs_img, width=IMAGE_MAX_WIDTH)
```
The image is added inside a `run` inside a centred `paragraph`. The `width` parameter scales the image proportionally — height is calculated automatically to preserve the aspect ratio.

---

### `_add_horizontal_line(doc)` — Line 117

```python
def _add_horizontal_line(doc: Document):
```

**Purpose:**
Inserts a thin grey horizontal rule between activity sections to visually separate them. Word does not have a direct "add horizontal line" API, so this function uses raw XML manipulation via the `OxmlElement` API from `python-docx`.

**How it works:**
1. Adds an empty paragraph
2. Gets or creates the paragraph's properties element (`<w:pPr>`)
3. Creates a `<w:pBdr>` (paragraph border) XML element
4. Creates a `<w:bottom>` border with style `"single"`, size `4` (in eighths of a point = 0.5pt), colour `#AAAAAA`
5. Appends the border element into the paragraph properties

**Why XML directly:**
`python-docx` doesn't expose paragraph borders through its Python API. Direct XML construction via `OxmlElement` is the standard workaround used in the `python-docx` community for features not yet in the high-level API.

---

### `_add_page_break(doc)` — Line 131

```python
def _add_page_break(doc: Document):
```

**Purpose:**
Inserts a page break so the detailed per-activity sections start on a fresh page after the Executive Summary table.

**How it works:**
Calls `doc.add_page_break()` — a single built-in `python-docx` method that inserts a `<w:br w:type="page"/>` element.

---

## 6. `mail_sender.py`

Handles all email sending logic using Windows COM automation to control the locally installed Microsoft Outlook application.

---

### Module-level import guard — Lines 22–28

```python
try:
    import win32com.client
except ImportError:
    raise ImportError("pywin32 is required...")
```

**Purpose:**
The `try/except` around the import means the module gives a clear, human-readable error message if `pywin32` is not installed, instead of a confusing `ModuleNotFoundError` traceback. This is called a "graceful import failure."

---

### `send_report(html_body, attachment_path, to_recipients, cc_recipients, subject)` — Line 37

```python
def send_report(
    html_body:       str,
    attachment_path: str,
    to_recipients:   list[str],
    cc_recipients:   list[str] | None = None,
    subject:         str = DEFAULT_SUBJECT,
) -> None:
```

**Purpose:**
The sole public function in `mail_sender.py`. Validates inputs, connects to Outlook via COM, composes the email, attaches the DOCX, and sends it.

**Parameters explained:**

| Parameter | Type | What it is |
|---|---|---|
| `html_body` | `str` | The full HTML string from `html_generator.py` |
| `attachment_path` | `str` | Path to the `.docx` file from `doc_generator.py` |
| `to_recipients` | `list[str]` | List of TO email addresses |
| `cc_recipients` | `list[str] \| None` | List of CC addresses, or `None` for no CC |
| `subject` | `str` | Email subject, defaults to `DEFAULT_SUBJECT` |

**How it works step by step:**

**Step 1 — Input validation:**
```python
if not to_recipients:
    raise ValueError(...)
if not os.path.isfile(abs_attachment):
    raise FileNotFoundError(...)
```
Validates before touching Outlook so the error message is specific and actionable.

**Step 2 — Connect to Outlook (with fallback):**
```python
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
except:
    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
```
`Dispatch` is the standard COM connection method. If it fails (e.g. the COM type cache is corrupted), `EnsureDispatch` is tried — it forces a rebuild of the local COM cache, which often fixes the connection.

**Step 3 — Create and configure mail item:**
```python
mail = outlook.CreateItem(0)   # 0 = olMailItem constant
mail.To      = "; ".join(to_recipients)
mail.CC      = "; ".join(cc_list)
mail.Subject = subject
mail.HTMLBody = html_body
```
`CreateItem(0)` creates a new email. Recipients are joined with `"; "` — Outlook parses semicolon-separated addresses as multiple recipients. `HTMLBody` (not `Body`) is used so Outlook renders the HTML instead of displaying it as raw text.

**Step 4 — Attach the DOCX:**
```python
mail.Attachments.Add(abs_attachment)
```
`abs_attachment` must be an absolute path — Outlook's COM interface does not resolve relative paths correctly. That's why `os.path.abspath()` is called earlier.

**Step 5 — Send:**
```python
mail.Send()
```
Sends immediately. The email appears in Outlook's **Sent Items** folder. Outlook must be installed and have a configured profile — it does not need to be open/visible.

**Raises:**
- `ValueError` — empty recipients list
- `FileNotFoundError` — DOCX file not found
- `RuntimeError` — any Outlook COM failure, with a human-readable message suggesting fixes

---

## Quick Function Index

| Function | File | Lines | Called by |
|---|---|---|---|
| `main()` | main.py | 39–100 | Python runtime (`__main__`) |
| `normalize_status()` | utils.py | 57–65 | `json_loader._validate_record()` |
| `resolve_path()` | utils.py | 68–77 | `utils.image_exists()` |
| `image_exists()` | utils.py | 80–89 | `doc_generator._add_activity_section()` |
| `load_activities()` | json_loader.py | 16–48 | `main.main()` |
| `_validate_record()` | json_loader.py | 51–83 | `json_loader.load_activities()` |
| `generate_html_body()` | html_generator.py | 37–81 | `main.main()` |
| `_build_table()` | html_generator.py | 84–127 | `html_generator.generate_html_body()` |
| `_escape()` | html_generator.py | 130–138 | `html_generator._build_table()` |
| `generate_docx()` | doc_generator.py | 22–46 | `main.main()` |
| `_set_page_margins()` | doc_generator.py | 49–55 | `doc_generator.generate_docx()` |
| `_add_document_title()` | doc_generator.py | 57–68 | `doc_generator.generate_docx()` |
| `_add_summary_table()` | doc_generator.py | 70–89 | `doc_generator.generate_docx()` |
| `_add_activity_section()` | doc_generator.py | 91–115 | `doc_generator.generate_docx()` (in loop) |
| `_add_horizontal_line()` | doc_generator.py | 117–130 | `doc_generator._add_activity_section()` |
| `_add_page_break()` | doc_generator.py | 131–135 | `doc_generator.generate_docx()` |
| `send_report()` | mail_sender.py | 37–113 | `main.main()` |
