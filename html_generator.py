"""
html_generator.py - Build the HTML email body (greeting + status table + footer).
"""

from utils import STATUS_COLORS_HTML, STATUS_BG_HTML, logger


# ---------------------------------------------------------------------------
# Inline CSS constants
# ---------------------------------------------------------------------------

_TABLE_STYLE = (
    "border-collapse:collapse;"
    "width:100%;"
    "max-width:650px;"
    "font-family:Arial,Helvetica,sans-serif;"
    "font-size:13px;"
)

_TH_STYLE = (
    "background-color:#003366;"
    "color:#ffffff;"
    "padding:10px 14px;"
    "text-align:left;"
    "letter-spacing:0.5px;"
)

_TD_STYLE_BASE = (
    "padding:9px 14px;"
    "border-bottom:1px solid #dee2e6;"
)

_ROW_EVEN_BG  = "#f8f9fa"
_ROW_ODD_BG   = "#ffffff"


def generate_html_body(activities: list[dict]) -> str:
    """
    Return the full HTML string to be placed in the Outlook email body.

    The output contains:
      - A brief greeting paragraph
      - A colour-coded status table
      - A closing paragraph
    """
    logger.info("Generating HTML email body (%d rows)...", len(activities))

    table_html = _build_table(activities)

    html = f"""
<html>
<body style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#212529;margin:0;padding:20px;">

  <p>Dear Team,</p>

  <p>
    Please find below the <strong>Weekly Sanity Check Report</strong> summarising
    the status of all monitored activities.  A detailed report with supporting
    screenshots is attached to this email.
  </p>

  {table_html}

  <br/>
  <p>
    Please review the attached document for full details and screenshots.
    Kindly address any items marked as <span style="color:{STATUS_COLORS_HTML['Warning']};font-weight:bold;">Warning</span>
    or <span style="color:{STATUS_COLORS_HTML['Failed']};font-weight:bold;">Failed</span> at the earliest.
  </p>

  <p>
    Regards,<br/>
    <strong>Sanity Check Automation</strong>
  </p>

</body>
</html>
""".strip()

    logger.info("HTML body generated successfully.")
    return html


def _build_table(activities: list[dict]) -> str:
    """Build and return the HTML <table> string."""

    header = f"""
<table style="{_TABLE_STYLE}">
  <thead>
    <tr>
      <th style="{_TH_STYLE}width:50px;">S.No</th>
      <th style="{_TH_STYLE}">Activity</th>
      <th style="{_TH_STYLE}width:100px;">Status</th>
    </tr>
  </thead>
  <tbody>""".strip()

    rows = []
    for idx, item in enumerate(activities):
        row_bg  = _ROW_EVEN_BG if idx % 2 == 0 else _ROW_ODD_BG
        status  = item.get("status", "N/A")
        fg      = STATUS_COLORS_HTML.get(status, "#333333")
        badge_bg = STATUS_BG_HTML.get(status, "#e9ecef")

        sno_td = (
            f'<td style="{_TD_STYLE_BASE}background-color:{row_bg};'
            f'text-align:center;color:#555555;">'
            f'{item["sno"]}</td>'
        )
        act_td = (
            f'<td style="{_TD_STYLE_BASE}background-color:{row_bg};">'
            f'{_escape(str(item["activity"]))}</td>'
        )
        status_td = (
            f'<td style="{_TD_STYLE_BASE}background-color:{row_bg};">'
            f'<span style="'
            f'display:inline-block;padding:3px 10px;border-radius:4px;'
            f'background-color:{badge_bg};color:{fg};font-weight:bold;'
            f'font-size:12px;">'
            f'{_escape(status)}</span></td>'
        )

        rows.append(f"    <tr>{sno_td}{act_td}{status_td}</tr>")

    footer = "  </tbody>\n</table>"

    return header + "\n" + "\n".join(rows) + "\n" + footer


def _escape(text: str) -> str:
    """Minimal HTML escaping for table cell content."""
    return (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
