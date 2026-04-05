"""
main.py - Entry point for the Weekly Sanity Check Report automation.

Usage:
    python main.py

Configure the TO / CC recipients and file paths in the CONFIG block below.
"""

import os
import sys

# ---------------------------------------------------------------------------
# CONFIG  –  edit these values before running
# ---------------------------------------------------------------------------

# Path to the input JSON file (relative to this script or absolute)
JSON_FILE = "data.json"

# Output path for the generated DOCX report
DOCX_OUTPUT = "Weekly_Sanity_Report.docx"

# Email recipients
TO_RECIPIENTS = [
    "charantej2714@outlook.com",
]

CC_RECIPIENTS = [
    "charan.mamidi@capgemini.com",
]

# Optional: override the email subject
EMAIL_SUBJECT = "Weekly Sanity Check Report"

# ---------------------------------------------------------------------------
# END CONFIG
# ---------------------------------------------------------------------------


def main():
    # Resolve paths relative to the directory of this script
    base_dir   = os.path.dirname(os.path.abspath(__file__))
    json_path  = os.path.join(base_dir, JSON_FILE)
    docx_path  = os.path.join(base_dir, DOCX_OUTPUT)

    print("=" * 60)
    print("  Weekly Sanity Check Report – Automation Started")
    print("=" * 60)

    # ------------------------------------------------------------------
    # Step 1: Load and validate JSON data
    # ------------------------------------------------------------------
    print("\n[1/4] Loading activity data from JSON...")
    try:
        from json_loader import load_activities
        activities = load_activities(json_path)
    except (FileNotFoundError, ValueError) as exc:
        print(f"\n[FATAL] {exc}")
        sys.exit(1)

    if not activities:
        print("[FATAL] No valid activities found in the JSON file.  Aborting.")
        sys.exit(1)

    # ------------------------------------------------------------------
    # Step 2: Generate HTML email body
    # ------------------------------------------------------------------
    print("\n[2/4] Generating HTML email body...")
    try:
        from html_generator import generate_html_body
        html_body = generate_html_body(activities)
    except Exception as exc:
        print(f"\n[FATAL] HTML generation failed: {exc}")
        sys.exit(1)

    # ------------------------------------------------------------------
    # Step 3: Generate DOCX report
    # ------------------------------------------------------------------
    print("\n[3/4] Generating DOCX report...")
    try:
        from doc_generator import generate_docx
        docx_file = generate_docx(activities, docx_path)
        print(f"       Report saved → {docx_file}")
    except RuntimeError as exc:
        print(f"\n[FATAL] DOCX generation failed: {exc}")
        sys.exit(1)

    # ------------------------------------------------------------------
    # Step 4: Send email via Outlook COM automation
    # ------------------------------------------------------------------
    print("\n[4/4] Sending email via local Outlook application...")
    try:
        from mail_sender import send_report
        send_report(
            html_body       = html_body,
            attachment_path = docx_file,
            to_recipients   = TO_RECIPIENTS,
            cc_recipients   = CC_RECIPIENTS,
            subject         = EMAIL_SUBJECT,
        )
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"\n[FATAL] Email sending failed: {exc}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  All steps completed successfully!")
    print("=" * 60)


if __name__ == "__main__":
    main()
