"""
mail_sender.py - Send an email via local Outlook application (Windows only).

Uses Windows COM automation to send emails through your installed Outlook app.
No SMTP configuration, passwords, or authentication required!

Requirements:
  - Windows operating system
  - Microsoft Outlook installed and configured with your account
  - pywin32 library (install via: pip install pywin32)

Benefits over SMTP:
  - No authentication issues or app passwords needed
  - Bypasses "basic authentication disabled" restrictions
  - Emails appear in your Outlook Sent Items folder
  - Works with any Outlook account (personal or work)
"""

import os
from pathlib import Path

try:
    import win32com.client
except ImportError:
    raise ImportError(
        "pywin32 is required for Outlook automation.\n"
        "Install it with: pip install pywin32"
    )

from utils import logger

# ---------------------------------------------------------------------------
DEFAULT_SUBJECT = "Weekly Sanity Check Report"
# ---------------------------------------------------------------------------


def send_report(
    html_body:       str,
    attachment_path: str,
    to_recipients:   list[str],
    cc_recipients:   list[str] | None = None,
    subject:         str = DEFAULT_SUBJECT,
) -> None:
    """
    Compose and send an HTML email with a DOCX attachment via local Outlook.

    Parameters
    ----------
    html_body       : Full HTML string for the email body.
    attachment_path : Absolute path to the DOCX file to attach.
    to_recipients   : List of TO email addresses.
    cc_recipients   : List of CC email addresses (may be empty / None).
    subject         : Email subject line.

    Raises
    ------
    FileNotFoundError  – attachment file is missing.
    ValueError         – to_recipients is empty.
    RuntimeError       – Outlook COM automation fails.
    """

    # --- Input validation ---------------------------------------------------
    if not to_recipients:
        raise ValueError("to_recipients must contain at least one email address.")

    abs_attachment = os.path.abspath(attachment_path)
    if not os.path.isfile(abs_attachment):
        raise FileNotFoundError(f"Attachment not found: {abs_attachment}")

    cc_list = cc_recipients or []

    # --- Create email via Outlook COM ---------------------------------------
    logger.info("Creating email via Outlook application...")
    try:
        # Connect to Outlook (try multiple methods)
        outlook = None
        try:
            # Method 1: Standard Dispatch
            outlook = win32com.client.Dispatch("Outlook.Application")
        except:
            # Method 2: EnsureDispatch (forces cache rebuild)
            logger.info("Standard Dispatch failed, trying EnsureDispatch...")
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        
        mail = outlook.CreateItem(0)  # 0 = olMailItem

        # Set recipients
        mail.To = "; ".join(to_recipients)
        if cc_list:
            mail.CC = "; ".join(cc_list)

        # Set subject and body
        mail.Subject = subject
        mail.HTMLBody = html_body

        # Attach the DOCX file
        mail.Attachments.Add(abs_attachment)

        # Send the email
        logger.info("Sending email via Outlook...")
        mail.Send()

        logger.info("Email sent successfully to: %s", ", ".join(to_recipients))
        if cc_list:
            logger.info("CC: %s", ", ".join(cc_list))

    except Exception as exc:
        raise RuntimeError(
            f"Failed to send email via Outlook: {exc}\n"
            "  • Make sure Microsoft Outlook is installed and configured\n"
            "  • Try opening Outlook manually to verify it's working\n"
            "  • Check that your Outlook profile is set up correctly"
        ) from exc
