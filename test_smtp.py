"""Quick SMTP authentication test"""
import smtplib
from dotenv import load_dotenv
import os
from pathlib import Path

# Load .env
_env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=_env_path)

SMTP_SENDER = os.getenv("SMTP_SENDER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_HOST = "smtp-mail.outlook.com"
SMTP_PORT = 587

print(f"Testing SMTP authentication...")
print(f"Email: {SMTP_SENDER}")
print(f"Password: {'*' * len(SMTP_PASSWORD) if SMTP_PASSWORD else 'NOT SET'}")
print(f"Server: {SMTP_HOST}:{SMTP_PORT}")
print("-" * 60)

try:
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
        server.set_debuglevel(1)  # Show detailed SMTP conversation
        print("\n[1] Connecting...")
        server.ehlo()
        print("\n[2] Starting TLS...")
        server.starttls()
        print("\n[3] EHLO after TLS...")
        server.ehlo()
        print("\n[4] Logging in...")
        server.login(SMTP_SENDER, SMTP_PASSWORD)
        print("\n✓ SUCCESS! SMTP authentication worked.")
except smtplib.SMTPAuthenticationError as e:
    print(f"\n✗ AUTHENTICATION FAILED: {e}")
    print("\nPossible issues:")
    print("1. App password is incorrect or expired")
    print("2. Two-factor authentication not enabled on your Microsoft account")
    print("3. Account security settings blocking less secure apps")
except Exception as e:
    print(f"\n✗ ERROR: {e}")
