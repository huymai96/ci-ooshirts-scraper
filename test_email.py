"""
Test Email Configuration
Run this to verify email alerts are working.
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

# Import config
try:
    from email_config import ALERT_EMAIL, SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASS
    print("[OK] Config loaded from email_config.py")
except ImportError:
    print("[ERROR] Could not load email_config.py")
    print("Create email_config.py from email_config.example.py first.")
    exit(1)

print(f"\nEmail Settings:")
print(f"  Server: {SMTP_SERVER}:{SMTP_PORT}")
print(f"  From: {SMTP_USER}")
print(f"  To: {ALERT_EMAIL}")
print()

# Create test message
msg = MIMEMultipart()
msg['From'] = SMTP_USER
msg['To'] = ALERT_EMAIL
msg['Subject'] = "[TEST] Scraper Email Alert Test"

body = f"""
This is a TEST email from the CI/Ooshirts Scraper.

If you received this, email alerts are working correctly!

----------------------------------------
Computer: {os.environ.get('COMPUTERNAME', 'Unknown')}
----------------------------------------
"""
msg.attach(MIMEText(body, 'plain'))

print("Sending test email...")
try:
    if SMTP_PORT == 465:
        print(f"  Using SSL connection on port 465...")
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
    else:
        print(f"  Using TLS connection on port {SMTP_PORT}...")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
    
    print(f"\n[SUCCESS] Test email sent to {ALERT_EMAIL}")
    print("  Check your inbox!")
    
except Exception as e:
    print(f"\n[FAILED] {e}")
    print("\nTroubleshooting:")
    print("  - Verify username/password are correct")
    print("  - Check if the mail server allows SMTP access")

input("\nPress Enter to close...")
