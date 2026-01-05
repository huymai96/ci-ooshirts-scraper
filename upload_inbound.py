"""
Upload inbound.csv to Promos Ink Supply Chain API
--------------------------------------------------
Uploads the inbound tracking data to the cloud dashboard.
Includes headless operation, logging, and email alerts.

Requires:
  pip install requests
"""

import os
import time
import logging
import traceback
import smtplib
import requests
from pathlib import Path
from logging.handlers import RotatingFileHandler
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --------- Paths ---------
SCRIPT_DIR = Path(__file__).resolve().parent
LOG_PATH = SCRIPT_DIR / "upload_inbound.log"
INBOUND_CSV = SCRIPT_DIR / "inbound.csv"

# --------- API Settings ---------
API_URL = "https://supplychain.promosinkwall-e.com/api/manifests"
API_KEY = "promos-ink-2024"

# --------- Logging ---------
logger = logging.getLogger("upload_inbound")
logger.setLevel(logging.INFO)
if not logger.handlers:
    fh = RotatingFileHandler(str(LOG_PATH), maxBytes=1_000_000, backupCount=2, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
    logger.addHandler(fh)
    sh = logging.StreamHandler()
    sh.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(sh)

# --------- Email Settings ---------
# Import from email_config.py if available
try:
    from email_config import ALERT_EMAIL, SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASS
except ImportError:
    ALERT_EMAIL = "alerts@example.com"
    SMTP_SERVER = "smtp.example.com"
    SMTP_PORT = 587
    SMTP_USER = ""  # Configure in email_config.py
    SMTP_PASS = ""

def send_error_email(subject, error_message):
    """Send email notification when upload fails."""
    if not SMTP_USER or not SMTP_PASS:
        logger.warning("[Email] SMTP credentials not configured - skipping email alert")
        return False
    
    try:
        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = ALERT_EMAIL
        msg['Subject'] = f"[Inbound Upload Alert] {subject}"
        
        body = f"""
Inbound CSV Upload encountered an error:

{error_message}

----------------------------------------
Time: {time.strftime('%Y-%m-%d %H:%M:%S')}
Computer: {os.environ.get('COMPUTERNAME', 'Unknown')}
File: {INBOUND_CSV}
Log file: {LOG_PATH}
----------------------------------------

This is an automated message from the CI/Ooshirts Scraper.
"""
        msg.attach(MIMEText(body, 'plain'))
        
        if SMTP_PORT == 465:
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                server.login(SMTP_USER, SMTP_PASS)
                server.send_message(msg)
        else:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.send_message(msg)
        
        logger.info(f"[Email] Alert sent to {ALERT_EMAIL}")
        return True
    except Exception as e:
        logger.error(f"[Email] Failed to send alert: {e}")
        return False

def upload_inbound_csv():
    """Upload inbound.csv to the Supply Chain API."""
    logger.info("=" * 50)
    logger.info("INBOUND CSV UPLOAD STARTED")
    logger.info(f"Time: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 50)
    
    # Check if file exists
    if not INBOUND_CSV.exists():
        error_msg = f"File not found: {INBOUND_CSV}"
        logger.error(error_msg)
        send_error_email("File Not Found", error_msg)
        return False
    
    # Get file info
    file_size = INBOUND_CSV.stat().st_size
    file_mtime = time.ctime(INBOUND_CSV.stat().st_mtime)
    logger.info(f"File: {INBOUND_CSV}")
    logger.info(f"Size: {file_size:,} bytes")
    logger.info(f"Modified: {file_mtime}")
    
    # Upload to API
    try:
        logger.info(f"Uploading to: {API_URL}")
        
        with open(INBOUND_CSV, 'rb') as f:
            response = requests.post(
                API_URL,
                files={'file': (INBOUND_CSV.name, f, 'text/csv')},
                data={'type': 'inbound'},
                headers={'x-api-key': API_KEY},
                timeout=120  # 2 minute timeout for large files
            )
        
        if response.status_code == 200:
            logger.info(f"[OK] Upload successful! Status: {response.status_code}")
            try:
                result = response.json()
                logger.info(f"API Response: {result}")
            except:
                logger.info(f"API Response: {response.text[:500]}")
            return True
        else:
            error_msg = f"Upload failed with status {response.status_code}: {response.text[:500]}"
            logger.error(error_msg)
            send_error_email(f"Upload Failed (HTTP {response.status_code})", error_msg)
            return False
            
    except requests.exceptions.Timeout:
        error_msg = "Upload timed out after 120 seconds"
        logger.error(error_msg)
        send_error_email("Upload Timeout", error_msg)
        return False
    except requests.exceptions.ConnectionError as e:
        error_msg = f"Connection error: {e}"
        logger.error(error_msg)
        send_error_email("Connection Error", error_msg)
        return False
    except Exception as e:
        error_msg = f"Unexpected error: {traceback.format_exc()}"
        logger.error(error_msg)
        send_error_email("Upload Error", error_msg)
        return False

def main():
    success = upload_inbound_csv()
    if success:
        logger.info("[OK] Inbound CSV upload completed successfully")
        print("\n[OK] Upload completed successfully!")
    else:
        logger.error("[FAILED] Inbound CSV upload failed")
        print("\n[FAILED] Upload failed. Check upload_inbound.log for details.")
    return success

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = traceback.format_exc()
        logger.exception("Fatal error")
        send_error_email("Script Crashed", error_msg)
        print("\n[!] An error occurred. See upload_inbound.log for details.")
