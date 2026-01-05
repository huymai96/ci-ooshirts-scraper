# CI & Ooshirts Order Scraper → Supply Chain API

Automated scraper that extracts orders from **CustomInk** and **Ooshirts** portals, then uploads them to the **Promos Ink Supply Chain API** for the package confirmation dashboard.

## 🏗️ Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           SCRAPER COMPUTER                                   │
│  (Windows PC with Chrome + Python)                                          │
│                                                                              │
│  ┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐         │
│  │  ooshirts_      │    │  CI_order_      │    │  upload_        │         │
│  │  order_scraper  │    │  scraper.py     │    │  inbound.py     │         │
│  │  .py            │    │                 │    │                 │         │
│  └────────┬────────┘    └────────┬────────┘    └────────┬────────┘         │
│           │                      │                      │                   │
│           ▼                      ▼                      ▼                   │
│  ┌─────────────────────────────────────────────────────────────────┐       │
│  │                  customink_orders.xlsx (local)                   │       │
│  └──────────────────────────────┬──────────────────────────────────┘       │
│                                 │                                           │
│                                 ▼                                           │
│  ┌─────────────────────────────────────────────────────────────────┐       │
│  │              HTTPS POST to Supply Chain API                      │       │
│  │       https://supplychain.promosinkwall-e.com/api/manifests     │       │
│  │              (type: "customink" or "inbound")                    │       │
│  └──────────────────────────────┬──────────────────────────────────┘       │
└─────────────────────────────────┼───────────────────────────────────────────┘
                                  │
                                  ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                        VERCEL (Cloud)                                        │
│                                                                              │
│  ┌─────────────────────────────────────────────────────────────────┐       │
│  │                   Vercel Blob Storage                            │       │
│  │   manifests/customink-*.xlsx    manifests/inbound-*.csv         │       │
│  └──────────────────────────────┬──────────────────────────────────┘       │
│                                 │                                           │
│                                 ▼                                           │
│  ┌─────────────────────────────────────────────────────────────────┐       │
│  │   /api/cron/rebuild-index (hourly) → tracking-index.json        │       │
│  └──────────────────────────────┬──────────────────────────────────┘       │
│                                 │                                           │
│                                 ▼                                           │
│  ┌─────────────────────────────────────────────────────────────────┐       │
│  │        /api/label-lookup?tracking=XXX                           │       │
│  │        (Used by Label Print GUI at receiving stations)           │       │
│  └─────────────────────────────────────────────────────────────────┘       │
└─────────────────────────────────────────────────────────────────────────────┘
```

## 📁 File Structure

```
ci-ooshirts-scraper/
├── run_scrapers.bat           # Main batch file (runs all 3 scripts)
├── CI_order_scraper.py        # CustomInk scraper + cloud upload
├── ooshirts_order_scraper.py  # Ooshirts scraper (local only)
├── upload_inbound.py          # Uploads inbound.csv to cloud
├── email_config.example.py    # Template for credentials
├── test_email.py              # Test email sending
├── Run Receiving Tool Scrapers.xml  # Windows Task Scheduler export
├── setup_task.ps1             # PowerShell script to create scheduled task
├── update_task_hourly.bat     # Updates task to run hourly
└── README.md                  # This file
```

## 🔧 Prerequisites

### 1. Python 3.8+
```powershell
python --version
```

### 2. Required Python Packages
```powershell
pip install selenium openpyxl requests
```

### 3. Chrome Browser
Google Chrome must be installed. The scraper uses headless Chrome for automation.

## 📋 Setup Instructions

### Step 1: Clone Repository
```powershell
git clone https://github.com/huymai96/ci-ooshirts-scraper.git
cd ci-ooshirts-scraper
```

### Step 2: Install Dependencies
```powershell
pip install -r requirements.txt
```

### Step 3: Configure Email Alerts
```powershell
copy email_config.example.py email_config.py
# Edit email_config.py with your SMTP credentials
```

### Step 4: Create Scheduled Task
```powershell
.\setup_task.ps1
```

## ⚙️ How It Works

1. **Ooshirts Scraper** - Logs into 2 accounts, scrapes orders, saves locally
2. **CustomInk Scraper** - Scrapes CI portal, saves locally, **uploads to Supply Chain API**
3. **Inbound Upload** - **Uploads inbound.csv to Supply Chain API**

### API Details

**Endpoint:** `https://supplychain.promosinkwall-e.com/api/manifests`

**Headers:** `x-api-key: promos-ink-2024`

**Body:** `file` + `type` (customink or inbound)

## 🕐 Schedule

Runs **every hour** via Windows Task Scheduler (~25 min total runtime).

## 📧 Error Alerts

Email alerts sent on failures - configure in `email_config.py`.

## 🏷️ Related Systems

- **Package Confirmation App** - The cloud dashboard
- **Label Print GUI** - Receiving station software

---

**Maintained by:** Promos Ink Engineering Team
