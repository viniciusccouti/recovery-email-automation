# Recovery Email Automation

Automate setting recovery emails across multiple Google Workspace domains.

## What it does

- Reads a spreadsheet with domain and user data
- Logs into Google Admin Console for each domain
- Sets a recovery email address for each user
- Tracks progress in the spreadsheet
- Uses isolated browser contexts per domain
- Handles 2FA and manual verification when needed

## Tech Stack

- **Playwright** - Browser automation
- **openpyxl** - Excel file handling
- **Python 3.8+**

## Usage

```bash
# Process 20 domains (default)
python recovery_email_automation.py

# Process 5 domains (for testing)
python recovery_email_automation.py 5
```

## Installation

```bash
pip install playwright openpyxl
playwright install chromium
```

## Features

- **Isolated contexts** - Each domain runs in its own browser context
- **Progress persistence** - Results saved to Excel after each domain
- **Manual checkpoints** - Pauses for 2FA or verification
- **Error logging** - Failed domains logged for retry
- **Graceful degradation** - One failure doesn't stop the batch

## Workflow

1. Admin logs in via Google account chooser
2. Navigate to Users list in Google Admin
3. Set recovery email for each user
4. Repeat for next domain
5. Mark domain as complete

---

Built for automating Google Workspace bulk operations.
