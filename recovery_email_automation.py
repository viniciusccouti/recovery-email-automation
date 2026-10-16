"""
Sets the recovery email (contact@wolfscale.com) for all users across each
pending domain in "Mailbox Recovery Project.xlsx", logging in as the admin
of each workspace in a visible browser (Playwright), with an isolated context
per domain (equivalent to an incognito tab).

Usage:
    python recovery_email_automation.py           # processes 20 domains (default)
    python recovery_email_automation.py 5          # processes only 5 (good for first test)

Progress is saved to the spreadsheet (done column) after each domain,
so running again continues from where it left off.
"""

import sys
import time
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

XLSX_PATH = Path(__file__).resolve().parent.parent / "Mailbox Recovery Project.xlsx"
LOCK_PATH = XLSX_PATH.parent / f"~$Mailbox Recovery Project.xlsx"
FAILURES_LOG = Path(__file__).resolve().parent / "failures.log"
SHEET_NAME = "mailboxes_export"
RECOVERY_EMAIL = "contact@wolfscale.com"
DEFAULT_BATCH_SIZE = 20

COL_CUSTOMER = 1
COL_EMAIL = 2
COL_PASSWORD = 3
COL_BACKUP = 4
COL_DOMAIN = 5
COL_DONE = 6
COL_NOTES = 7

PASSWORD_LOOKUP_SHEET = "Sheet1"
PASSWORD_LOOKUP_EMAIL_COL = 1  # A
PASSWORD_LOOKUP_PASSWORD_COL = 2  # B


def build_password_lookup(wb):
    # Column C of the main sheet is a VLOOKUP formula to Sheet1, but openpyxl
    # does not recalculate formulas (cached values are lost when saving). The
    # true source is always Sheet1.
    ws2 = wb[PASSWORD_LOOKUP_SHEET]
    lookup = {}
    for row in ws2.iter_rows(values_only=True):
        email = row[PASSWORD_LOOKUP_EMAIL_COL - 1]
        password = row[PASSWORD_LOOKUP_PASSWORD_COL - 1]
        if email:
            lookup[email] = password
    return lookup


def is_pending(done):
    # The done column can be a native Excel boolean (False) or the string "FALSE"
    # (after the spreadsheet was converted to English text), so accept both.
    if done is False:
        return True
    if isinstance(done, str) and done.strip().upper() == "FALSE":
        return True
    return False


def load_target_domains(wb, password_lookup):
    ws = wb[SHEET_NAME]
    rows = []
    for row_idx in range(2, ws.max_row + 1):
        email = ws.cell(row_idx, COL_EMAIL).value
        domain = ws.cell(row_idx, COL_DOMAIN).value
        done = ws.cell(row_idx, COL_DONE).value
        notes = ws.cell(row_idx, COL_NOTES).value
        if is_pending(done) and notes is None:
            password = password_lookup.get(email)
            rows.append({"row": row_idx, "email": email, "password": password, "domain": domain})

    by_domain = {}
    for r in rows:
        by_domain.setdefault(r["domain"], []).append(r)

    domains = []
    for domain, users in by_domain.items():
        no_dot = [u for u in users if "." not in u["email"].split("@")[0]]
        admin = None
        for u in no_dot:
            local = u["email"].split("@")[0]
            is_prefixed_by_another = any(
                o["email"] != u["email"] and local.startswith(o["email"].split("@")[0])
                for o in no_dot
            )
            if not is_prefixed_by_another:
                admin = u
                break
        if admin is None:
            # Should not happen: domains without a clear admin were already
            # marked "off" or "find who is the admin" in the spreadsheet.
            continue
        if not admin["password"]:
            print(f"  Warning: no password found in Sheet1 for {admin['email']}, marking {domain} and skipping")
            for u in users:
                ws.cell(u["row"], COL_NOTES).value = "password not found"
            continue
        domains.append({"domain": domain, "admin": admin, "users": users})
    return domains


def wait_for_manual_checkpoint(page, reason):
    # NOT ALL ACCOUNTS will hit this. Most log in directly without pausing.
    # This only pauses when GOOGLE specifically asks for something extra
    # (verification code, captcha, confirm recovery, etc) for THAT account.
    # As soon as you resolve it in the browser window and press ENTER here,
    # the script CONTINUES exactly from where it paused (no restart, no skip,
    # no repeat). The input() below blocks until you press ENTER.
    try:
        page.bring_to_front()
    except Exception:
        pass
    print(f"\n>>> PAUSED: {reason}")
    print(">>> Only happens on accounts where Google requests extra verification.")
    print(">>> NO RUSH and NO TIME LIMIT. Can wait for the code to arrive")
    print(">>> in email, take as long as needed. DON'T press Ctrl+C, it interrupts")
    print(">>> the script mid-run (doesn't cancel just this pause). Only press ENTER when done.")
    input(">>> Resolve it in the browser window and press ENTER here to continue from where it paused...\n")


def looks_like_verification_prompt(page):
    indicators = [
        "2-Step Verification",
        "Verify it's you",
        "Confirm your recovery",
        "unusual activity",
    ]
    try:
        content = page.content()
    except Exception:
        return False
    return any(text.lower() in content.lower() for text in indicators)


def login_as_admin(context, admin):
    page = context.new_page()
    page.goto(
        "https://accounts.google.com/ServiceLogin?service=accountsettings&continue=https://admin.google.com/&hl=en"
    )
    page.wait_for_selector("#identifierId", timeout=30000)
    page.fill("#identifierId", admin["email"])
    page.click("#identifierNext button")
    page.wait_for_timeout(2000)

    if looks_like_verification_prompt(page):
        wait_for_manual_checkpoint(page, "Unexpected screen after entering email (may be captcha/verification).")

    page.wait_for_selector('input[type="password"]', timeout=30000)
    page.fill('input[type="password"]', admin["password"])
    page.click("#passwordNext button")
    print("  Checking if Google will ask for extra confirmation, wait a few seconds (normal)...")
    page.wait_for_timeout(2500)

    if looks_like_verification_prompt(page):
        wait_for_manual_checkpoint(page, "2FA / additional verification detected.")

    print("  Loading admin console, can take up to 20s (normal, not a freeze)...")
    try:
        page.wait_for_url("**/admin.google.com/**", timeout=20000)
    except Exception:
        wait_for_manual_checkpoint(
            page, "Still not at admin.google.com. Confirm manually and press ENTER when there."
        )

    # After login, Google does some internal redirects (rapt=..., pli=1, etc)
    # before stabilizing. Give it time to finish before navigating again,
    # otherwise one navigation will interfere with another.
    try:
        page.wait_for_load_state("load", timeout=15000)
    except Exception:
        pass
    page.wait_for_timeout(3000)

    return page


def goto_with_retry(page, url, attempts=3):
    last_error = None
    for _ in range(attempts):
        try:
            page.goto(url)
            return
        except Exception as e:
            last_error = e
            page.wait_for_timeout(2000)
    raise last_error


def goto_users_list(page):
    goto_with_retry(page, "https://admin.google.com/ac/users?hl=en")
    # "Users" text appears many times on the screen (menu, title, etc),
    # which confuses the wait. Confirm by URL instead, which is unique.
    page.wait_for_url("**/ac/users**", timeout=20000)
    page.wait_for_timeout(1500)


def set_recovery_email_for_user(page, user_email):
    # The Users list already shows everyone in the domain directly on screen
    # (small domains, few users), so click directly on the row instead of
    # trying to search (the search box at top is "Ask me anything" AI, not
    # a user list filter).
    #
    # Multiple users can have the SAME name (e.g., "Joe Davy" in 3 variants
    # of banzaibrand.com), so can't click by name. Instead, find the TABLE ROW
    # containing the exact email, then click the real link inside it
    # (<a href="./ac/users/{id}">).
    row = page.locator("tr", has_text=user_email)
    row.locator('a[href*="/ac/users/"]').first.click()
    page.wait_for_timeout(2000)

    # "Security" also exists in the sidebar menu (org-wide settings), so
    # get_by_text would get both. Use role="tab" to get only the tab
    # within the user's panel.
    try:
        page.get_by_role("tab", name="Security", exact=True).click()
    except Exception:
        page.get_by_text("Security", exact=True).last.click()
    page.wait_for_timeout(1500)

    # "Recovery information" shows "Add a recovery email" as a link (not a
    # ready field), it becomes a text field only after clicking it.
    page.get_by_text("Add a recovery email", exact=True).first.click()
    page.wait_for_timeout(1000)

    # The modal has 2 fields (Email and Phone), must target Email specifically,
    # otherwise text goes to the wrong field (Phone).
    recovery_field = page.get_by_role("textbox", name="Email", exact=False).first
    recovery_field.click()
    recovery_field.fill(RECOVERY_EMAIL)
    page.wait_for_timeout(500)
    page.get_by_role("button", name="Save").first.click()
    page.wait_for_timeout(1500)

    # Return to user list to process next user
    goto_users_list(page)


DEBUG_DIR = Path(__file__).resolve().parent / "debug_screenshots"


def process_domain(browser, entry, wb):
    domain = entry["domain"]
    admin = entry["admin"]
    print(f"\n--- {domain} (admin: {admin['email']}, {len(entry['users'])} user(s)) ---")

    context = browser.new_context()
    page = None
    try:
        page = login_as_admin(context, admin)
        goto_users_list(page)
        for u in entry["users"]:
            print(f"  recovery email -> {u['email']}")
            set_recovery_email_for_user(page, u["email"])
        for u in entry["users"]:
            wb[SHEET_NAME].cell(u["row"], COL_DONE).value = "TRUE"
        return True, None
    except Exception as e:
        print(f"  ERROR: {e}")
        # Notes intentionally NOT marked here: domain stays eligible to retry
        # in the next run without manual cleanup. Failures logged to file below.
        with open(FAILURES_LOG, "a", encoding="utf-8") as f:
            f.write(f"{domain} | admin={admin['email']} | {e}\n")
        if page is not None:
            DEBUG_DIR.mkdir(exist_ok=True)
            safe_name = domain.replace(".", "_")
            try:
                page.screenshot(path=str(DEBUG_DIR / f"{safe_name}.png"), full_page=True)
                (DEBUG_DIR / f"{safe_name}.html").write_text(page.content(), encoding="utf-8")
                print(f"  screenshot saved to debug_screenshots/{safe_name}.png")
            except Exception as screenshot_error:
                print(f"  (couldn't save screenshot: {screenshot_error})")
        return False, str(e)
    finally:
        context.close()


def main():
    if LOCK_PATH.exists():
        print(f"The file '{XLSX_PATH.name}' is open in Excel. Close it and run again.")
        sys.exit(1)

    batch_size = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_BATCH_SIZE

    wb = openpyxl.load_workbook(XLSX_PATH, data_only=False)
    password_lookup = build_password_lookup(wb)
    domains = load_target_domains(wb, password_lookup)
    wb.save(XLSX_PATH)  # persist "password not found" marks now
    batch = domains[:batch_size]

    print(f"Total pending domains: {len(domains)}")
    print(f"Processing {len(batch)} now (batch_size={batch_size})\n")

    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        for entry in batch:
            ok, err = process_domain(browser, entry, wb)
            results.append((entry["domain"], ok, err))
            wb.save(XLSX_PATH)
        browser.close()

    success = sum(1 for _, ok, _ in results if ok)
    failed = len(results) - success

    print("\n=== Batch Summary ===")
    for domain, ok, err in results:
        status = "OK" if ok else f"FAILED: {err}"
        print(f"  {domain}: {status}")

    print(f"\n{success} completed successfully, {failed} failed in this batch.")
    print(f"Total pending domains still (including failures, retryable): {len(domains) - success}")


if __name__ == "__main__":
    main()
