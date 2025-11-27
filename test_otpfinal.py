import time
import random
import string
import re

import pytest
from pywinauto import Desktop, keyboard
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


REPORT = []


def log_step(desc, status="PASS"):
    REPORT.append((desc, status))
    print(f"{desc}: {status}")


# -------------------------------------------------------------
#  ALERT HANDLER
# -------------------------------------------------------------
def accept_alert_if_present(driver, timeout=5):
    if not driver:
        return
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        log_step("Browser alert detected.")
        alert.accept()
        log_step("Browser alert accepted.")
    except Exception:
        log_step("No browser alert present, continuing normally", "INFO")


# -------------------------------------------------------------
#  REALISTIC NAME GENERATOR
# -------------------------------------------------------------
FIRST_NAMES = [
    "John", "Alice", "David", "Sophia", "Liam", "Emma", "Noah", "Ava",
    "Ethan", "Mia", "Oliver", "Isabella", "James", "Charlotte", "Amelia",
    "Benjamin", "Harper", "Lucas", "Ella", "Henry",
]

LAST_NAMES = [
    "Smith", "Johnson", "Brown", "Williams", "Taylor", "Miller", "Wilson",
    "Davis", "Anderson", "Thomas", "Jackson", "White", "Harris", "Martin",
    "Garcia", "Clark", "Rodriguez", "Lewis", "Lee", "Walker",
]


def get_random_real_name():
    first = random.choice(FIRST_NAMES)
    last = random.choice(LAST_NAMES)
    return first, last


# -------------------------------------------------------------
#  RANDOM MAILBOX GENERATOR
# -------------------------------------------------------------
def generate_random_mailbox():
    prefix = ''.join(random.choices(string.ascii_lowercase, k=4))
    return prefix + "test"   # e.g. abcdtest


# -------------------------------------------------------------
#  BUILD EMAIL
# -------------------------------------------------------------
def build_email(first_name, last_name):
    # generate a random suffix
    suffix = generate_random_mailbox()   # e.g., abcdtest

    # full mailbox local-part used by mailsac
    mailbox = f"{first_name.lower()}.{last_name.lower()}.{suffix}"

    # full email address
    email = f"{mailbox}@mailsac.com"
    return email, mailbox


# -------------------------------------------------------------
#  HP SMART LAUNCH & ACCOUNT CREATION
# -------------------------------------------------------------
def launch_hp_smart():
    try:
        keyboard.send_keys("{VK_LWIN}HP Smart{ENTER}")
        log_step("Sent keys to launch HP Smart app.")

        desktop = Desktop(backend="uia")
        main_win = desktop.window(title_re=".*HP Smart.*")
        main_win.wait('exists visible enabled ready', timeout=30)
        main_win.set_focus()
        log_step("Focused HP Smart main window.")

        manage_account_btn = main_win.child_window(
            title="Manage HP Account",
            auto_id="HpcSignedOutIcon",
            control_type="Button",
        )
        manage_account_btn.wait('visible enabled ready', timeout=15)
        manage_account_btn.click_input()
        log_step("Clicked Manage HP Account button.")

        create_account_btn = main_win.child_window(
            auto_id="HpcSignOutFlyout_CreateBtn",
            control_type="Button",
        )
        create_account_btn.wait('visible enabled ready', timeout=15)
        create_account_btn.click_input()
        log_step("Clicked Create Account button.")

        return desktop

    except Exception as e:
        log_step(f"Error launching HP Smart: {e}", "FAIL")
        return None


# -------------------------------------------------------------
#  FILL ACCOUNT FORM
# -------------------------------------------------------------
def fill_account_form(desktop, first_name, last_name, email_id):
    try:
        browser_win = desktop.window(title_re=".*HP account.*")
        browser_win.wait('exists visible enabled ready', timeout=30)
        browser_win.set_focus()
        log_step("Focused HP Account browser window.")

        browser_win.child_window(auto_id="firstName", control_type="Edit").type_keys(first_name)
        browser_win.child_window(auto_id="lastName", control_type="Edit").type_keys(last_name)
        browser_win.child_window(auto_id="email", control_type="Edit").type_keys(email_id)
        browser_win.child_window(auto_id="password", control_type="Edit").type_keys("SecurePassword123")

        create_btn = browser_win.child_window(auto_id="sign-up-submit", control_type="Button")
        create_btn.wait('visible enabled ready', timeout=10)
        create_btn.click_input()

        log_step(f"Filled account form with Name: {first_name} {last_name} | Email: {email_id}")
        time.sleep(6)

    except Exception as e:
        log_step(f"Error filling account form: {e}", "FAIL")


# -------------------------------------------------------------
#  FETCH OTP (SELENIUM)
# -------------------------------------------------------------
def fetch_otp_from_mailsac(mailbox_name, max_wait=30, poll_interval=3):
    driver = None
    try:
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)

        driver.get("https://mailsac.com")
        log_step("Opened Mailsac website.")

        mailbox_field = wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='mailbox']"))
        )
        mailbox_field.send_keys(mailbox_name)

        check_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Check the mail!']"))
        )
        check_btn.click()
        log_step("Opened Mailsac inbox.")

        start_time = time.time()
        otp = None

        while time.time() - start_time < max_wait:
            try:
                email_row = WebDriverWait(driver, poll_interval).until(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         "//table[contains(@class,'inbox-table')]/tbody/tr[contains(@class,'clickable')][1]")
                    )
                )
                email_row.click()
                log_step("Clicked on first email row.")
                break
            except Exception:
                # Refresh inbox and retry
                try:
                    driver.find_element(By.XPATH, "//button[normalize-space()='Check the mail!']").click()
                    log_step("Refreshed Mailsac inbox.", "INFO")
                except Exception as refresh_err:
                    log_step(f"Unable to refresh inbox: {refresh_err}", "FAIL")
                    break

        body_elem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#emailBody")))
        email_body = body_elem.text

        otp_match = re.search(r"\b(\d{6})\b", email_body)
        if otp_match:
            otp = otp_match.group(1)
            log_step(f"Extracted OTP: {otp}")
        else:
            log_step("OTP not found in email.", "FAIL")

        return otp, driver

    except Exception as e:
        log_step(f"Error fetching OTP: {e}", "FAIL")
        if driver:
            driver.quit()
        return None, None


# -------------------------------------------------------------
#  OTP ENTRY — PYWINAUTO
# -------------------------------------------------------------
def complete_web_verification_in_app(otp):
    try:
        desktop = Desktop(backend="uia")
        otp_window = desktop.window(title_re=".*HP account.*")
        otp_window.wait('exists visible enabled ready', timeout=20)
        otp_window.set_focus()
        log_step("Focused OTP input screen.")

        otp_box = otp_window.child_window(auto_id="code", control_type="Edit")
        otp_box.wait('visible enabled ready', timeout=10)

        pyperclip.copy(otp)
        time.sleep(1)
        otp_box.click_input()
        otp_box.type_keys("^v")  # Paste OTP
        log_step("OTP pasted successfully.")

        verify_btn = otp_window.child_window(auto_id="submit-code", control_type="Button")
        verify_btn.wait('visible enabled ready', timeout=10)
        verify_btn.click_input()
        log_step("Clicked Verify button.")
        time.sleep(4)

    except Exception as e:
        log_step(f"OTP verification failed: {e}", "FAIL")


# -------------------------------------------------------------
#  REPORT
# -------------------------------------------------------------
def generate_report():
    html = """<html><head><title>Automation Report</title></head><body>
<h2>HP Account Automation Report</h2><table border='1'>
<tr><th>Step</th><th>Status</th></tr>"""
    for desc, status in REPORT:
        html += f"<tr><td>{desc}</td><td>{status}</td></tr>"
    html += "</table></body></html>"

    with open("automation_report.html", "w") as f:
        f.write(html)

    print("Report generated: automation_report.html")


# -------------------------------------------------------------
#  MAIN FLOW
# -------------------------------------------------------------
def main():
    # 1. random name
    first_name, last_name = get_random_real_name()

    # 2. email + mailbox (both aligned)
    email_id, mailbox = build_email(first_name, last_name)
    log_step(f"Generated email: {email_id}")

    # 3. HP Smart – use email_id in form
    desktop = launch_hp_smart()
    if desktop:
        fill_account_form(desktop, first_name, last_name, email_id)
    else:
        log_step("Desktop handle is None, aborting flow.", "FAIL")
        generate_report()
        return

    # 4. Use the SAME mailbox for Mailsac
    otp, driver = fetch_otp_from_mailsac(mailbox)
    if otp:
        complete_web_verification_in_app(otp)
    else:
        log_step("OTP was not retrieved. Skipping verification.", "FAIL")

    accept_alert_if_present(driver)

    if driver:
        driver.close()
        driver.quit()

    generate_report()


# -------------------------------------------------------------
#  PYTEST ENTRY POINT
# -------------------------------------------------------------
if __name__ == "__main__":
    main()


def test_hp_account_automation():
    main()
    assert True
