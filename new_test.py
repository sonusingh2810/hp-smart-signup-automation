import time

import pytest
from pywinauto import Desktop, keyboard
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


REPORT = []


def log_step(desc, status="PASS"):
    REPORT.append((desc, status))
    print(f"{desc}: {status}")


# -------------------------------------------------------------
#  ALERT HANDLER (placeholder for future Selenium use)
# -------------------------------------------------------------
def accept_alert_if_present(driver, timeout=5):
    """Handle any browser alert if a Selenium driver is provided."""
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
#  HP SMART LAUNCH & SIGN-IN ENTRY
# -------------------------------------------------------------
def launch_hp_smart():
    """
    Launch HP Smart app and navigate to the HP Account sign-in flow.

    Currently uses Win+Search to launch HP Smart.
    """
    try:
        # Launch HP Smart via Windows search
        keyboard.send_keys("{VK_LWIN}HP Smart{ENTER}")
        log_step("Sent keys to launch HP Smart app.")

        desktop = Desktop(backend="uia")
        main_win = desktop.window(title_re=".*HP Smart.*")
        main_win.wait("exists visible enabled ready", timeout=30)
        main_win.set_focus()
        log_step("Focused HP Smart main window.")

        # Open Manage HP Account flyout
        manage_account_btn = main_win.child_window(
            title="Manage HP Account",
            auto_id="HpcSignedOutIcon",
            control_type="Button",
        )
        manage_account_btn.wait("visible enabled ready", timeout=15)
        manage_account_btn.click_input()
        log_step("Clicked Manage HP Account button.")

        # Try to click explicit "Sign in" button in the flyout (if present)
        try:
            sign_in_btn = main_win.child_window(
                auto_id="HpcSignOutFlyout_SignInBtn",
                control_type="Button",
            )
            sign_in_btn.wait("visible enabled ready", timeout=10)
            sign_in_btn.click_input()
            log_step("Clicked Sign In button.")
        except Exception:
            log_step(
                "Sign In button not found in flyout, assuming browser is already opened.",
                "INFO",
            )

        return desktop

    except Exception as e:
        log_step(f"Error launching HP Smart or navigating to sign-in: {e}", "FAIL")
        return None


# -------------------------------------------------------------
#  SIGN-IN FORM FILLING
# -------------------------------------------------------------
def sign_in_hp_account(desktop, email, password):
    """
    HP Account SIGN-IN flow:

      1) Focus HP account browser window
      2) Type email into the username/email textbox (already focused)
      3) Click 'Use password' button (NOT 'Sign in with mobile number')
      4) Type password on the next screen
      5) Click 'Sign in' button
    """
    try:
        # Attach to the HP account browser window (Chrome / Edge etc.)
        browser_win = Desktop(backend="uia").window(title_re=".*HP account.*")
        browser_win.wait("exists visible enabled ready", timeout=60)
        browser_win.set_focus()
        log_step("Focused HP Account sign-in browser window.")

        time.sleep(2)  # let the page finish loading

        # --- STEP 1: Type email / username ---
        keyboard.send_keys("^a{BACKSPACE}")  # clear any existing text
        keyboard.send_keys(email, with_spaces=True)
        log_step(f"Typed email/username: {email}")

        time.sleep(1)

        # --- STEP 2: Click 'Use password' button explicitly ---
        use_pwd_btn = browser_win.child_window(
            title="Use password",
            control_type="Button",
        )
        use_pwd_btn.wait("visible enabled ready", timeout=30)
        use_pwd_btn.click_input()
        log_step("Clicked 'Use password' button.")

        # Wait for the password screen to appear
        time.sleep(3)

        # --- STEP 3: Type password into the password field ---
        # From HTML: id="password" => auto_id="password"
        try:
            pwd_box = browser_win.child_window(
                auto_id="password",
                control_type="Edit",
            )
            pwd_box.wait("visible enabled ready", timeout=30)
        except Exception:
            edits = browser_win.descendants(control_type="Edit")
            log_step(
                f"Password field by auto_id not found, falling back to first Edit. Count={len(edits)}",
                "INFO",
            )
            if not edits:
                log_step("No Edit control found for password field.", "FAIL")
                # browser_win.print_control_identifiers()
                return
            pwd_box = edits[0]

        pwd_box.click_input()
        keyboard.send_keys("^a{BACKSPACE}")
        keyboard.send_keys(password, with_spaces=True)
        log_step("Typed password.")

        time.sleep(1)

        # --- STEP 4: Click final 'Sign in' button ---
        # HTML: id="sign-in" name="sign-in"
        try:
            sign_in_btn = browser_win.child_window(
                auto_id="sign-in",
                control_type="Button",
            )
            sign_in_btn.wait("visible enabled ready", timeout=30)
            sign_in_btn.click_input()
            log_step("Clicked final 'Sign in' button using auto_id.")
        except Exception:
            # Fallback: search all buttons containing 'Sign in' text
            buttons = browser_win.descendants(control_type="Button")
            target = None
            for b in buttons:
                text = b.window_text().strip()
                if "Sign in" in text:
                    target = b
                    break

            if not target:
                log_step("Could not locate 'Sign in' button.", "FAIL")
                # browser_win.print_control_identifiers()
                return

            target.wait("visible enabled ready", timeout=30)
            target.click_input()
            log_step("Clicked final 'Sign in' button using text fallback.")

        time.sleep(6)  # wait for login to complete

    except Exception as e:
        log_step(f"Error during HP Account sign-in: {e}", "FAIL")
        # For deeper debugging, you can temporarily enable:
        # try:
        #     browser_win.print_control_identifiers()
        # except Exception:
        #     pass


# -------------------------------------------------------------
#  CLICK SCAN BUTTON ON HP SMART MAIN WINDOW
# -------------------------------------------------------------
def click_scan_button(desktop):
    """
    Bring HP Smart main window to front and click the 'Scan' tile/button.
    """
    try:
        main_win = desktop.window(title_re=".*HP Smart.*")
        main_win.wait("exists visible enabled ready", timeout=30)
        main_win.set_focus()
        log_step("Refocused HP Smart main window before clicking 'Scan'.")

        # First try to locate by exact title
        try:
            scan_btn = main_win.child_window(
                title="Scan",
                control_type="Button",
            )
            scan_btn.wait("visible enabled ready", timeout=15)
        except Exception:
            # Fallback: search all buttons whose text contains 'Scan'
            buttons = main_win.descendants(control_type="Button")
            target = None
            for b in buttons:
                txt = b.window_text().strip()
                if "Scan" in txt:
                    target = b
                    break

            if not target:
                log_step("Could not find 'Scan' button on HP Smart main window.", "FAIL")
                # For debugging you can uncomment:
                # main_win.print_control_identifiers()
                return

            scan_btn = target

        scan_btn.click_input()
        log_step("Clicked 'Scan' button on HP Smart home screen.")

        time.sleep(5)  # wait for Scan screen to open

    except Exception as e:
        log_step(f"Error while clicking 'Scan' button: {e}", "FAIL")


# -------------------------------------------------------------
#  CLICK RETURN HOME BUTTON ON SCAN SCREEN
# -------------------------------------------------------------
def click_return_home_button(desktop):
    """
    On the Scan screen, click the 'Return Home' button
    when 'Scanning is Currently Unavailable' is shown.
    """
    try:
        main_win = desktop.window(title_re=".*HP Smart.*")
        main_win.wait("exists visible enabled ready", timeout=30)
        main_win.set_focus()
        log_step("Focused HP Smart Scan screen to click 'Return Home'.")

        # Try by exact title first
        try:
            return_btn = main_win.child_window(
                title="Return Home",
                control_type="Button",
            )
            return_btn.wait("visible enabled ready", timeout=15)
        except Exception:
            # Fallback: search all buttons containing 'Return Home'
            buttons = main_win.descendants(control_type="Button")
            target = None
            for b in buttons:
                txt = b.window_text().strip()
                if "Return Home" in txt:
                    target = b
                    break

            if not target:
                log_step("Could not find 'Return Home' button on Scan screen.", "FAIL")
                # For debugging:
                # main_win.print_control_identifiers()
                return

            return_btn = target

        return_btn.click_input()
        log_step("Clicked 'Return Home' button on Scan screen.")
        time.sleep(5)  # wait to navigate back to home

    except Exception as e:
        log_step(f"Error while clicking 'Return Home' button: {e}", "FAIL")


# -------------------------------------------------------------
#  REPORT
# -------------------------------------------------------------
def generate_report():
    html = """<html><head><title>Automation Report</title></head><body>
<h2>HP Account Automation Report - SIGN IN, SCAN & RETURN HOME</h2><table border='1'>
<tr><th>Step</th><th>Status</th></tr>"""
    for desc, status in REPORT:
        html += f"<tr><td>{desc}</td><td>{status}</td></tr>"
    html += "</table></body></html>"

    with open("automation_report.html", "w") as f:
        f.write(html)

    print("Report generated: automation_report.html")


# -------------------------------------------------------------
#  MAIN FLOW (SIGN-IN + CLICK SCAN + RETURN HOME)
# -------------------------------------------------------------
def main():
    # Fixed credentials for sign-in
    email_id = "billu123@mailsac.com"
    password = "Sonu@123"
    log_step(f"Using sign-in credentials: {email_id} / ********")

    # Launch HP Smart and navigate to sign-in
    desktop = launch_hp_smart()
    if not desktop:
        log_step("Desktop handle is None, aborting flow.", "FAIL")
        generate_report()
        return

    # Give the browser time to open the HP account sign-in page
    time.sleep(5)

    # Fill HP Account sign-in form
    sign_in_hp_account(desktop, email_id, password)

    # No Selenium driver currently used; keep placeholder for future
    accept_alert_if_present(driver=None)

    # After successful sign-in, click on the Scan button in HP Smart
    click_scan_button(desktop)

    # On the Scan screen, click "Return Home"
    click_return_home_button(desktop)

    # Generate final HTML report
    generate_report()


# -------------------------------------------------------------
#  PYTEST ENTRY POINT
# -------------------------------------------------------------
if __name__ == "__main__":
    main()


def test_hp_account_sign_in():
    main()
    assert True
