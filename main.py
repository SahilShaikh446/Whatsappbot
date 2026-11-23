# SmartEdge WhatsApp Assistant
# Windows 7 legacy-compatible, manual-send workflow.
# Opens opted-in contacts, prepares an image + caption, and waits for you
# to manually press Send in WhatsApp Web. It does NOT auto-send or evade detection.

from __future__ import print_function

import re
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import pyperclip
from openpyxl.styles import Font, PatternFill
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "students_contacts.xlsx"
POSTS_DIR = BASE_DIR / "POSTS"
PROFILE_DIR = BASE_DIR / "whatsapp_chrome_profile"
DEBUG_DIR = BASE_DIR / "whatsapp_debug"
TIMEOUT = 30
MAX_CONTACTS_PER_RUN = 10   # Set None to process all eligible contacts manually
REQUIRE_OPT_IN = True
DRY_RUN = False             # True = prepare preview only; no manual-send prompt

MESSAGE_TEMPLATE = """Hi {Name}! 👋

🎓 Degree complete, but still looking for a job?

Build job-ready skills with SmartEdge Academy.

✅ Python | Java | Data Analytics
✅ AI & Automation | Full Stack | Digital Marketing
✅ Practical training and real projects
✅ Resume building and interview preparation
✅ 100% Job Assistance (not a job guarantee)

🎁 FREE Career Counselling!

Reply YES if you would like details.

🌐 academy.smartedgegroup.in
📞 9820624436
📍 Bhiwandi"""

CHAT_BOX_XPATH = '//div[@id="main"]//footer//*[@contenteditable="true"]'
ATTACH_XPATH = (
    '//div[@id="main"]//footer//*[@title="Attach" or @aria-label="Attach"]'
    ' | //div[@id="main"]//footer//*[@data-icon="plus"]'
    '/ancestor::*[self::button or @role="button"][1]'
)
IMAGE_INPUT_CSS = 'input[type="file"][accept*="image"]'

class AutomationError(RuntimeError):
    pass

def phone_number(value):
    if value is None:
        return ""
    if isinstance(value, float):
        if not value.is_integer():
            return ""
        value = int(value)
    digits = re.sub(r"\D", "", str(value).strip())
    if len(digits) == 10:
        digits = "91" + digits
    return digits if len(digits) == 12 and digits.startswith("91") else ""

def choose_post():
    images = sorted(
        [p for p in POSTS_DIR.iterdir()
         if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp")],
        key=lambda p: p.name.lower()
    )
    if not images:
        raise AutomationError("No JPG, JPEG, PNG or WEBP files found in POSTS folder.")
    print("\nAvailable posts:")
    for i, path in enumerate(images, 1):
        print("  {}. {}".format(i, path.name))
    while True:
        answer = input("Choose post number (or Q to quit): ").strip().lower()
        if answer == "q":
            return None
        if answer.isdigit() and 1 <= int(answer) <= len(images):
            return images[int(answer) - 1]
        print("Invalid selection. Try again.")

def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=" + str(PROFILE_DIR))
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    # Use a normal visible browser session. No stealth / anti-detection flags.
    driver_path = BASE_DIR / "chromedriver.exe"
    if driver_path.is_file():
        return webdriver.Chrome(executable_path=str(driver_path), options=options)
    # Selenium Manager may not work on Windows 7 / older Selenium.
    # If this fails, put a matching chromedriver.exe beside this script.
    return webdriver.Chrome(options=options)

def is_visible(element):
    try:
        return element.is_displayed() and element.size.get("width", 0) > 0
    except (StaleElementReferenceException, WebDriverException):
        return False

def wait_until(driver, fn, timeout=TIMEOUT, label="UI element"):
    try:
        return WebDriverWait(driver, timeout, poll_frequency=0.5).until(fn)
    except TimeoutException as exc:
        raise AutomationError("Timed out waiting for {}".format(label)) from exc

def visible_matches(driver, by, selector):
    return [e for e in driver.find_elements(by, selector) if is_visible(e)]

def click(driver, element):
    try:
        element.click()
    except ElementClickInterceptedException:
        raise AutomationError("The WhatsApp control was obstructed. Please clear the screen and retry.")

def capture_failure(driver, row_idx, stage):
    DEBUG_DIR.mkdir(exist_ok=True)
    filename = DEBUG_DIR / "row_{}_{}_{}.png".format(
        row_idx, stage, datetime.now().strftime("%Y%m%d_%H%M%S")
    )
    try:
        driver.save_screenshot(str(filename))
        print("Screenshot saved:", filename)
    except WebDriverException:
        pass

def set_status(wb, ws, row_idx, status, notes=""):
    ws.cell(row_idx, 3).value = status
    ws.cell(row_idx, 4).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.cell(row_idx, 5).value = str(notes)[:300]
    palette = {
        "Sent": ("DCFCE7", "166534"),
        "Failed": ("FEE2E2", "991B1B"),
        "Review": ("FEF3C7", "92400E"),
        "Preview OK": ("DBEAFE", "1E40AF"),
        "Skipped": ("E5E7EB", "374151"),
    }
    if status in palette:
        bg, fg = palette[status]
        cell = ws.cell(row_idx, 3)
        cell.fill = PatternFill(fill_type="solid", fgColor=bg)
        cell.font = Font(name="Calibri", size=10, bold=True, color=fg)
    try:
        wb.save(str(EXCEL_FILE))
    except PermissionError as exc:
        raise AutomationError(
            "Close '{}' in Microsoft Excel before running.".format(EXCEL_FILE.name)
        ) from exc

def require_media_input(driver):
    inputs = [e for e in driver.find_elements(By.CSS_SELECTOR, IMAGE_INPUT_CSS) if e.is_enabled()]
    if not inputs:
        buttons = visible_matches(driver, By.XPATH, ATTACH_XPATH)
        if not buttons:
            raise AutomationError("WhatsApp Attach (+) control not found.")
        click(driver, buttons[0])
        inputs = wait_until(
            driver,
            lambda d: [e for e in d.find_elements(By.CSS_SELECTOR, IMAGE_INPUT_CSS) if e.is_enabled()],
            timeout=12,
            label="Photos & Videos file input",
        )
    return inputs[0]

def locate_caption(driver):
    xpath = (
        '//*[@contenteditable="true" and not(ancestor::footer) '
        'and not(ancestor::*[@id="side"])]'
        '[contains(translate(@aria-label,"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"caption") '
        'or contains(translate(@aria-placeholder,"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"caption") '
        'or contains(translate(@data-placeholder,"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"caption") '
        'or contains(translate(@data-testid,"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"caption")]'
    )
    candidates = visible_matches(driver, By.XPATH, xpath)
    if len(candidates) == 1:
        return candidates[0]
    # Do not guess if the preview caption field cannot be identified uniquely.
    return None

def prepare_image_caption(driver, image_path, message):
    wait_until(driver, lambda d: visible_matches(d, By.XPATH, CHAT_BOX_XPATH),
               timeout=45, label="chat composer")
    media_input = require_media_input(driver)
    media_input.send_keys(str(image_path.resolve()))
    caption = wait_until(driver, lambda d: locate_caption(d), timeout=35,
                         label="unambiguous image caption field")
    caption.click()
    pyperclip.copy(message)
    caption.send_keys("\ue009", "v")  # CTRL+V
    wait_until(
        driver,
        lambda d: message.splitlines()[0] in (caption.text or ""),
        timeout=10,
        label="caption text",
    )
    print("Image and caption prepared. Review them in WhatsApp Web.")

def main():
    if not EXCEL_FILE.is_file():
        sys.exit("Excel file not found: {}".format(EXCEL_FILE))
    post = choose_post()
    if post is None:
        print("Cancelled.")
        return

    wb = openpyxl.load_workbook(str(EXCEL_FILE))
    ws = wb.active

    if ws.max_column < 6 or str(ws.cell(1, 6).value or "").strip().lower().replace(" ", "_") not in (
        "opt_in", "opt-in", "consent"
    ):
        ws.cell(1, 6).value = "Opt_In"
        wb.save(str(EXCEL_FILE))
        wb.close()
        sys.exit("Column F header set to Opt_In. Mark YES only for people who agreed to receive messages, close Excel, and run again.")

    tasks, seen = [], set()
    for row in range(2, ws.max_row + 1):
        name = str(ws.cell(row, 1).value or "there").strip()
        phone = phone_number(ws.cell(row, 2).value)
        status = str(ws.cell(row, 3).value or "").strip().lower()
        opted_in = str(ws.cell(row, 6).value or "").strip().lower() in ("yes", "y", "true", "1")
        if status in ("sent", "review", "preview ok"):
            continue
        if REQUIRE_OPT_IN and not opted_in:
            continue
        if not phone:
            set_status(wb, ws, row, "Failed", "Invalid Indian mobile number")
            continue
        if phone in seen:
            set_status(wb, ws, row, "Review", "Duplicate number in workbook")
            continue
        seen.add(phone)
        tasks.append((row, name, phone))

    if MAX_CONTACTS_PER_RUN is not None:
        tasks = tasks[:MAX_CONTACTS_PER_RUN]
    if not tasks:
        print("No eligible contacts. Check Opt_In=Yes and status columns.")
        wb.close()
        return

    print("\nSelected post:", post.name)
    print("Eligible contacts:", len(tasks))
    print("Manual-send mode: the script will NEVER click WhatsApp's Send button.")
    driver = None
    try:
        driver = get_driver()
        driver.get("https://web.whatsapp.com")
        print("\nIf needed, scan the WhatsApp Web QR code.")
        wait_until(driver, lambda d: visible_matches(d, By.CSS_SELECTOR, "#side")
                   or visible_matches(d, By.CSS_SELECTOR, "#pane-side"),
                   timeout=150, label="WhatsApp Web chat list")
        print("WhatsApp Web is open.")

        for index, (row, name, phone) in enumerate(tasks, 1):
            print("\n[{}/{}] {} (+{})".format(index, len(tasks), name, phone))
            try:
                driver.get("https://web.whatsapp.com/send?phone={}".format(phone))
                wait_until(driver, lambda d: visible_matches(d, By.XPATH, CHAT_BOX_XPATH),
                           timeout=60, label="recipient chat")
                message = MESSAGE_TEMPLATE.format(Name=name)
                prepare_image_caption(driver, post, message)

                if DRY_RUN:
                    set_status(wb, ws, row, "Preview OK", "Preview prepared; not sent")
                    input("Press Enter for next contact...")
                    continue

                print("\nMANUAL ACTION REQUIRED:")
                print("1. Check recipient, image and caption.")
                print("2. Click Send yourself in WhatsApp Web if appropriate.")
                print("3. Return here and record the result.")
                choice = input("Enter S=you sent it, R=review/unsure, K=skip, Q=quit: ").strip().lower()
                if choice == "s":
                    set_status(wb, ws, row, "Sent", "User confirmed manual send; not independently verified")
                elif choice == "r":
                    set_status(wb, ws, row, "Review", "User marked for manual verification")
                elif choice == "k":
                    set_status(wb, ws, row, "Skipped", "Skipped by user")
                elif choice == "q":
                    set_status(wb, ws, row, "Review", "Stopped by user; check WhatsApp before retrying")
                    break
                else:
                    set_status(wb, ws, row, "Review", "Unrecognised response; verify chat manually")
            except (AutomationError, WebDriverException, OSError) as exc:
                capture_failure(driver, row, "error")
                print("Error:", type(exc).__name__, str(exc))
                set_status(wb, ws, row, "Failed", "{}: {}".format(type(exc).__name__, exc))
                input("Press Enter to continue, or Ctrl+C to stop...")
    except KeyboardInterrupt:
        print("\nStopped by user. Workbook statuses have been saved.")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        wb.close()

if __name__ == "__main__":
    main()
