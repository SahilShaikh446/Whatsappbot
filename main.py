# -*- coding: utf-8 -*-
"""
SmartEdge WhatsApp Assistant (Windows 7 / MS Office 2007 Edition)
- Fully automated workflow: reads contacts from Excel, poster from POSTS/, and message from message.txt
- Choice of sending mode: Both (Poster + Caption), Only Poster, Only Message, or Both Separate
- Compatible with: Windows 7, Python 3.8.10, ChromeDriver 109, MS Office 2007 (.xlsx)
"""

from __future__ import print_function

import json
import os
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
import pyperclip
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

# Base paths
BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
EXCEL_FILE = BASE_DIR / "students_contacts.xlsx"
MESSAGE_FILE = BASE_DIR / "message.txt"
POSTS_DIR = BASE_DIR / "POSTS"
PROFILE_DIR = BASE_DIR / "whatsapp_chrome_profile"
DEBUG_DIR = BASE_DIR / "whatsapp_debug"

# Ensure directories exist
DEBUG_DIR.mkdir(exist_ok=True)
POSTS_DIR.mkdir(exist_ok=True)

DEFAULT_CONFIG = {
    "excel_file": "students_contacts.xlsx",
    "message_file": "message.txt",
    "posts_folder": "POSTS",
    "auto_send": True,
    "min_delay_seconds": 25,
    "max_delay_seconds": 45,
    "batch_size": 15,
    "batch_pause_minutes": 5,
    "require_opt_in": True,
    "max_contacts_per_run": None,
}


class AutomationError(RuntimeError):
    pass


def load_config():
    """Load config.json or return defaults."""
    if CONFIG_FILE.is_file():
        try:
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                cfg = json.load(f)
                merged = dict(DEFAULT_CONFIG)
                merged.update(cfg)
                return merged
        except Exception as e:
            print("[WARN] Error loading config.json, using defaults: {}".format(e))
    return dict(DEFAULT_CONFIG)


def load_message_template():
    """Load message template from message.txt (created if missing)."""
    if not MESSAGE_FILE.is_file():
        default_tpl = (
            "Hi {Name}! 👋\n\n"
            "🎓 Degree complete, but still looking for a job?\n\n"
            "Build job-ready skills with SmartEdge Academy!\n\n"
            "✅ Python | Java | Data Analytics\n"
            "✅ AI & Automation | Full Stack | Digital Marketing\n"
            "✅ Practical training and real projects\n"
            "✅ Resume building and interview preparation\n"
            "✅ 100% Job Assistance (not a job guarantee)\n\n"
            "🎁 FREE Career Counselling!\n\n"
            "Reply YES if you would like details.\n\n"
            "🌐 academy.smartedgegroup.in\n"
            "📞 9820624436\n"
            "📍 Bhiwandi"
        )
        try:
            with open(str(MESSAGE_FILE), "w", encoding="utf-8") as f:
                f.write(default_tpl)
            print("[INFO] Created default message.txt. You can edit this file in Notepad.")
        except Exception:
            pass
        return default_tpl

    try:
        with open(str(MESSAGE_FILE), "r", encoding="utf-8") as f:
            content = f.read().strip()
            if content:
                return content
    except Exception as e:
        print("[WARN] Could not read message.txt with UTF-8: {}".format(e))
        try:
            with open(str(MESSAGE_FILE), "r", encoding="latin-1") as f:
                return f.read().strip()
        except Exception:
            pass

    return "Hi {Name}! 👋\n\nHope you are doing well!"


def get_post_image(posts_dir):
    """Find a valid image in POSTS/ folder."""
    if not posts_dir.is_dir():
        return None
    valid_extensions = (".jpg", ".jpeg", ".png", ".webp")
    images = sorted(
        [p for p in posts_dir.iterdir() if p.is_file() and p.suffix.lower() in valid_extensions],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not images:
        return None
    return images[0]


def format_phone_number(value):
    """
    Sanitize and format phone number for WhatsApp.
    Handles 10 digits, 11 digits (with leading 0), or international with country code.
    """
    if value is None:
        return ""
    if isinstance(value, float):
        if not value.is_integer():
            return ""
        value = int(value)
    digits = re.sub(r"\D", "", str(value).strip())
    if len(digits) == 11 and digits.startswith("0"):
        digits = "91" + digits[1:]
    elif len(digits) == 10:
        digits = "91" + digits
    return digits if 10 <= len(digits) <= 15 else ""


def is_visible(element):
    """Safely check if an element is visible in the DOM."""
    try:
        return element.is_displayed() and element.size.get("width", 0) > 0
    except (StaleElementReferenceException, WebDriverException):
        return False


def wait_until(driver, condition_fn, timeout=30, label="element"):
    """Wait for condition or raise AutomationError."""
    try:
        return WebDriverWait(driver, timeout, poll_frequency=0.5).until(condition_fn)
    except TimeoutException as exc:
        raise AutomationError("Timed out waiting for {}".format(label)) from exc


def visible_elements(driver, by, selector):
    """Return all visible elements matching selector."""
    try:
        return [e for e in driver.find_elements(by, selector) if is_visible(e)]
    except Exception:
        return []


def capture_debug_screenshot(driver, row_idx, stage):
    """Capture failure screenshot for troubleshooting."""
    try:
        filename = DEBUG_DIR / "row_{}_{}_{}.png".format(
            row_idx, stage, datetime.now().strftime("%Y%m%d_%H%M%S")
        )
        driver.save_screenshot(str(filename))
        print("   [DEBUG] Screenshot saved to: {}".format(filename.name))
    except Exception:
        pass


def get_driver():
    """Initialize Selenium WebDriver with Windows 7 Chrome 109 support."""
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=" + str(PROFILE_DIR))
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--ignore-certificate-errors")

    # Anti-crash switches for low-spec Windows 7
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver_path = BASE_DIR / "chromedriver.exe"
    if driver_path.is_file():
        try:
            from selenium.webdriver.chrome.service import Service
            service = Service(executable_path=str(driver_path))
            return webdriver.Chrome(service=service, options=options)
        except Exception:
            return webdriver.Chrome(executable_path=str(driver_path), options=options)

    return webdriver.Chrome(options=options)


def outgoing_message_count(driver):
    """Count visible outgoing messages in the active chat."""
    try:
        return len(driver.find_elements(By.CSS_SELECTOR, "#main .message-out"))
    except Exception:
        return 0


def open_chat_for_contact(driver, name, phone):
    """
    Reliably opens the chat with recipient.
    Strategy 1: Direct URL navigation (web.whatsapp.com/send?phone=...)
                Waits for chat composer. Only checks for real visible dialogs.
    Strategy 2: Automatic fallback to searching contact in WhatsApp Web search bar
                (handles contacts synced via Google Contacts).
    """
    url = "https://web.whatsapp.com/send?phone={}".format(phone)
    driver.get(url)

    def check_chat_state(d):
        # 1. Chat composer is ready!
        if visible_elements(d, By.XPATH, '//div[@id="main"]//footer//*[@contenteditable="true"]'):
            return "ready"

        # 2. Check ONLY for an ACTUAL visible modal popup dialog on screen
        dialogs = visible_elements(d, By.XPATH, '//div[@role="dialog"]')
        for dlg in dialogs:
            try:
                dlg_text = dlg.text.lower()
                if (
                    "phone number shared via url is invalid" in dlg_text
                    or "not on whatsapp" in dlg_text
                    or "url is invalid" in dlg_text
                ):
                    ok_btns = dlg.find_elements(By.XPATH, './/button')
                    if ok_btns:
                        try:
                            ok_btns[0].click()
                        except Exception:
                            pass
                    return "invalid_dialog"
            except Exception:
                pass
        return False

    try:
        state = wait_until(driver, check_chat_state, timeout=35, label="chat composer")
        if state == "ready":
            return True
    except AutomationError:
        pass

    # Strategy 2: Search via sidebar (especially for Google Contacts or URL restriction)
    print("   [i] Trying contact search in WhatsApp sidebar...")
    try:
        search_boxes = visible_elements(
            driver, By.XPATH, '//div[@id="side"]//div[@contenteditable="true"]'
        )
        if search_boxes:
            s_box = search_boxes[0]
            s_box.click()
            time.sleep(0.3)
            s_box.send_keys(Keys.CONTROL, "a")
            s_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.3)

            # Try last 10 digits
            search_query = phone[-10:] if len(phone) >= 10 else phone
            s_box.send_keys(search_query)
            time.sleep(2.5)

            contact_rows = visible_elements(
                driver,
                By.XPATH,
                '//div[@id="pane-side"]//div[@role="listitem"] | //div[@id="pane-side"]//div[@role="row"]'
            )
            if contact_rows:
                contact_rows[0].click()
                time.sleep(1.5)
                if visible_elements(driver, By.XPATH, '//div[@id="main"]//footer//*[@contenteditable="true"]'):
                    return True

            # If not found by phone, try searching by Name
            if name and name.lower() not in ("student", "there", "customer", ""):
                s_box.send_keys(Keys.CONTROL, "a")
                s_box.send_keys(Keys.BACKSPACE)
                time.sleep(0.3)
                s_box.send_keys(name)
                time.sleep(2.5)
                contact_rows = visible_elements(
                    driver,
                    By.XPATH,
                    '//div[@id="pane-side"]//div[@role="listitem"] | //div[@id="pane-side"]//div[@role="row"]'
                )
                if contact_rows:
                    contact_rows[0].click()
                    time.sleep(1.5)
                    if visible_elements(driver, By.XPATH, '//div[@id="main"]//footer//*[@contenteditable="true"]'):
                        return True
    except Exception as e:
        print("   [DEBUG] Sidebar search fallback error: {}".format(e))

    return False


def send_poster(driver, image_path, caption=""):
    """
    Attaches image via WhatsApp Web's Attach (+) menu, inserts caption in the preview dialog,
    and sends the post.
    """
    if not image_path or not image_path.is_file():
        raise AutomationError("Image file not found: {}".format(image_path))

    count_before = outgoing_message_count(driver)

    # Step 1: Click the Attach (+) button in chat footer
    attach_btns = visible_elements(
        driver,
        By.XPATH,
        '//div[@id="main"]//footer//*[@data-icon="plus"]/ancestor::*[self::button or @role="button"][1] | '
        '//div[@id="main"]//footer//*[@title="Attach" or @aria-label="Attach"] | '
        '//div[@id="main"]//footer//span[@data-icon="plus"]'
    )
    if not attach_btns:
        raise AutomationError("Attach (+) button not found in chat footer")

    try:
        attach_btns[0].click()
    except Exception:
        driver.execute_script("arguments[0].click();", attach_btns[0])

    time.sleep(1.0)

    # Step 2: Locate the Photos & Videos file input
    file_inputs = [
        e for e in driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]') if e.is_enabled()
    ]
    if not file_inputs:
        time.sleep(1.0)
        file_inputs = [
            e for e in driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]') if e.is_enabled()
        ]

    if not file_inputs:
        raise AutomationError("Could not find file upload input after clicking Attach (+)")

    image_inputs = [
        e for e in file_inputs if "image" in (e.get_attribute("accept") or "").lower()
    ]
    target_input = image_inputs[0] if image_inputs else file_inputs[0]

    # Step 3: Send file path to trigger WhatsApp media preview modal
    target_input.send_keys(str(image_path.resolve()))
    print("   [+] Poster attached via (+) menu. Waiting for preview dialog...")

    # Step 4: Wait for the media preview modal (div[role="dialog"])
    def get_preview_dialog(d):
        dialogs = visible_elements(d, By.XPATH, '//div[@role="dialog"]')
        for dlg in dialogs:
            if dlg.find_elements(By.XPATH, './/*[@data-icon="send"] | .//img | .//canvas | .//*[@contenteditable="true"]'):
                return dlg
        return False

    preview_dialog = wait_until(driver, get_preview_dialog, timeout=35, label="media preview dialog")

    # Step 5: Insert caption inside the preview dialog if provided
    if caption:
        caption_editors = preview_dialog.find_elements(By.XPATH, './/div[@contenteditable="true"]')
        if caption_editors:
            c_box = caption_editors[0]
            try:
                c_box.click()
            except Exception:
                pass
            time.sleep(0.3)
            pyperclip.copy(caption)
            c_box.send_keys(Keys.CONTROL, "v")
            time.sleep(0.8)
            print("   [+] Caption inserted into preview dialog")

    # Step 6: Click Send inside preview dialog
    send_btns = preview_dialog.find_elements(
        By.XPATH,
        './/*[@data-icon="send"]/ancestor::*[self::button or @role="button"][1] | .//button[@aria-label="Send"]'
    )
    if send_btns:
        try:
            send_btns[0].click()
        except Exception:
            driver.execute_script("arguments[0].click();", send_btns[0])
    else:
        # Fallback: Enter key inside caption editor
        caption_editors = preview_dialog.find_elements(By.XPATH, './/div[@contenteditable="true"]')
        if caption_editors:
            caption_editors[0].send_keys(Keys.ENTER)

    print("   [+] Send clicked on poster. Waiting for delivery...")

    # Step 7: Wait for preview dialog to close
    start_time = time.time()
    confirmed = False
    while time.time() - start_time < 30:
        if not is_visible(preview_dialog) or outgoing_message_count(driver) > count_before:
            confirmed = True
            break
        time.sleep(0.8)

    return "Sent" if confirmed else "Review"


def send_text_only(driver, message_text):
    """Sends text-only message via chat composer in footer."""
    count_before = outgoing_message_count(driver)
    chat_box_matches = visible_elements(
        driver, By.XPATH, '//div[@id="main"]//footer//*[@contenteditable="true"]'
    )
    if not chat_box_matches:
        raise AutomationError("Chat input box not found in footer")

    box = chat_box_matches[0]
    box.click()
    time.sleep(0.3)
    pyperclip.copy(message_text)
    box.send_keys(Keys.CONTROL, "v")
    time.sleep(0.8)

    send_btns = visible_elements(
        driver,
        By.XPATH,
        '//div[@id="main"]//footer//*[@data-icon="send"]/ancestor::*[self::button or @role="button"][1]'
    )
    if send_btns:
        try:
            send_btns[0].click()
        except Exception:
            box.send_keys(Keys.ENTER)
    else:
        box.send_keys(Keys.ENTER)

    print("   [+] Message sent. Waiting for delivery...")
    start_time = time.time()
    confirmed = False
    while time.time() - start_time < 20:
        if outgoing_message_count(driver) > count_before:
            confirmed = True
            break
        time.sleep(0.8)

    return "Sent" if confirmed else "Review"


def set_excel_status(wb, ws, row_idx, status, notes=""):
    """Update status, timestamp, and notes in MS Office 2007 Excel sheet."""
    ws.cell(row_idx, 3).value = status
    ws.cell(row_idx, 4).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.cell(row_idx, 5).value = str(notes)[:250]

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
    except PermissionError:
        print("   [WARN] Could not save Excel file (is it open in Excel? Close it).")


def main():
    print("=" * 65)
    print(" SMARTEDGE WHATSAPP ASSISTANT (WINDOWS 7 / EXCEL 2007)")
    print("=" * 65)

    config = load_config()
    template = load_message_template()

    # Step: User Choice of what to send
    print("\nWHAT WOULD YOU LIKE TO SEND TODAY?")
    print("  1 = Both (Poster with Message caption) [RECOMMENDED]")
    print("  2 = Only Poster (Image only, no text)")
    print("  3 = Only Message (Text from message.txt only)")
    print("  4 = Both Separate (Poster first, then Message right after)")
    print("-" * 65)
    choice = input("Enter choice (1, 2, 3, or 4) [Press Enter for 1]: ").strip()
    if choice == "2":
        send_mode = "poster_only"
        print("[*] Selected: ONLY POSTER")
    elif choice == "3":
        send_mode = "message_only"
        print("[*] Selected: ONLY MESSAGE")
    elif choice == "4":
        send_mode = "both_separate"
        print("[*] Selected: BOTH SEPARATE (Poster, then Message)")
    else:
        send_mode = "both_caption"
        print("[*] Selected: BOTH (Poster with Caption)")

    print("[*] Message template loaded from: {}".format(MESSAGE_FILE.name))
    print("[*] Excel file: {}".format(EXCEL_FILE.name))

    if not EXCEL_FILE.is_file():
        sys.exit("[ERROR] Excel file not found: {}".format(EXCEL_FILE))

    post_image = None
    if send_mode in ("both_caption", "poster_only", "both_separate"):
        post_image = get_post_image(POSTS_DIR)
        if post_image:
            print("[*] Selected poster image: {}".format(post_image.name))
        else:
            if send_mode == "poster_only":
                sys.exit("[ERROR] No image found in POSTS/ folder. Add an image and retry.")
            else:
                print("[WARN] No image found in POSTS/ folder. Falling back to message only.")
                send_mode = "message_only"

    wb = openpyxl.load_workbook(str(EXCEL_FILE))
    ws = wb.active

    # Validate headers
    headers = [str(ws.cell(1, col).value or "").strip().lower() for col in range(1, ws.max_column + 1)]
    if ws.max_column < 6 or "opt_in" not in headers:
        ws.cell(1, 6).value = "Opt_In"
        wb.save(str(EXCEL_FILE))

    # Collect pending contacts
    queued, seen = [], set()
    for row in range(2, ws.max_row + 1):
        name = str(ws.cell(row, 1).value or "Student").strip()
        phone = format_phone_number(ws.cell(row, 2).value)
        status = str(ws.cell(row, 3).value or "").strip().lower()
        opted_in = str(ws.cell(row, 6).value or "").strip().lower() in ("yes", "y", "true", "1")

        if status in ("sent", "preview ok"):
            continue
        if config.get("require_opt_in", True) and not opted_in:
            continue
        if not phone:
            set_excel_status(wb, ws, row, "Failed", "Invalid phone number")
            continue
        if phone in seen:
            set_excel_status(wb, ws, row, "Review", "Duplicate number in Excel")
            continue

        seen.add(phone)
        queued.append({"row": row, "name": name, "phone": phone})

    max_limit = config.get("max_contacts_per_run")
    if max_limit and isinstance(max_limit, int):
        queued = queued[:max_limit]

    print("[*] Eligible contacts queued: {}".format(len(queued)))
    if not queued:
        print("[*] No contacts to process. Check Excel Status or Opt_In column.")
        wb.close()
        return

    auto_send = config.get("auto_send", True)
    min_delay = int(config.get("min_delay_seconds", 25))
    max_delay = int(config.get("max_delay_seconds", 45))
    batch_size = int(config.get("batch_size", 15))
    batch_pause_min = int(config.get("batch_pause_minutes", 5))

    print("[*] Mode: Fully Automated")
    print("[*] Delay between messages: {}s - {}s".format(min_delay, max_delay))
    print("[*] Batch cooling: pause {} min every {} messages".format(batch_pause_min, batch_size))
    print("-" * 65)

    driver = None
    sent_count = 0
    try:
        driver = get_driver()
        driver.get("https://web.whatsapp.com")
        print("\n[!] Scan QR code if not already logged in...")

        wait_until(
            driver,
            lambda d: visible_elements(d, By.CSS_SELECTOR, "#side")
            or visible_elements(d, By.CSS_SELECTOR, "#pane-side"),
            timeout=150,
            label="WhatsApp Web login",
        )
        print("[OK] WhatsApp Web session active!\n")

        for idx, contact in enumerate(queued, start=1):
            row_idx = contact["row"]
            name = contact["name"]
            phone = contact["phone"]
            personalized_msg = template.replace("{Name}", name).replace("{name}", name)

            print("[{}/{}] Recipient: {} (+{})".format(idx, len(queued), name, phone))

            try:
                chat_opened = open_chat_for_contact(driver, name, phone)
                if not chat_opened:
                    print("   [-] Could not open chat for +{} (number not found or restricted).".format(phone))
                    capture_debug_screenshot(driver, row_idx, "chat_open_failed")
                    set_excel_status(wb, ws, row_idx, "Failed", "Could not open chat")
                    continue

                # Execute chosen send mode
                if send_mode == "poster_only":
                    res = send_poster(driver, post_image, caption="")

                elif send_mode == "message_only":
                    res = send_text_only(driver, personalized_msg)

                elif send_mode == "both_separate":
                    res = send_poster(driver, post_image, caption="")
                    if res == "Sent":
                        time.sleep(2.0)
                        res = send_text_only(driver, personalized_msg)

                else:  # both_caption
                    res = send_poster(driver, post_image, caption=personalized_msg)

                set_excel_status(
                    wb,
                    ws,
                    row_idx,
                    res,
                    "Delivered successfully" if res == "Sent" else "Verify in WhatsApp",
                )
                if res == "Sent":
                    sent_count += 1
                    print("   [✓] Successfully sent to {}!".format(name))

            except Exception as exc:
                capture_debug_screenshot(driver, row_idx, "error")
                print("   [!] Error: {}".format(exc))
                set_excel_status(wb, ws, row_idx, "Failed", str(exc)[:200])

            # Batch cooling pause
            if idx % batch_size == 0 and idx < len(queued):
                print("\n[PAUSE] Batch limit reached ({} messages). Resting for {} minutes...".format(
                    batch_size, batch_pause_min
                ))
                time.sleep(batch_pause_min * 60)

            # Delay between individual messages
            elif idx < len(queued):
                delay = random.uniform(min_delay, max_delay)
                print("   [⏳] Waiting {:.1f}s before next contact...".format(delay))
                time.sleep(delay)

        print("\n" + "=" * 65)
        print("Completed! Total messages processed: {}/{}".format(sent_count, len(queued)))
        print("Excel statuses updated in: {}".format(EXCEL_FILE.name))
        print("=" * 65)

    except KeyboardInterrupt:
        print("\n[!] Execution interrupted by user.")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        wb.close()


if __name__ == "__main__":
    main()
