"""Send consented WhatsApp messages from contacts.xlsx and save reports."""

from datetime import date, datetime
import json
from pathlib import Path
import random
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


FOLDER = Path(__file__).resolve().parent
CONTACTS_FILE = FOLDER / "contacts.xlsx"
SCREENSHOT_FOLDER = FOLDER / "whatsapp_screenshots"
DEFAULT_MESSAGE = "Hello {name}, hope you are doing well."

# Delays in milliseconds. Lower = faster, but very short gaps between contacts
# look like spam to WhatsApp and can get the number restricted.
STEP_DELAY_MS = (800, 1800)          # small pauses inside one contact's send
BETWEEN_CONTACTS_MS = (3000, 6000)   # pause before moving to the next contact

# Selectors for the media preview's Send control (WhatsApp uses a div role="button",
# not a <button>). Add more here if WhatsApp changes its markup.
SEND_SELECTOR = (
    '[role="button"][aria-label="Send"], button[aria-label="Send"], '
    '[role="button"]:has([data-icon="send"]), '
    '[role="button"]:has([data-icon="wds-ic-send-filled"]), '
    '[data-icon="send"], [data-icon="wds-ic-send-filled"]'
)


def read_contacts():
    """Read Name and Phone from Excel; Message and Poster are optional."""
    if not CONTACTS_FILE.exists():
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Name", "Phone", "Message", "Poster"])
        workbook.save(CONTACTS_FILE)
        raise FileNotFoundError(f"Created {CONTACTS_FILE.name}. Add contacts, then run again.")

    sheet = load_workbook(CONTACTS_FILE, data_only=True).active
    headers = [str(cell.value or "").strip().lower() for cell in sheet[1]]
    if "name" not in headers or "phone" not in headers:
        raise ValueError("contacts.xlsx must have Name and Phone columns; Message is optional.")

    columns = {header: index for index, header in enumerate(headers)}
    contacts = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        name = str(values[columns["name"]] or "").strip()
        phone = re.sub(r"\D", "", str(values[columns["phone"]] or ""))
        message = str(values[columns["message"]] or "").strip() if "message" in columns else ""
        poster = str(values[columns["poster"]] or "").strip() if "poster" in columns else ""
        if name or phone:
            contacts.append({"name": name, "phone": phone, "message": message, "poster": poster})
    return contacts


def save_reports(results):
    """Write the full JSON report and a readable Excel summary."""
    date_tag = date.today().isoformat()
    json_path = FOLDER / f"whatsapp_report_{date_tag}.json"
    excel_path = FOLDER / f"whatsapp_report_{date_tag}.xlsx"

    json_path.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Message Report"
    sheet.append(["Name", "Phone", "Status", "Message", "Poster", "Last 3 Messages", "Screenshot", "Error"])
    for result in results:
        sheet.append([
            result["name"], result["phone"], result["status"], result["message"], result["poster"],
            "\n".join(result["last_3_messages"]) or "No messages found in chat",
            result["screenshot"], result["error"],
        ])

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="146B55")
    for row in sheet.iter_rows(min_row=2, min_col=6, max_col=6):
        row[0].alignment = Alignment(wrap_text=True, vertical="top")
    for column, width in {"A": 22, "B": 20, "C": 18, "D": 42, "E": 42, "F": 48, "G": 42, "H": 36}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    try:
        workbook.save(excel_path)
    except PermissionError:
        # Excel can lock an open workbook, so preserve this run under a new filename.
        excel_path = FOLDER / f"whatsapp_report_{date_tag}_{datetime.now():%H%M%S}.xlsx"
        workbook.save(excel_path)
        print("The dated Excel report was open or locked; saved a timestamped copy instead.")
    print(f"Reports saved:\n{json_path}\n{excel_path}")


def wait_for_login_confirmation():
    """Require an interactive confirmation before sending any messages."""
    try:
        input("Scan the QR code if needed. After WhatsApp chats load, press Enter here.")
    except EOFError:
        print("Interactive input is unavailable; no messages were sent.")
        return False
    return True


def attach_poster(page, poster_path):
    """Attach an image through the Attach menu; fall back to the hidden file input."""
    try:
        page.locator(
            '[data-icon="plus-rounded"], [data-icon="attach-menu-plus"], '
            '[aria-label="Attach"], [title="Attach"]'
        ).first.click(timeout=10000)
        with page.expect_file_chooser(timeout=10000) as chooser_info:
            page.get_by_text("Photos & videos").first.click()
        chooser_info.value.set_files(str(poster_path))
    except PlaywrightTimeoutError:
        # Fallback: set the file directly on the image input.
        page.keyboard.press("Escape")
        image_input = page.locator('input[type="file"][accept*="image"]').first
        if image_input.count() == 0:
            raise RuntimeError("WhatsApp's image upload control was not found.")
        image_input.set_input_files(str(poster_path))


def dump_preview_debug(page):
    """Save a screenshot and print clickable/labelled elements to help fix selectors."""
    SCREENSHOT_FOLDER.mkdir(exist_ok=True)
    shot = SCREENSHOT_FOLDER / f"debug_preview_{datetime.now():%H%M%S}.png"
    page.screenshot(path=str(shot))
    items = page.evaluate(
        """() => [...document.querySelectorAll('[aria-label],[data-icon],[role="button"]')]
            .filter(e => e.offsetParent !== null)
            .slice(-40)
            .map(e => `${e.tagName} role=${e.getAttribute('role')} aria-label=${e.getAttribute('aria-label')} icon=${e.getAttribute('data-icon')}`)"""
    )
    print("--- DEBUG: visible elements (last 40) ---")
    for item in items:
        print(item)
    print(f"--- DEBUG screenshot: {shot} ---")
    return shot


# Finds the green round Send button in the media preview by colour + position,
# so it does not depend on WhatsApp's changing element names.
FIND_SEND_JS = r"""() => {
    const green = c => {
        const m = c.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
        if (!m) return false;
        const r = +m[1], g = +m[2], b = +m[3];
        return r < 70 && g > 140 && b > 60 && b < 170;
    };
    let best = null;
    for (const e of document.querySelectorAll('div, button, span')) {
        const r = e.getBoundingClientRect();
        if (r.width < 30 || r.width > 110 || r.height < 30 || r.height > 110) continue;
        if (r.bottom <= 0 || r.top >= innerHeight) continue;
        if (r.left < innerWidth * 0.5) continue;
        if (!green(getComputedStyle(e).backgroundColor)) continue;
        if (!best || r.top > best.top) {
            best = {x: r.left + r.width / 2, y: r.top + r.height / 2, top: r.top};
        }
    }
    return best;
}"""


def find_send_button(page):
    """Return {'x','y'} of the preview's Send button, or None if no preview is open."""
    pos = page.evaluate(FIND_SEND_JS)
    if pos:
        return pos
    try:
        loc = page.locator(SEND_SELECTOR).last
        if loc.count() and loc.is_visible():
            box = loc.bounding_box()
            if box:
                return {"x": box["x"] + box["width"] / 2, "y": box["y"] + box["height"] / 2}
    except Exception:
        pass
    return None


def wait_for_send_button(page, timeout=20000, gone=False):
    """Wait until the Send button appears (gone=False) or disappears (gone=True)."""
    for _ in range(max(1, timeout // 300)):
        found = find_send_button(page) is not None
        if found != gone:
            return True
        page.wait_for_timeout(300)
    return False


def send_poster(page, poster_path, message):
    """Attach an image, add the caption, and send it automatically."""
    attach_poster(page, poster_path)

    if not wait_for_send_button(page, 25000):
        shot = dump_preview_debug(page)
        raise RuntimeError(f"Poster preview did not open. See {shot}")
    page.wait_for_timeout(2500)  # let the image finish loading

    # The caption box is normally focused already; click it if we can find it.
    try:
        page.locator('[aria-placeholder*="aption" i]').last.click(timeout=3000)
    except PlaywrightTimeoutError:
        pass
    page.keyboard.type(message, delay=15)
    page.wait_for_timeout(1000)

    # 1) Enter key
    page.keyboard.press("Enter")
    if wait_for_send_button(page, 5000, gone=True):
        print("  poster sent via Enter key")
        return

    # 2) Mouse click on the green Send button, up to 3 tries
    for attempt in range(1, 4):
        pos = find_send_button(page)
        if pos is None:
            print("  poster sent")
            return
        page.mouse.click(pos["x"], pos["y"])
        if wait_for_send_button(page, 5000, gone=True):
            print(f"  poster sent via mouse click (try {attempt})")
            return

    # 3) JavaScript click on whatever is at the button's position
    pos = find_send_button(page)
    if pos:
        page.evaluate(
            "([x, y]) => { const e = document.elementFromPoint(x, y);"
            " if (e) (e.closest('[role=button],button') || e).click(); }",
            [pos["x"], pos["y"]],
        )
        if wait_for_send_button(page, 5000, gone=True):
            print("  poster sent via JS click")
            return

    shot = dump_preview_debug(page)
    raise RuntimeError(f"Poster preview stayed open after all send attempts. See {shot}")


READ_MESSAGES_JS = r"""() => {
    const root = document.querySelector('#main') || document;
    let nodes = [...root.querySelectorAll('[data-id]')]
        .filter(e => /^(true|false)_/.test(e.getAttribute('data-id')));
    if (!nodes.length) nodes = [...root.querySelectorAll('[data-pre-plain-text]')];
    if (!nodes.length) nodes = [...root.querySelectorAll('[role="row"]')];
    const seen = new Set();
    const out = [];
    for (const e of nodes) {
        const id = e.getAttribute('data-id') || '';
        if (id) { if (seen.has(id)) continue; seen.add(id); }
        let text = (e.innerText || '').replace(/\s+/g, ' ').trim();
        if (e.querySelector('img[src^="blob:"]')) text = ('[image] ' + text).trim();
        if (!text) continue;
        out.push((id.startsWith('true_') ? 'You: ' : id ? 'Them: ' : '') + text);
    }
    return out.slice(-3);
}"""

COUNT_OUT_IMAGES_JS = r"""() => document.querySelectorAll(
    '#main [data-id^="true_"] img[src^="blob:"]').length"""


def read_last_messages(page):
    """Return the last 3 chat messages (prefixed You:/Them:) without relying on CSS class names."""
    try:
        return page.evaluate(READ_MESSAGES_JS)
    except Exception:
        return []


def send_to_contact(page, contact):
    """Open a WhatsApp chat by phone, send one message, then collect replies."""
    phone = contact["phone"]
    if not contact["name"] or not 8 <= len(phone) <= 15:
        raise ValueError("Name is required and Phone must include country code (8-15 digits).")

    message_template = contact["message"] or DEFAULT_MESSAGE
    message = message_template.replace("{name}", contact["name"])
    poster_value = contact.get("poster", "")
    poster_path = Path(poster_value) if poster_value else None
    if poster_path and not poster_path.is_absolute():
        poster_path = FOLDER / poster_path
    if poster_path and not poster_path.is_file():
        raise FileNotFoundError(f"Poster image not found: {poster_path}")

    page.goto(f"https://web.whatsapp.com/send?phone={phone}", wait_until="domcontentloaded")
    page.wait_for_timeout(random.randint(*STEP_DELAY_MS))

    # Wait for the chat composer; an unavailable chat is recorded as a contact error.
    page.wait_for_selector('footer div[contenteditable="true"]', state="visible", timeout=30000)

    outgoing_image_count = 0
    if poster_path:
        outgoing_image_count = page.evaluate(COUNT_OUT_IMAGES_JS)
        send_poster(page, poster_path, message)
    else:
        composer = page.locator('footer div[contenteditable="true"]').last
        composer.fill(message)
        page.wait_for_timeout(random.randint(*STEP_DELAY_MS))
        composer.press("Enter")

    # WhatsApp's message container classes can change; look for the visible message text.
    try:
        page.get_by_text(message, exact=True).last.wait_for(state="visible", timeout=20000)
        if poster_path:
            page.wait_for_function(
                "n => document.querySelectorAll('#main [data-id^=\"true_\"] img[src^=\"blob:\"]').length > n",
                arg=outgoing_image_count, timeout=20000,
            )
        status = "Sent"
        confirmation_error = ""
    except PlaywrightTimeoutError:
        # The send may have succeeded; don't retry and risk sending a duplicate.
        status = "Unconfirmed"
        confirmation_error = "Message and/or poster may have been sent; check this chat before retrying."
    page.wait_for_timeout(random.randint(*STEP_DELAY_MS))

    SCREENSHOT_FOLDER.mkdir(exist_ok=True)
    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", contact["name"]).strip("_") or phone
    screenshot = SCREENSHOT_FOLDER / f"{safe_name}_{datetime.now():%H%M%S}.png"
    page.screenshot(path=str(screenshot))

    page.wait_for_timeout(1000)  # let the latest bubbles render
    last_messages = read_last_messages(page)
    return message, str(poster_path) if poster_path else "", str(screenshot), last_messages, status, confirmation_error


def main():
    contacts = read_contacts()
    if not contacts:
        raise ValueError("No contacts found. Add contact rows to contacts.xlsx first.")

    print(f"Loaded {len(contacts)} contact(s). Only proceed if they agreed to receive messages.")
    results = []
    with sync_playwright() as playwright:
        # A persistent profile remembers the WhatsApp login after the first QR scan.
        context = playwright.chromium.launch_persistent_context(
            str(FOLDER / "whatsapp_profile"), channel="chrome", headless=False
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")
            if not wait_for_login_confirmation():
                return
            print("Sending to the contacts in contacts.xlsx. Only message people who agreed to receive them.")

            for index, contact in enumerate(contacts):
                result = {
                    "name": contact["name"], "phone": contact["phone"], "status": "Failed",
                    "message": "", "poster": contact.get("poster", ""),
                    "last_3_messages": [], "screenshot": "", "error": "",
                }
                try:
                    message, poster, screenshot, last_messages, status, confirmation_error = send_to_contact(page, contact)
                    result.update(
                        status=status,
                        message=message,
                        poster=poster,
                        screenshot=screenshot,
                        last_3_messages=last_messages,
                        error=confirmation_error,
                    )
                except Exception as error:
                    result["error"] = str(error)
                    print(f"Could not message {contact['name']}: {error}")
                results.append(result)
                print(f"{contact['name']}: {result['status']}")
                if index < len(contacts) - 1:
                    page.wait_for_timeout(random.randint(*BETWEEN_CONTACTS_MS))
        finally:
            context.close()

    save_reports(results)


if __name__ == "__main__":
    main()