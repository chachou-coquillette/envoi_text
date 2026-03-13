"""
agent.py – Send an SMS to every contact listed in contacts.csv
via the Windows "Phone Link" (formerly "Your Phone") application.

Requirements
------------
- Windows 10 / 11 with the Phone Link app installed and connected to an
  Android device that has SMS permissions granted.
- Python packages listed in requirements.txt.

Usage
-----
    python agent.py
"""

import sys
import time
import logging
import re
from pathlib import Path

import pandas as pd
import pyperclip
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
from pywinauto import mouse

import config

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

KEY_ENTER = "{VK_RETURN}"
KEY_ESCAPE = "{VK_ESCAPE}"
KEY_TAB = "{VK_TAB}"
KEY_PASTE_ALT = "+{INSERT}"

PHONE_LINK_TITLE_RE = re.compile(
    r"(Phone Link|Lien avec Windows|Votre telephone|Mobile connecte|Mobile connecté)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_contacts(path: str) -> list[dict[str, str]]:
    """Return a list of {'name': str, 'phone': str} dicts from a CSV file."""
    if not Path(path).exists():
        raise FileNotFoundError(f"contacts CSV not found: {path}")
    df = pd.read_csv(path, dtype=str)
    required = {"name", "phone"}
    missing = required - set(df.columns.str.lower())
    if missing:
        raise ValueError(f"contacts CSV is missing columns: {missing}")
    df.columns = df.columns.str.lower()
    return df[["name", "phone"]].dropna().to_dict(orient="records")


def open_phone_link() -> "Application":
    """Connect to an already-running Phone Link instance or launch it."""
    title_pattern = r".*(Phone Link|Lien avec Windows|Votre telephone|Mobile connecte|Mobile connecté).*"

    def _connect_via_desktop(timeout_s: int) -> "Application | None":
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            # Enumerate all top-level UIA windows and match known localized titles.
            for win in Desktop(backend="uia").windows():
                title = (win.window_text() or "").strip()
                if not title:
                    continue
                if PHONE_LINK_TITLE_RE.search(title):
                    try:
                        return Application(backend="uia").connect(handle=win.handle)
                    except Exception:
                        continue
            time.sleep(0.5)
        return None

    try:
        app = Application(backend="uia").connect(title_re=title_pattern, timeout=5)
        logger.info("Connected to running Phone Link instance.")
    except Exception:
        # Some localized builds expose a different title than "Phone Link".
        app = _connect_via_desktop(timeout_s=2)
        if app is not None:
            logger.info("Connected to running Phone Link instance (desktop lookup).")
            return app

        logger.info("Phone Link not found -- launching it now...")
        Application(backend="uia").start(
            r"explorer.exe shell:AppsFolder\Microsoft.YourPhone_8wekyb3d8bbwe!App"
        )

        # explorer.exe is only a launcher; wait for the actual app window.
        app = _connect_via_desktop(timeout_s=config.UI_TIMEOUT)
        if app is None:
            raise RuntimeError(
                "Phone Link did not open in time. Open it manually, then retry."
            )
    return app


def get_main_window(app: "Application"):
    """Return the Phone Link main window wrapper."""
    deadline = time.time() + config.UI_TIMEOUT
    best_fallback = None
    best_area = -1

    while time.time() < deadline:
        # 1) Prefer explicit localized title matches.
        try:
            for win in app.windows():
                title = (win.window_text() or "").strip()
                if title and PHONE_LINK_TITLE_RE.search(title):
                    return app.window(handle=win.handle)
        except Exception:
            pass

        # 2) Fallback: look for a large visible WinUI window from this app.
        try:
            for win in app.windows():
                try:
                    if not win.is_visible() or not win.is_enabled():
                        continue
                    rect = win.rectangle()
                    area = rect.width() * rect.height()
                    class_name = (win.class_name() or "").strip()
                    if area > best_area and (
                        "WinUI" in class_name or "Window" in class_name
                    ):
                        best_area = area
                        best_fallback = win.handle
                except Exception:
                    continue
        except Exception:
            pass

        # 3) Cross-check Desktop windows in case app handle mapping is delayed.
        try:
            for win in Desktop(backend="uia").windows():
                title = (win.window_text() or "").strip()
                if title and PHONE_LINK_TITLE_RE.search(title):
                    return Desktop(backend="uia").window(handle=win.handle)
        except Exception:
            pass

        time.sleep(0.5)

    if best_fallback is not None:
        logger.info("Main window selected using fallback handle detection.")
        return Desktop(backend="uia").window(handle=best_fallback)

    raise RuntimeError("Could not locate Phone Link main window.")


def bring_window_to_front(window) -> None:
    """Force the app window to foreground so keyboard input goes to Phone Link."""
    try:
        # Restore if minimized/collapsed, then request focus.
        window.restore()
    except Exception:
        pass

    try:
        window.set_focus()
        time.sleep(0.2)
    except Exception:
        pass

    # Some WinUI windows need a click to become the true foreground target.
    try:
        rect = window.rectangle()
        x = rect.left + min(80, max(20, rect.width() // 8))
        y = rect.top + min(20, max(8, rect.height() // 20))
        mouse.click(button="left", coords=(x, y))
        time.sleep(0.2)
    except Exception:
        pass

    try:
        window.set_focus()
    except Exception:
        pass


def navigate_to_messages(window) -> None:
    """Click the 'Messages' navigation item."""
    selectors = [
        {"title_re": "Messages", "control_type": "ListItem"},
        {"title_re": "Messages", "control_type": "Button"},
        {"title_re": "Messages", "control_type": "TabItem"},
    ]
    for selector in selectors:
        try:
            messages_nav = window.child_window(**selector)
            if messages_nav.exists(timeout=1):
                messages_nav.click_input()
                time.sleep(1)
                return
        except Exception:
            continue
    raise RuntimeError("Could not find 'Messages' navigation item.")


def start_new_conversation(window) -> None:
    """Click the 'New message' / compose button."""
    selectors = [
        {
            "title_re": "New message|Nouveau message|Compose|Écrire",
            "control_type": "Button",
        },
        {
            "title_re": "Nouveau|New",
            "control_type": "Button",
        },
    ]

    for selector in selectors:
        try:
            compose_btn = window.child_window(**selector)
            if compose_btn.exists(timeout=1):
                compose_btn.click_input()
                time.sleep(1)
                return
        except Exception:
            continue

    # Fallback: open compose via keyboard shortcut when button is not discoverable.
    window.set_focus()
    send_keys("^v")
    time.sleep(1)


def _score_recipient_candidate(ctrl, window_rect=None) -> int:
    """Return a score for controls likely to be the recipient input."""
    try:
        info = ctrl.element_info
        name = (getattr(info, "name", "") or "").strip().lower()
        auto_id = (getattr(info, "automation_id", "") or "").strip().lower()
        ctype = (getattr(info, "control_type", "") or "").strip()
        rect = ctrl.rectangle()
    except Exception:
        return 0

    if ctype not in {"Edit", "ComboBox"}:
        return 0

    score = 1
    if any(token in name for token in ("to", "à", "destin", "recipient")):
        score += 6
    if any(token in name for token in ("search", "rechercher")):
        score -= 8
    if any(token in auto_id for token in ("to", "recipient", "search", "people", "picker")):
        score += 5
    if "search" in auto_id:
        score -= 5

    # Prefer the right pane top input (compose recipient field), not left search box.
    if window_rect is not None:
        mid_x = (window_rect.left + window_rect.right) // 2
        if rect.left >= mid_x - 20:
            score += 7
        if rect.top <= window_rect.top + 230:
            score += 3
        if rect.width() >= 220:
            score += 2

    try:
        if ctrl.is_visible():
            score += 2
        if ctrl.is_enabled():
            score += 1
    except Exception:
        pass

    return score


def _find_recipient_field(window):
    """Find recipient field using exact selectors then scored descendant fallback."""
    selectors = [
        {"title_re": r"(To|À)\\s*:?.*", "control_type": "Edit"},
        {"title_re": r"(Search|Rechercher).*", "control_type": "Edit"},
        {"auto_id": "ToTextBox", "control_type": "Edit"},
        {"auto_id": "PeoplePicker", "control_type": "ComboBox"},
    ]
    for selector in selectors:
        try:
            candidate = window.child_window(**selector)
            if candidate.exists(timeout=1):
                return candidate
        except Exception:
            continue

    scored = []
    window_rect = None
    try:
        window_rect = window.rectangle()
    except Exception:
        pass
    try:
        for ctrl in window.descendants():
            score = _score_recipient_candidate(ctrl, window_rect)
            if score > 0:
                scored.append((score, ctrl))
    except Exception:
        pass

    if not scored:
        return None

    scored.sort(key=lambda item: item[0], reverse=True)
    return scored[0][1]


def _log_recipient_candidates(window) -> None:
    """Log top controls that look like recipient fields for troubleshooting."""
    try:
        scored = []
        window_rect = None
        try:
            window_rect = window.rectangle()
        except Exception:
            pass
        for ctrl in window.descendants():
            score = _score_recipient_candidate(ctrl, window_rect)
            if score > 0:
                info = ctrl.element_info
                scored.append(
                    (
                        score,
                        {
                            "name": getattr(info, "name", "") or "",
                            "auto_id": getattr(info, "automation_id", "") or "",
                            "control_type": getattr(info, "control_type", "") or "",
                        },
                    )
                )

        scored.sort(key=lambda item: item[0], reverse=True)
        for score, meta in scored[:8]:
            logger.info(
                "Recipient candidate score=%s type=%s auto_id=%s name=%s",
                score,
                meta["control_type"],
                meta["auto_id"],
                meta["name"],
            )
    except Exception as exc:
        logger.info("Could not enumerate recipient candidates: %s", exc)


def _recipient_queries(phone: str) -> list[str]:
    """Build phone queries that improve matching in contact pickers."""
    raw = (phone or "").strip()
    queries = [raw]

    # Prefer international '+' format first for contact matching.
    if raw.startswith("00"):
        queries = ["+" + raw[2:], raw]
    if raw.startswith("+"):
        queries.append("00" + raw[1:])

    # Deduplicate while preserving order.
    unique = []
    for value in queries:
        if value and value not in unique:
            unique.append(value)
    return unique


def _fill_text_field(field, text: str) -> None:
    """Fill a text field without using Ctrl+V (reserved by Phone Link)."""
    try:
        field.set_edit_text(text)
        return
    except Exception:
        pass

    field.click_input()
    send_keys("^a{BACKSPACE}")
    pyperclip.copy(text)
    # In this app Ctrl+V opens compose; use Shift+Insert to paste into fields.
    send_keys(KEY_PASTE_ALT)
    time.sleep(0.2)


def _escape_for_send_keys(text: str) -> str:
    """Escape characters interpreted as key modifiers by send_keys."""
    escaped = text
    for char in ["+", "^", "%", "~", "(", ")", "[", "]", "{", "}"]:
        escaped = escaped.replace(char, "{" + char + "}")
    return escaped


def _fill_focused_field(text: str) -> None:
    """Type/paste into currently focused field."""
    send_keys("^a{BACKSPACE}")
    pyperclip.copy(text)
    send_keys(KEY_PASTE_ALT)
    time.sleep(0.2)


def _find_file_dialog_filename_field(dialog):
    """Find the filename input in the system Open dialog."""
    selectors = [
        {"auto_id": "1148", "control_type": "Edit"},
        {"title_re": r"File name:|Nom du fichier\s*:?.*", "control_type": "Edit"},
    ]
    for selector in selectors:
        try:
            field = dialog.child_window(**selector)
            if field.exists(timeout=1):
                return field
        except Exception:
            continue

    # Fallback: pick the widest visible edit, usually the filename box.
    best = None
    best_width = -1
    try:
        for ctrl in dialog.descendants(control_type="Edit"):
            try:
                if not ctrl.is_visible() or not ctrl.is_enabled():
                    continue
                width = ctrl.rectangle().width()
                if width > best_width:
                    best_width = width
                    best = ctrl
            except Exception:
                continue
    except Exception:
        pass
    return best


def _find_open_dialog_address_field(dialog):
    """Find the address/path input in the system Open dialog."""
    selectors = [
        {"title_re": r"Address|Adresse|Emplacement.*", "control_type": "Edit"},
        {"auto_id": "41477", "control_type": "Edit"},
        {"auto_id": "1001", "control_type": "Edit"},
    ]
    for selector in selectors:
        try:
            field = dialog.child_window(**selector)
            if field.exists(timeout=1):
                return field
        except Exception:
            continue

    best = None
    best_top = None
    try:
        dialog_rect = dialog.rectangle()
        for ctrl in dialog.descendants(control_type="Edit"):
            try:
                if not ctrl.is_visible() or not ctrl.is_enabled():
                    continue
                rect = ctrl.rectangle()
                if rect.top > dialog_rect.top + 130:
                    continue
                if best_top is None or rect.top < best_top:
                    best = ctrl
                    best_top = rect.top
            except Exception:
                continue
    except Exception:
        pass
    return best


def _activate_window(window) -> None:
    """Bring any window or dialog to the foreground for keyboard input."""
    try:
        window.restore()
    except Exception:
        pass
    try:
        window.set_focus()
        time.sleep(0.2)
    except Exception:
        pass
    try:
        rect = window.rectangle()
        mouse.click(button="left", coords=(rect.left + 40, rect.top + 20))
        time.sleep(0.2)
    except Exception:
        pass
    try:
        window.set_focus()
    except Exception:
        pass


def _get_open_dialog():
    """Return the Open dialog using the most reliable backend available."""
    deadline = time.time() + config.UI_TIMEOUT
    while time.time() < deadline:
        for backend in ("win32", "uia"):
            try:
                dialog = Desktop(backend=backend).window(title_re=r"Open|Ouvrir")
                if dialog.exists(timeout=1):
                    return dialog
            except Exception:
                continue
        time.sleep(0.3)
    raise RuntimeError("Open dialog not found")


def _click_browse_this_pc_if_present(window) -> None:
    """Click the intermediate 'Browse this PC' button shown by some Phone Link builds."""
    selectors = [
        {
            "title_re": r"Browse this PC|Parcourir ce PC|Browse|Parcourir",
            "control_type": "Button",
        },
        {
            "title_re": r"Browse this PC|Parcourir ce PC|Browse|Parcourir",
            "control_type": "ListItem",
        },
        {
            "title_re": r"Browse this PC|Parcourir ce PC|Browse|Parcourir",
            "control_type": "MenuItem",
        },
    ]

    # Check first inside the app window, then on desktop-level popups.
    search_roots = [window, Desktop(backend="uia")]
    for root in search_roots:
        for selector in selectors:
            try:
                candidate = root.child_window(**selector)
                if candidate.exists(timeout=1):
                    candidate.click_input()
                    time.sleep(0.8)
                    return
            except Exception:
                continue


def _pick_file_in_open_dialog(dialog, image_file: Path) -> None:
    """Select a file in the Open dialog, preferring a keyboard-first full-path flow."""
    _activate_window(dialog)

    # Preferred path: use the standard Windows accelerator for the filename field.
    try:
        send_keys("%n")
        time.sleep(0.3)
        _fill_focused_field(str(image_file))
        time.sleep(0.2)
        send_keys(KEY_ENTER)
        return
    except Exception:
        pass

    # Most Windows Open dialogs accept a full path directly in the filename box.
    filename_field = _find_file_dialog_filename_field(dialog)
    if filename_field is not None:
        _fill_text_field(filename_field, str(image_file))
        time.sleep(0.2)
        send_keys(KEY_ENTER)
        return

    address_field = _find_open_dialog_address_field(dialog)
    if address_field is not None:
        _fill_text_field(address_field, str(image_file.parent))
        time.sleep(0.2)
        send_keys(KEY_ENTER)
        time.sleep(0.8)
    else:
        # Fallback if the address bar is not exposed via UIA.
        send_keys("%d")
        time.sleep(0.2)
        _fill_focused_field(str(image_file.parent))
        send_keys(KEY_ENTER)
        time.sleep(0.8)

    # Fill filename box and confirm.
    filename_field = _find_file_dialog_filename_field(dialog)
    if filename_field is None:
        raise RuntimeError("Open dialog filename field not found")

    _fill_text_field(filename_field, image_file.name)
    time.sleep(0.2)
    send_keys(KEY_ENTER)


def attach_image(window, image_path: str) -> None:
    """Attach an image in compose view (for MMS) if an image path is provided."""
    raw = (image_path or "").strip()
    if not raw:
        return

    image_file = Path(raw)
    if not image_file.exists():
        raise RuntimeError(f"Image file not found: {image_file}")

    # Open attachment picker.
    attach_selectors = [
        {"title_re": r"Attach|Joindre|Ajouter|Image|Photo|Fichier", "control_type": "Button"},
        {"title_re": r"Attach|Joindre|Ajouter|Image|Photo|Fichier", "control_type": "SplitButton"},
        {"auto_id": "AttachButton", "control_type": "Button"},
    ]
    attach_btn = None
    for selector in attach_selectors:
        try:
            candidate = window.child_window(**selector)
            if candidate.exists(timeout=1):
                attach_btn = candidate
                break
        except Exception:
            continue

    if attach_btn is None:
        raise RuntimeError("Attachment button not found in compose view")

    attach_btn.click_input()
    time.sleep(1)

    # Some versions show an intermediate step before the file dialog.
    _click_browse_this_pc_if_present(window)

    dialog = _get_open_dialog()
    dialog.wait("exists enabled visible ready", timeout=config.UI_TIMEOUT)

    _pick_file_in_open_dialog(dialog, image_file)
    # Wait for attachment preview/upload in compose area.
    time.sleep(1.5)


def search_and_select_contact(window, contact: dict) -> bool:
    """
    Type the contact phone number in the recipient field and select them.

    Returns True on success, False if the contact could not be found.
    """
    try:
        recipient_field = _find_recipient_field(window)

        if recipient_field is None:
            # Compose view may not be active; trigger it via keyboard and try again.
            window.set_focus()
            send_keys("^v")
            time.sleep(1)
            recipient_field = _find_recipient_field(window)

        if recipient_field is None:
            _log_recipient_candidates(window)
            raise RuntimeError("Recipient field not found")

        recipient_field.click_input()

        # Use preferred phone format (e.g. +33...) for better matching.
        query = _recipient_queries(contact["phone"])[0]
        _fill_text_field(recipient_field, query)
        time.sleep(0.8)
        send_keys(KEY_ENTER)
        time.sleep(0.4)
        # Move from recipient field to message field.
        send_keys(KEY_TAB)
        send_keys(KEY_TAB)
        time.sleep(0.3)
        return True
    except Exception as exc:
        logger.warning("Could not select contact %s: %s", contact["name"], exc)
        return False


def type_and_send_message(window, message: str) -> None:
    """Paste the message into the compose field and send it."""
    try:
        window.set_focus()
        _fill_focused_field(message)
        time.sleep(0.5)
        send_keys(KEY_ENTER)
        time.sleep(0.5)
    except Exception as exc:
        raise RuntimeError("Could not send message.") from exc


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    try:
        contacts = load_contacts(config.CONTACTS_FILE)
        if not contacts:
            logger.error("No contacts found in %s – aborting.", config.CONTACTS_FILE)
            sys.exit(1)

        logger.info("Loaded %d contacts.", len(contacts))

        app = open_phone_link()
        window = get_main_window(app)
        bring_window_to_front(window)
        time.sleep(1)

        navigate_to_messages(window)

        success_count = 0
        failure_count = 0

        for contact in contacts:
            logger.info("Sending to %s (%s)...", contact["name"], contact["phone"])
            try:
                start_new_conversation(window)
                if not search_and_select_contact(window, contact):
                    logger.warning("Skipping %s -- contact not found.", contact["name"])
                    failure_count += 1
                    send_keys(KEY_ESCAPE)
                    time.sleep(1)
                    continue
                attach_image(window, getattr(config, "IMAGE_FILE", ""))
                type_and_send_message(window, config.MESSAGE)
                logger.info("✓ Sent to %s", contact["name"])
                success_count += 1
            except Exception as exc:
                logger.error("Failed to send to %s: %s", contact["name"], exc)
                failure_count += 1
                try:
                    send_keys(KEY_ESCAPE)
                except Exception:
                    pass

            time.sleep(config.DELAY_BETWEEN_MESSAGES)

        logger.info(
            "Done. Sent: %d  |  Failed: %d  |  Total: %d",
            success_count,
            failure_count,
            len(contacts),
        )
    except Exception as exc:
        logger.exception("Fatal error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
