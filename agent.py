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

import pandas as pd
import pyperclip
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

import config

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_contacts(path: str) -> list[dict[str, str]]:
    """Return a list of {'name': str, 'phone': str} dicts from a CSV file."""
    df = pd.read_csv(path, dtype=str)
    required = {"name", "phone"}
    missing = required - set(df.columns.str.lower())
    if missing:
        raise ValueError(f"contacts CSV is missing columns: {missing}")
    df.columns = df.columns.str.lower()
    return df[["name", "phone"]].dropna().to_dict(orient="records")


def open_phone_link() -> "Application":
    """Connect to an already-running Phone Link instance or launch it."""
    try:
        app = Application(backend="uia").connect(title_re=".*Phone Link.*", timeout=5)
        logger.info("Connected to running Phone Link instance.")
    except Exception:
        logger.info("Phone Link not found -- launching it now...")
        app = Application(backend="uia").start(
            r"explorer.exe shell:AppsFolder\Microsoft.YourPhone_8wekyb3d8bbwe!App"
        )
        time.sleep(config.UI_TIMEOUT)
    return app


def get_main_window(app: "Application"):
    """Return the Phone Link main window wrapper."""
    return app.window(title_re=".*Phone Link.*")


def navigate_to_messages(window) -> None:
    """Click the 'Messages' navigation item."""
    try:
        messages_nav = window.child_window(title_re="Messages", control_type="ListItem")
        messages_nav.click_input()
        time.sleep(1)
    except Exception as exc:
        raise RuntimeError("Could not find 'Messages' navigation item.") from exc


def start_new_conversation(window) -> None:
    """Click the 'New message' / compose button."""
    try:
        compose_btn = window.child_window(
            title_re="New message|Nouveau message|Compose|Écrire",
            control_type="Button",
        )
        compose_btn.click_input()
        time.sleep(1)
    except Exception as exc:
        raise RuntimeError("Could not find 'New message' button.") from exc


def search_and_select_contact(window, contact: dict) -> bool:
    """
    Type the contact phone number in the recipient field and select them.

    Returns True on success, False if the contact could not be found.
    """
    try:
        recipient_field = window.child_window(
            title_re="To:|À:|Search|Rechercher",
            control_type="Edit",
        )
        recipient_field.click_input()
        # Use clipboard paste so special characters like '+' are not misinterpreted
        pyperclip.copy(contact["phone"])
        send_keys("^v")
        time.sleep(1)

        # Pick the first suggestion or press Enter to confirm the raw number
        suggestions = window.child_window(control_type="List")
        if suggestions.exists(timeout=2):
            children = suggestions.children()
            if children:
                children[0].click_input()
            else:
                send_keys("{ENTER}")
        else:
            send_keys("{ENTER}")
        time.sleep(0.5)
        return True
    except Exception as exc:
        logger.warning("Could not select contact %s: %s", contact["name"], exc)
        return False


def type_and_send_message(window, message: str) -> None:
    """Paste the message into the compose field and send it."""
    try:
        message_field = window.child_window(
            title_re="Message|Aa|Type a message|Écrivez un message",
            control_type="Edit",
        )
        message_field.click_input()
        # Use clipboard paste to handle special characters reliably
        pyperclip.copy(message)
        send_keys("^v")
        time.sleep(0.5)
        send_keys("{ENTER}")
        time.sleep(0.5)
    except Exception as exc:
        raise RuntimeError("Could not send message.") from exc


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    contacts = load_contacts(config.CONTACTS_FILE)
    if not contacts:
        logger.error("No contacts found in %s – aborting.", config.CONTACTS_FILE)
        sys.exit(1)

    logger.info("Loaded %d contacts.", len(contacts))

    app = open_phone_link()
    window = get_main_window(app)
    window.set_focus()
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
                send_keys("{ESCAPE}")
                time.sleep(1)
                continue
            type_and_send_message(window, config.MESSAGE)
            logger.info("✓ Sent to %s", contact["name"])
            success_count += 1
        except Exception as exc:
            logger.error("Failed to send to %s: %s", contact["name"], exc)
            failure_count += 1
            try:
                send_keys("{ESCAPE}")
            except Exception:
                pass

        time.sleep(config.DELAY_BETWEEN_MESSAGES)

    logger.info(
        "Done. Sent: %d  |  Failed: %d  |  Total: %d",
        success_count,
        failure_count,
        len(contacts),
    )


if __name__ == "__main__":
    main()
