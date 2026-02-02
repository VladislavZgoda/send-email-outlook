from json import load
from pathlib import Path
from typing import TypedDict

import win32com.client as win32

Email = TypedDict("Email", {"email": str, "files": list[str]})

try:
    with open("emails.json", mode="r", encoding="UTF-8") as f:
        emails: list[Email] = load(f)
except FileNotFoundError:
    print("Error! File 'emails.json' not found!")


def main(emails: list[Email]) -> None:
    script_dir = Path(__file__).resolve().parent
    outlook = win32.Dispatch("outlook.application")
    subject = "АСКУЭ Тимашевск"
    attachment_path = script_dir / "assets"

    for email in emails:
        mail = outlook.CreateItem(0)
        mail.To = email["email"]
        mail.Subject = subject
        mail.Body = ""

        for f in email["files"]:
            attachment = str(attachment_path / f)
            mail.Attachments.Add(attachment)

        mail.Send()
    print("OK")


if __name__ == "__main__":
    main(emails)
