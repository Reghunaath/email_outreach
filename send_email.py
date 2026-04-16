import win32com.client
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────
TEMPLATE_FILE  = "template.html"
RESUME_FILE    = "Reghunaath_Resume_Feb_N.pdf"

RECIPIENTS = [
    {"name": "Sharon", "email": "sharon@example.com", "company": "CREWASIS"},
    # {"name": "John",   "email": "john@example.com",   "company": "Acme Corp"},
]
# ──────────────────────────────────────────────────────────────────────────────


def load_template(name: str, company: str) -> str:
    template = Path(TEMPLATE_FILE).read_text(encoding="utf-8")
    return template.format(name=name, company=company)


def send_email(name: str, email: str, company: str) -> None:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = MailItem

    mail.To = email
    mail.Subject = f"How I Can Contribute to {company}"
    mail.HTMLBody = load_template(name, company)
    mail.Attachments.Add(str(Path(RESUME_FILE).resolve()))

    mail.Send()
    print(f"Sent to {name} at {company} ({email})")


if __name__ == "__main__":
    for recipient in RECIPIENTS:
        send_email(
            name=recipient["name"],
            email=recipient["email"],
            company=recipient["company"],
        )
