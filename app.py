import customtkinter as ctk
import win32com.client
import json
import csv
from pathlib import Path
from datetime import datetime, timedelta

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

TEMPLATE_FILE = "template.html"
RESUME_FILE   = "Reghunaath_Resume_Feb_N.pdf"
DATA_FILE     = "data.json"
LOG_FILE      = "log.csv"

LOG_HEADERS = ["name", "email", "company", "url_extension", "sent_at", "scheduled_for"]


def log_email(name: str, email: str, company: str, url_extension: int, scheduled_for: str = "") -> None:
    log_path = Path(LOG_FILE)
    write_header = not log_path.exists()
    with log_path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=LOG_HEADERS)
        if write_header:
            writer.writeheader()
        writer.writerow({
            "name":          name,
            "email":         email,
            "company":       company,
            "url_extension":         url_extension,
            "sent_at":       datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "scheduled_for": scheduled_for,
        })


def read_data() -> dict:
    return json.loads(Path(DATA_FILE).read_text(encoding="utf-8"))


def increment_url_extension() -> None:
    data = read_data()
    data["url_extension"] += 1
    Path(DATA_FILE).write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")


def load_template(name: str, company: str) -> str:
    url_extension = read_data()["url_extension"]
    return Path(TEMPLATE_FILE).read_text(encoding="utf-8").format(name=name, company=company, url_extension=url_extension)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Email Outreach")
        self.geometry("420x500")
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self, text="Email Outreach", font=ctk.CTkFont(size=20, weight="bold")
        ).pack(padx=24, pady=(24, 20), anchor="w")

        # ── Recipient ──
        ctk.CTkLabel(
            self, text="RECIPIENT", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(0, 8), anchor="w")

        self.name_entry    = ctk.CTkEntry(self, placeholder_text="Name",    width=372, height=38)
        self.email_entry   = ctk.CTkEntry(self, placeholder_text="Email",   width=372, height=38)
        self.company_entry = ctk.CTkEntry(self, placeholder_text="Company", width=372, height=38)
        for entry in (self.name_entry, self.email_entry, self.company_entry):
            entry.pack(padx=24, pady=(0, 10))

        # ── When ──
        ctk.CTkLabel(
            self, text="WHEN", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(8, 8), anchor="w")

        self.send_mode = ctk.CTkSegmentedButton(
            self, values=["Send Now", "Schedule"], command=self._toggle_schedule, width=372
        )
        self.send_mode.set("Send Now")
        self.send_mode.pack(padx=24, pady=(0, 10))

        # Schedule date/time row (hidden initially)
        self.schedule_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.date_entry = ctk.CTkEntry(
            self.schedule_frame, placeholder_text="MM/DD/YYYY", width=178, height=38
        )
        self.date_entry.pack(side="left", padx=(0, 16))
        self.time_entry = ctk.CTkEntry(
            self.schedule_frame, placeholder_text="HH:MM AM/PM", width=178, height=38
        )
        self.time_entry.pack(side="left")

        # ── Send button ──
        self.send_btn = ctk.CTkButton(
            self, text="Send", width=372, height=42, command=self._send
        )
        self.send_btn.pack(padx=24, pady=(16, 10))

        # ── Status ──
        self.status = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12))
        self.status.pack(padx=24)

    def _toggle_schedule(self, value: str):
        if value == "Schedule":
            self.schedule_frame.pack(padx=24, pady=(0, 10), before=self.send_btn)
            self.geometry("420x560")
            self._set_schedule_defaults()
        else:
            self.schedule_frame.pack_forget()
            self.geometry("420x500")

    def _set_schedule_defaults(self):
        now = datetime.now()
        default_date = now.date()
        if now.hour >= 15:  # after 3 PM → next day
            default_date = default_date + timedelta(days=1)
        self.date_entry.delete(0, "end")
        self.date_entry.insert(0, default_date.strftime("%m/%d/%Y"))
        self.time_entry.delete(0, "end")
        self.time_entry.insert(0, "10:10 AM")

    def _set_status(self, msg: str, ok: bool = True):
        self.status.configure(text=msg, text_color=("green" if ok else "red"))

    def _send(self):
        name    = self.name_entry.get().strip()
        email   = self.email_entry.get().strip()
        company = self.company_entry.get().strip()

        if not all([name, email, company]):
            self._set_status("Fill in all fields.", ok=False)
            return

        try:
            url_extension   = read_data()["url_extension"]
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail    = outlook.CreateItem(0)
            mail.To       = email
            mail.Subject  = f"How I Can Contribute to {company}"
            mail.HTMLBody = load_template(name, company)
            mail.Attachments.Add(str(Path(RESUME_FILE).resolve()))

            if self.send_mode.get() == "Schedule":
                date_str = self.date_entry.get().strip()
                time_str = self.time_entry.get().strip()
                if not date_str or not time_str:
                    self._set_status("Enter date and time.", ok=False)
                    return
                try:
                    dt = datetime.strptime(f"{date_str} {time_str}", "%m/%d/%Y %I:%M %p")
                except ValueError:
                    self._set_status("Use MM/DD/YYYY and HH:MM AM/PM.", ok=False)
                    return
                mail.DeferredDeliveryTime = dt.strftime("%m/%d/%Y %I:%M %p")
                mail.Send()
                increment_url_extension()
                log_email(name, email, company, url_extension, scheduled_for=dt.strftime("%Y-%m-%d %H:%M:%S"))
                self._set_status(f"Scheduled for {dt.strftime('%b %d at %I:%M %p')}.")
            else:
                mail.Send()
                increment_url_extension()
                log_email(name, email, company, url_extension)
                self._set_status(f"Sent to {name} at {company}.")

            self.name_entry.delete(0, "end")
            self.email_entry.delete(0, "end")
            self.company_entry.delete(0, "end")

        except Exception as e:
            self._set_status(f"Error: {e}", ok=False)


if __name__ == "__main__":
    app = App()
    app.mainloop()
