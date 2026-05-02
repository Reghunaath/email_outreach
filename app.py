import customtkinter as ctk
import win32com.client
import json
import csv
from pathlib import Path
from datetime import datetime, timedelta

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

TEMPLATE_FILE           = "template.html"
TEMPLATE_RECRUITER_FILE = "template_recruiter.html"
RESUME_FILE             = "Reghunaath_Resume_Feb_N.pdf"
DATA_FILE               = "data.json"
LOG_FILE                = "log.csv"

DEFAULT_SUBJECT_FOUNDER   = "How I Can Contribute to {company}"
DEFAULT_SUBJECT_RECRUITER = "Reaching out so I'm more than just a PDF"

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
            "url_extension": url_extension,
            "sent_at":       datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "scheduled_for": scheduled_for,
        })


def read_data() -> dict:
    return json.loads(Path(DATA_FILE).read_text(encoding="utf-8"))


def increment_url_extension() -> None:
    data = read_data()
    data["url_extension"] += 1
    Path(DATA_FILE).write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")


def read_template_raw(mode: str = "Founder") -> str:
    file = TEMPLATE_RECRUITER_FILE if mode == "Recruiter" else TEMPLATE_FILE
    return Path(file).read_text(encoding="utf-8").strip()


class BodyEditModal(ctk.CTkToplevel):
    def __init__(self, parent, initial_text: str):
        super().__init__(parent)
        self.title("Edit Email Body")
        self.geometry("560x500")
        self.resizable(False, False)
        self.grab_set()

        self._saved_text = initial_text

        self.textbox = ctk.CTkTextbox(self, width=512, height=400, wrap="word")
        self.textbox.insert("1.0", initial_text)
        self.textbox.pack(padx=24, pady=(20, 12))

        ctk.CTkButton(
            self, text="Save", width=512, height=38, command=self._save
        ).pack(padx=24, pady=(0, 20))

    def _save(self):
        self._saved_text = self.textbox.get("1.0", "end").strip()
        self.destroy()

    def get_text(self) -> str:
        return self._saved_text


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Email Outreach")
        self.geometry("420x672")
        self.resizable(False, False)
        self._body_text = read_template_raw("Founder")
        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self, text="Email Outreach", font=ctk.CTkFont(size=20, weight="bold")
        ).pack(padx=24, pady=(24, 20), anchor="w")

        # ── Mode ──
        ctk.CTkLabel(
            self, text="MODE", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(0, 8), anchor="w")

        self.mode_toggle = ctk.CTkSegmentedButton(
            self, values=["Founder", "Recruiter"], command=self._toggle_mode, width=372
        )
        self.mode_toggle.set("Founder")
        self.mode_toggle.pack(padx=24, pady=(0, 10))

        # ── Recipient ──
        ctk.CTkLabel(
            self, text="RECIPIENT", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(8, 8), anchor="w")

        self.name_entry    = ctk.CTkEntry(self, placeholder_text="Name",    width=372, height=38)
        self.email_entry   = ctk.CTkEntry(self, placeholder_text="Email",   width=372, height=38)
        self.company_entry = ctk.CTkEntry(self, placeholder_text="Company", width=372, height=38)
        for entry in (self.name_entry, self.email_entry, self.company_entry):
            entry.pack(padx=24, pady=(0, 10))

        # Job IDs entry — hidden until Recruiter mode is selected
        self.job_ids_entry = ctk.CTkEntry(
            self, placeholder_text="Job IDs (e.g. 12345, 67890)", width=372, height=38
        )

        # Subject — always visible, pre-filled with mode default
        self.subject_entry = ctk.CTkEntry(self, placeholder_text="Subject", width=372, height=38)
        self.subject_entry.insert(0, DEFAULT_SUBJECT_FOUNDER)
        self.subject_entry.pack(padx=24, pady=(0, 10))

        # ── Edit Body button ──
        self.edit_body_btn = ctk.CTkButton(
            self, text="Edit Body", width=372, height=38,
            fg_color="transparent", border_width=1,
            command=self._open_body_modal,
        )
        self.edit_body_btn.pack(padx=24, pady=(0, 10))

        # ── When ──
        self.when_label = ctk.CTkLabel(
            self, text="WHEN", font=ctk.CTkFont(size=11), text_color="gray"
        )
        self.when_label.pack(padx=24, pady=(8, 8), anchor="w")

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

    def _open_body_modal(self):
        modal = BodyEditModal(self, self._body_text)
        self.wait_window(modal)
        self._body_text = modal.get_text()

    def _toggle_mode(self, value: str):
        if value == "Recruiter":
            self.job_ids_entry.pack(padx=24, pady=(0, 10), before=self.subject_entry)
            self.subject_entry.delete(0, "end")
            self.subject_entry.insert(0, DEFAULT_SUBJECT_RECRUITER)
        else:
            self.job_ids_entry.pack_forget()
            self.subject_entry.delete(0, "end")
            self.subject_entry.insert(0, DEFAULT_SUBJECT_FOUNDER)
        self._body_text = read_template_raw(value)
        self._update_geometry()

    def _toggle_schedule(self, value: str):
        if value == "Schedule":
            self.schedule_frame.pack(padx=24, pady=(0, 10), before=self.send_btn)
            self._set_schedule_defaults()
        else:
            self.schedule_frame.pack_forget()
        self._update_geometry()

    def _update_geometry(self):
        h = 672
        if self.mode_toggle.get() == "Recruiter":
            h += 48
        if self.send_mode.get() == "Schedule":
            h += 60
        self.geometry(f"420x{h}")

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
        names   = [n.strip() for n in self.name_entry.get().split(",") if n.strip()]
        emails  = [e.strip() for e in self.email_entry.get().split(",") if e.strip()]
        company = self.company_entry.get().strip()
        mode    = self.mode_toggle.get()
        job_ids = self.job_ids_entry.get().strip() if mode == "Recruiter" else ""

        if not names or not emails or not company:
            self._set_status("Fill in all fields.", ok=False)
            return

        subject = self.subject_entry.get().strip().replace("{company}", company)

        if mode == "Recruiter" and not job_ids:
            self._set_status("Enter at least one Job ID.", ok=False)
            return

        if not subject:
            self._set_status("Subject cannot be empty.", ok=False)
            return

        if len(names) != len(emails):
            self._set_status(f"Name/email count mismatch ({len(names)} vs {len(emails)}).", ok=False)
            return

        schedule_mode = self.send_mode.get() == "Schedule"
        dt = None
        if schedule_mode:
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

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            resume_path = str(Path(RESUME_FILE).resolve())

            body_template = self._body_text

            for name, email in zip(names, emails):
                url_extension = read_data()["url_extension"]
                body = (body_template
                    .replace("{name}", name)
                    .replace("{company}", company)
                    .replace("{url_extension}", str(url_extension))
                    .replace("{job_ids}", job_ids))
                mail = outlook.CreateItem(0)
                mail.To       = email
                mail.Subject  = subject
                mail.HTMLBody = body
                mail.Attachments.Add(resume_path)

                if schedule_mode:
                    mail.DeferredDeliveryTime = dt.strftime("%m/%d/%Y %I:%M %p")
                    mail.Send()
                    log_email(name, email, company, url_extension, scheduled_for=dt.strftime("%Y-%m-%d %H:%M:%S"))
                else:
                    mail.Send()
                    log_email(name, email, company, url_extension)
                increment_url_extension()

            self.name_entry.delete(0, "end")
            self.email_entry.delete(0, "end")
            self.company_entry.delete(0, "end")
            if mode == "Recruiter":
                self.job_ids_entry.delete(0, "end")
            self.subject_entry.delete(0, "end")
            self.subject_entry.insert(0, DEFAULT_SUBJECT_FOUNDER if mode == "Founder" else DEFAULT_SUBJECT_RECRUITER)
            self._body_text = read_template_raw(mode)

            count = len(names)
            if schedule_mode:
                self._set_status(f"Scheduled {count} email(s) for {dt.strftime('%b %d at %I:%M %p')}.")
            else:
                label = names[0] if count == 1 else f"{count} recipients"
                self._set_status(f"Sent to {label} at {company}.")

        except Exception as e:
            self._set_status(f"Error: {e}", ok=False)


if __name__ == "__main__":
    app = App()
    app.mainloop()
