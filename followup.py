import re
import customtkinter as ctk
import win32com.client
from pathlib import Path
from datetime import datetime, timedelta

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

TEMPLATE_FOLLOWUP_FOUNDER   = "template_followup_founder.html"
TEMPLATE_FOLLOWUP_RECRUITER = "template_followup_recruiter.html"
MAX_SENT_EMAILS             = 100


def read_template(mode: str) -> str:
    file = TEMPLATE_FOLLOWUP_RECRUITER if mode == "Recruiter" else TEMPLATE_FOLLOWUP_FOUNDER
    return Path(file).read_text(encoding="utf-8").strip()


def parse_first_name(body: str) -> str:
    match = re.search(r"Hi ([^,]+),", body)
    if match:
        return match.group(1).strip()
    return "there"


class FollowUpApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Follow-up")
        self.geometry("420x640")
        self.resizable(False, False)
        self._outlook    = None
        self._mail_items = []  # list of (BooleanVar, outlook_mail_item)
        self._build_ui()
        self._load_sent_emails()

    def _build_ui(self):
        ctk.CTkLabel(
            self, text="Follow-up", font=ctk.CTkFont(size=20, weight="bold")
        ).pack(padx=24, pady=(24, 20), anchor="w")

        # ── Mode ──
        ctk.CTkLabel(
            self, text="MODE", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(0, 8), anchor="w")

        self.mode_toggle = ctk.CTkSegmentedButton(
            self, values=["Founder", "Recruiter"], width=372
        )
        self.mode_toggle.set("Recruiter")
        self.mode_toggle.pack(padx=24, pady=(0, 10))

        # ── Sent emails header ──
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(padx=24, pady=(8, 8), fill="x")
        ctk.CTkLabel(
            header_frame, text="SENT EMAILS", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(side="left")
        ctk.CTkButton(
            header_frame, text="Refresh", width=72, height=26,
            fg_color="transparent", border_width=1,
            command=self._load_sent_emails,
        ).pack(side="right")

        # ── Scrollable email list ──
        self.scroll_frame = ctk.CTkScrollableFrame(self, width=356, height=260)
        self.scroll_frame.pack(padx=24, pady=(0, 10))

        # ── When ──
        ctk.CTkLabel(
            self, text="WHEN", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(8, 8), anchor="w")

        self.send_mode = ctk.CTkSegmentedButton(
            self, values=["Send Now", "Schedule"], command=self._toggle_schedule, width=372
        )
        self.send_mode.set("Send Now")
        self.send_mode.pack(padx=24, pady=(0, 10))

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

    def _load_sent_emails(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self._mail_items.clear()

        try:
            if self._outlook is None:
                self._outlook = win32com.client.Dispatch("Outlook.Application")
            namespace   = self._outlook.GetNamespace("MAPI")
            sent_folder = namespace.GetDefaultFolder(5)  # olFolderSentMail
            items       = sent_folder.Items
            items.Sort("[SentOn]", True)

            count = 0
            for item in items:
                if count >= MAX_SENT_EMAILS:
                    break
                try:
                    to_field = item.To or ""
                    subject  = item.Subject or "(no subject)"
                    date_str = item.SentOn.strftime("%b %d")
                    name     = parse_first_name(item.Body or "")
                    truncated_subject = subject[:38] + "…" if len(subject) > 38 else subject
                    label    = f"{name}  ·  {truncated_subject}  ·  {date_str}"

                    var = ctk.BooleanVar(value=False)
                    ctk.CTkCheckBox(
                        self.scroll_frame, text=label,
                        variable=var, onvalue=True, offvalue=False,
                        width=340,
                    ).pack(anchor="w", pady=3)
                    self._mail_items.append((var, item))
                    count += 1
                except Exception:
                    continue

            if count == 0:
                ctk.CTkLabel(
                    self.scroll_frame, text="No sent emails found.", text_color="gray"
                ).pack(pady=20)

        except Exception as e:
            self._set_status(f"Could not load emails: {e}", ok=False)

    def _toggle_schedule(self, value: str):
        if value == "Schedule":
            self.schedule_frame.pack(padx=24, pady=(0, 10), before=self.send_btn)
            self._set_schedule_defaults()
        else:
            self.schedule_frame.pack_forget()
        self._update_geometry()

    def _update_geometry(self):
        h = 640
        if self.send_mode.get() == "Schedule":
            h += 60
        self.geometry(f"420x{h}")

    def _set_schedule_defaults(self):
        now          = datetime.now()
        default_date = now.date()
        if now.hour >= 15:
            default_date = default_date + timedelta(days=1)
        self.date_entry.delete(0, "end")
        self.date_entry.insert(0, default_date.strftime("%m/%d/%Y"))
        self.time_entry.delete(0, "end")
        self.time_entry.insert(0, "10:10 AM")

    def _set_status(self, msg: str, ok: bool = True):
        self.status.configure(text=msg, text_color=("green" if ok else "red"))

    def _send(self):
        selected = [(var, item) for var, item in self._mail_items if var.get()]
        if not selected:
            self._set_status("Select at least one email.", ok=False)
            return

        mode     = self.mode_toggle.get()
        template = read_template(mode)

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
            for _, mail_item in selected:
                name  = parse_first_name(mail_item.Body or "")
                body  = template.replace("{name}", name)
                reply = mail_item.Reply()
                reply.To = mail_item.To
                reply.HTMLBody = body
                if schedule_mode:
                    reply.DeferredDeliveryTime = dt.strftime("%m/%d/%Y %I:%M %p")
                reply.Send()

            count = len(selected)
            if schedule_mode:
                self._set_status(f"Scheduled {count} follow-up(s) for {dt.strftime('%b %d at %I:%M %p')}.")
            else:
                self._set_status(f"Sent {count} follow-up(s).")

        except Exception as e:
            self._set_status(f"Error: {e}", ok=False)


if __name__ == "__main__":
    app = FollowUpApp()
    app.mainloop()
