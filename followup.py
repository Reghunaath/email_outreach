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


def parse_company(body: str) -> str:
    for pattern in [
        r"I stumbled upon (.+?) recently",          # Founder
        r"I'm genuinely excited about (.+?) and",   # Recruiter / Hiring Manager
    ]:
        match = re.search(pattern, body)
        if match:
            return match.group(1).strip()
    return "N/A"


class FollowUpApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Follow-up")
        self.geometry("590x690")
        self.resizable(False, False)
        self._outlook    = None
        self._mail_items = []  # list of (BooleanVar, outlook_mail_item)
        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self, text="Follow-up", font=ctk.CTkFont(size=20, weight="bold")
        ).pack(padx=24, pady=(24, 20), anchor="w")

        # ── Mode ──
        ctk.CTkLabel(
            self, text="MODE", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(0, 8), anchor="w")

        self.mode_toggle = ctk.CTkSegmentedButton(
            self, values=["Founder", "Recruiter"], width=542
        )
        self.mode_toggle.set("Recruiter")
        self.mode_toggle.pack(padx=24, pady=(0, 10))

        # ── Sent emails header ──
        ctk.CTkLabel(
            self, text="SENT EMAILS", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(8, 8), anchor="w")

        # ── Date filter + Search ──
        filter_frame = ctk.CTkFrame(self, fg_color="transparent")
        filter_frame.pack(padx=24, pady=(0, 8), fill="x")
        self._filter_from = ctk.CTkEntry(filter_frame, placeholder_text="From MM/DD/YYYY", width=190, height=34)
        self._filter_from.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(filter_frame, text="→", text_color="gray").pack(side="left", padx=(0, 8))
        self._filter_to = ctk.CTkEntry(filter_frame, placeholder_text="To MM/DD/YYYY", width=190, height=34)
        self._filter_to.pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            filter_frame, text="Search", width=110, height=34,
            command=self._load_sent_emails,
        ).pack(side="left")

        today        = datetime.now().date()
        this_monday  = today - timedelta(days=today.weekday())
        past_monday  = this_monday - timedelta(days=7)
        past_sunday  = this_monday - timedelta(days=1)
        self._filter_from.insert(0, past_monday.strftime("%m/%d/%Y"))
        self._filter_to.insert(0, past_sunday.strftime("%m/%d/%Y"))

        # ── Scrollable email list ──
        self.scroll_frame = ctk.CTkScrollableFrame(self, width=526, height=260)
        self.scroll_frame.pack(padx=24, pady=(0, 4))

        self._count_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=11), text_color="gray")
        self._count_label.pack(padx=24, pady=(0, 8), anchor="e")

        # ── When ──
        ctk.CTkLabel(
            self, text="WHEN", font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=24, pady=(8, 8), anchor="w")

        self.send_mode = ctk.CTkSegmentedButton(
            self, values=["Send Now", "Schedule"], command=self._toggle_schedule, width=542
        )
        self.send_mode.set("Send Now")
        self.send_mode.pack(padx=24, pady=(0, 10))

        self.schedule_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.date_entry = ctk.CTkEntry(
            self.schedule_frame, placeholder_text="MM/DD/YYYY", width=263, height=38
        )
        self.date_entry.pack(side="left", padx=(0, 16))
        self.time_entry = ctk.CTkEntry(
            self.schedule_frame, placeholder_text="HH:MM AM/PM", width=263, height=38
        )
        self.time_entry.pack(side="left")

        # ── Send button ──
        self.send_btn = ctk.CTkButton(
            self, text="Send", width=542, height=42, command=self._send
        )
        self.send_btn.pack(padx=24, pady=(16, 10))

        # ── Status ──
        self.status = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12))
        self.status.pack(padx=24)

    def _get_replied_conv_ids(self, items: list, filter_from) -> set:
        replied = set()
        if not items or filter_from is None:
            return replied
        try:
            namespace = self._outlook.GetNamespace("MAPI")
            inbox     = namespace.GetDefaultFolder(6)  # olFolderInbox

            target_conv_ids = {item.ConversationID for item in items}

            inbox_items = inbox.Items
            inbox_items.Sort("[ReceivedTime]", True)  # descending
            today = datetime.now().date()
            for inbox_item in inbox_items:
                try:
                    received_date = inbox_item.ReceivedTime.date()
                    if received_date > today:
                        continue
                    if received_date < filter_from:
                        break  # sorted descending; nothing older will match
                    if inbox_item.ConversationID in target_conv_ids:
                        replied.add(inbox_item.ConversationID)
                except Exception:
                    continue
        except Exception:
            pass
        return replied

    def _load_sent_emails(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self._mail_items.clear()

        filter_from = filter_to = None
        try:
            from_str = self._filter_from.get().strip()
            to_str   = self._filter_to.get().strip()
            if from_str:
                filter_from = datetime.strptime(from_str, "%m/%d/%Y").date()
            if to_str:
                filter_to = datetime.strptime(to_str, "%m/%d/%Y").date()
        except ValueError:
            pass

        try:
            if self._outlook is None:
                self._outlook = win32com.client.Dispatch("Outlook.Application")
            namespace   = self._outlook.GetNamespace("MAPI")
            sent_folder = namespace.GetDefaultFolder(5)  # olFolderSentMail
            items       = sent_folder.Items
            items.Sort("[SentOn]", True)

            collected = []
            count = 0
            for item in items:
                if count >= MAX_SENT_EMAILS:
                    break
                try:
                    sent_date = item.SentOn.date()
                    if filter_from and sent_date < filter_from:
                        break  # items are sorted descending; nothing older will match
                    if filter_to and sent_date > filter_to:
                        continue
                    collected.append(item)
                    count += 1
                except Exception:
                    continue

            replied_conv_ids = self._get_replied_conv_ids(collected, filter_from)

            for item in collected:
                try:
                    subject  = item.Subject or "(no subject)"
                    date_str = item.SentOn.strftime("%b %d")
                    body     = item.Body or ""
                    name     = parse_first_name(body)
                    company  = parse_company(body)
                    truncated_subject = subject[:55] + "…" if len(subject) > 55 else subject
                    has_reply = item.ConversationID in replied_conv_ids
                    reply_tag = "  ↩" if has_reply else ""
                    label     = f"{name} ({company})  ·  {truncated_subject}  ·  {date_str}{reply_tag}"

                    var = ctk.BooleanVar(value=False)
                    cb  = ctk.CTkCheckBox(
                        self.scroll_frame, text=label,
                        variable=var, onvalue=True, offvalue=False,
                        width=510,
                    )
                    if has_reply:
                        cb.configure(text_color="gray")
                    cb.pack(anchor="w", pady=3)
                    self._mail_items.append((var, item))
                except Exception:
                    continue

            if not collected:
                ctk.CTkLabel(
                    self.scroll_frame, text="No sent emails found.", text_color="gray"
                ).pack(pady=20)
                self._count_label.configure(text="")
            else:
                n = len(collected)
                self._count_label.configure(text=f"{n} email{'s' if n != 1 else ''} found")

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
        h = 690
        if self.send_mode.get() == "Schedule":
            h += 60
        self.geometry(f"590x{h}")

    def _set_schedule_defaults(self):
        now          = datetime.now()
        default_date = now.date()
        if now.hour >= 15:
            default_date = default_date + timedelta(days=1)
        self.date_entry.delete(0, "end")
        self.date_entry.insert(0, default_date.strftime("%m/%d/%Y"))
        self.time_entry.delete(0, "end")
        self.time_entry.insert(0, "9:10 AM")

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
