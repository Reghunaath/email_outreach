import json
import time
import threading
from pathlib import Path
import requests
import customtkinter as ctk

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

BASE_URL   = "https://app.apollo.io"
CREDS_FILE = Path(__file__).parent / "apollo_creds.json"

_HEADERS_BASE = {
    "Content-Type": "application/json",
    "x-referer-host": "app.apollo.io",
    "x-referer-path": "/people",
    "x-accept-language": "en",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/148.0.0.0 Safari/537.36"
    ),
}


def _parse_cookies(cookie_str: str) -> dict:
    cookies: dict = {}
    for part in cookie_str.split(";"):
        part = part.strip()
        if "=" in part:
            k, v = part.split("=", 1)
            cookies[k.strip()] = v.strip()
    return cookies


def _find_person(session: requests.Session, headers: dict, linkedin_url: str) -> dict | None:
    resp = session.post(
        f"{BASE_URL}/api/v1/mixed_people/search",
        headers=headers,
        json={"person_linkedin_urls": [linkedin_url], "page": 1, "per_page": 1},
    )
    resp.raise_for_status()
    people = resp.json().get("people", [])
    if people:
        return people[0]

    resp2 = session.post(
        f"{BASE_URL}/api/v1/contacts/search",
        headers=headers,
        json={"person_linkedin_urls": [linkedin_url], "page": 1, "per_page": 1},
    )
    resp2.raise_for_status()
    contacts = resp2.json().get("contacts", [])
    if contacts:
        c = contacts[0]
        return {
            "id": c.get("person_id"),
            "first_name": c.get("first_name"),
            "organization": {"name": c.get("organization_name", "")},
            "_contact": c,
        }
    return None


def _reveal_email(session: requests.Session, headers: dict, person_id: str) -> dict:
    resp = session.post(
        f"{BASE_URL}/api/v1/mixed_people/add_to_my_prospects",
        headers=headers,
        json={
            "entity_ids": [person_id],
            "analytics_context": "Searcher: Individual Add Button",
            "skip_fetching_people": False,
            "cta_name": "Access email",
            "cacheKey": int(time.time() * 1000),
        },
    )
    resp.raise_for_status()
    contacts = resp.json().get("contacts", [])
    return contacts[0] if contacts else {}


class ApolloFinderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Apollo Email Finder")
        self.geometry("580x560")
        self.resizable(False, False)

        self._build_credentials_page()
        self._build_lookup_page()
        self._load_creds()
        self._show_credentials_page()

    # ── Page builders ──────────────────────────────────────────────────────────

    def _build_credentials_page(self):
        self.creds_frame = ctk.CTkFrame(self, fg_color="transparent")

        ctk.CTkLabel(
            self.creds_frame,
            text="Apollo Credentials",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(pady=(28, 4))
        ctk.CTkLabel(
            self.creds_frame,
            text="Paste your Apollo session credentials from DevTools.",
            text_color="gray",
        ).pack(pady=(0, 20))

        ctk.CTkLabel(self.creds_frame, text="CSRF Token", anchor="w").pack(fill="x", padx=36)
        self.csrf_entry = ctk.CTkEntry(
            self.creds_frame, width=508, height=36, placeholder_text="x-csrf-token header value"
        )
        self.csrf_entry.pack(padx=36, pady=(4, 18))

        ctk.CTkLabel(self.creds_frame, text="Cookies", anchor="w").pack(fill="x", padx=36)
        self.cookies_box = ctk.CTkTextbox(self.creds_frame, width=508, height=230, wrap="none")
        self.cookies_box.pack(padx=36, pady=(4, 8))

        self.creds_error_label = ctk.CTkLabel(
            self.creds_frame, text="", text_color="red", wraplength=508
        )
        self.creds_error_label.pack(padx=36, pady=(0, 8))

        ctk.CTkButton(
            self.creds_frame, text="Next →", width=508, height=38, command=self._go_to_lookup
        ).pack(padx=36, pady=(0, 28))

    def _build_lookup_page(self):
        self.lookup_frame = ctk.CTkFrame(self, fg_color="transparent")

        ctk.CTkLabel(
            self.lookup_frame,
            text="LinkedIn Profile URLs",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(pady=(28, 4))
        ctk.CTkLabel(
            self.lookup_frame,
            text="One LinkedIn profile URL per line.",
            text_color="gray",
        ).pack(pady=(0, 12))

        self.urls_box = ctk.CTkTextbox(self.lookup_frame, width=508, height=310, wrap="none")
        self.urls_box.pack(padx=36, pady=(0, 12))

        self.status_label = ctk.CTkLabel(
            self.lookup_frame, text="", text_color="gray", wraplength=508
        )
        self.status_label.pack(padx=36, pady=(0, 12))

        btn_row = ctk.CTkFrame(self.lookup_frame, fg_color="transparent")
        btn_row.pack(padx=36, fill="x", pady=(0, 28))

        ctk.CTkButton(
            btn_row,
            text="← Back",
            width=110,
            height=38,
            fg_color="gray30",
            hover_color="gray40",
            command=self._show_credentials_page,
        ).pack(side="left")

        self.fetch_btn = ctk.CTkButton(
            btn_row, text="Fetch & Copy", height=38, command=self._start_fetch
        )
        self.fetch_btn.pack(side="right", fill="x", expand=True, padx=(12, 0))

    # ── Credential persistence ─────────────────────────────────────────────────

    def _load_creds(self) -> None:
        if not CREDS_FILE.exists():
            return
        try:
            data = json.loads(CREDS_FILE.read_text(encoding="utf-8"))
            self.csrf_entry.insert(0, data.get("csrf_token", ""))
            self.cookies_box.insert("1.0", data.get("cookies", ""))
        except Exception:
            pass

    def _save_creds(self) -> None:
        data = {
            "csrf_token": self.csrf_entry.get().strip(),
            "cookies": self.cookies_box.get("1.0", "end").strip(),
        }
        CREDS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")

    # ── Navigation ─────────────────────────────────────────────────────────────

    def _show_credentials_page(self):
        self.lookup_frame.pack_forget()
        self.creds_frame.pack(fill="both", expand=True)

    def _go_to_lookup(self):
        csrf = self.csrf_entry.get().strip()
        cookies = self.cookies_box.get("1.0", "end").strip()
        if not csrf or not cookies:
            self.creds_error_label.configure(text="Both CSRF token and cookies are required.")
            return
        self.creds_error_label.configure(text="")
        self._save_creds()
        self._show_lookup_page()

    def _show_lookup_page(self):
        self.creds_frame.pack_forget()
        self.lookup_frame.pack(fill="both", expand=True)

    # ── Fetch logic ────────────────────────────────────────────────────────────

    def _start_fetch(self):
        urls_raw = self.urls_box.get("1.0", "end").strip()
        urls = [u.strip().rstrip("/") for u in urls_raw.splitlines() if u.strip()]
        if not urls:
            self.status_label.configure(text="Enter at least one LinkedIn URL.", text_color="orange")
            return

        self.fetch_btn.configure(state="disabled", text="Fetching...")
        self.status_label.configure(text=f"Processing 0 / {len(urls)}...", text_color="gray")

        csrf = self.csrf_entry.get().strip()
        cookies_str = self.cookies_box.get("1.0", "end").strip()

        threading.Thread(
            target=self._fetch_all,
            args=(urls, csrf, cookies_str),
            daemon=True,
        ).start()

    def _fetch_all(self, urls: list[str], csrf: str, cookies_str: str) -> None:
        headers = {**_HEADERS_BASE, "x-csrf-token": csrf}
        session = requests.Session()
        session.cookies.update(_parse_cookies(cookies_str))

        names: list[str] = []
        emails: list[str] = []
        errors: list[str] = []
        total = len(urls)

        for i, url in enumerate(urls):
            self.after(0, lambda i=i: self.status_label.configure(
                text=f"Processing {i + 1} / {total}...", text_color="gray"
            ))
            try:
                person = _find_person(session, headers, url)
                if not person:
                    errors.append(url)
                    continue
                contact = person.get("_contact") or _reveal_email(session, headers, person["id"])
                first_name = (contact.get("first_name") or person.get("first_name", "")).strip().title()
                email = contact.get("email", "")
                email_status = contact.get("email_status", "")
                if first_name:
                    names.append(first_name)
                if email and email_status != "unavailable":
                    emails.append(email)
            except Exception as exc:
                errors.append(f"{url} ({exc})")

        self.after(0, lambda: self._on_fetch_done(names, emails, errors, total))

    def _on_fetch_done(
        self, names: list[str], emails: list[str], errors: list[str], total: int
    ) -> None:
        self.fetch_btn.configure(state="normal", text="Fetch & Copy")

        if not names and not emails:
            first_err = f" — {errors[0]}" if errors else ""
            self.status_label.configure(
                text=f"No results found.{first_err}", text_color="red"
            )
            return

        names_str = ", ".join(names)
        emails_str = ", ".join(emails)
        self.clipboard_clear()
        self.clipboard_append(f"{names_str}\t{emails_str}")

        error_note = f"  ({len(errors)} failed)" if errors else ""
        self.status_label.configure(
            text=f"Copied {len(names)} / {total} to clipboard.{error_note}",
            text_color="green",
        )


if __name__ == "__main__":
    app = ApolloFinderApp()
    app.mainloop()
