"""
Lead Generation Scraper — Desktop App
أداة توليد العملاء المحتملين

GUI built with CustomTkinter for easy use by non-technical clients.

Features:
- Search Google Maps for businesses by keyword + location
- Scrape business name, phone, email, website, address
- Export to Excel or CSV
- Clean modern UI

Requirements:
    pip install customtkinter requests beautifulsoup4 openpyxl selenium webdriver-manager

Usage:
    python lead_scraper.py
"""

import customtkinter as ctk
import threading
import csv
import json
import os
import re
import time
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox

try:
    import requests
    from bs4 import BeautifulSoup
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ─── Scraping Engine ────────────────────────────

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
}

EMAIL_PATTERN = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}"
)
PHONE_PATTERN = re.compile(
    r"(?:\+?\d{1,3}[\s\-]?)?(?:\(?\d{2,4}\)?[\s\-]?){2,4}\d{3,4}"
)


def extract_emails(text):
    """Extract all email addresses from text."""
    emails = EMAIL_PATTERN.findall(text)
    # Filter out common false positives
    return list(set(
        e for e in emails
        if not any(e.endswith(x) for x in ['.png', '.jpg', '.gif', '.css', '.js'])
    ))


def extract_phones(text):
    """Extract phone numbers from text."""
    phones = PHONE_PATTERN.findall(text)
    return list(set(phones[:3]))  # Max 3 phones


def scrape_website_contact(url, timeout=10):
    """Visit a website and extract contact info."""
    if not url or not url.startswith("http"):
        return {}, []

    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        soup = BeautifulSoup(resp.text, "html.parser")
        text = soup.get_text()

        emails = extract_emails(text)
        phones = extract_phones(text)

        # Try contact page
        contact_links = [
            a["href"] for a in soup.find_all("a", href=True)
            if any(word in a["href"].lower() for word in ["contact", "about", "اتصل"])
        ]

        if contact_links and not emails:
            contact_url = contact_links[0]
            if not contact_url.startswith("http"):
                from urllib.parse import urljoin
                contact_url = urljoin(url, contact_url)
            try:
                resp2 = requests.get(contact_url, headers=HEADERS, timeout=8)
                text2 = BeautifulSoup(resp2.text, "html.parser").get_text()
                emails = extract_emails(text2) or emails
                phones = extract_phones(text2) or phones
            except Exception:
                pass

        return {
            "emails": ", ".join(emails[:2]),
            "phones": ", ".join(phones[:2]),
        }, emails

    except Exception:
        return {}, []


def search_google(keyword, location, num_results=20, log_fn=None):
    """Search Google and extract business listings."""
    query = f"{keyword} {location} contact email phone"
    url = f"https://www.google.com/search?q={requests.utils.quote(query)}&num={num_results}"

    if log_fn:
        log_fn(f"Searching: {keyword} in {location}...")

    results = []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(resp.text, "html.parser")

        for g in soup.select("div.g, div[data-hveid]"):
            title_el = g.select_one("h3")
            link_el = g.select_one("a[href]")
            snippet_el = g.select_one("div.VwiC3b, span.aCOpRe, div[data-sncf]")

            if not title_el or not link_el:
                continue

            href = link_el.get("href", "")
            if not href.startswith("http") or "google.com" in href:
                continue

            title = title_el.get_text(strip=True)
            snippet = snippet_el.get_text(strip=True) if snippet_el else ""

            # Extract phones from snippet
            phones = extract_phones(snippet)

            results.append({
                "name": title,
                "website": href,
                "snippet": snippet[:200],
                "phone": phones[0] if phones else "",
                "email": "",
                "address": "",
                "source": "Google Search",
            })

    except Exception as e:
        if log_fn:
            log_fn(f"Search error: {e}")

    return results


def search_yelp(keyword, location, num_results=20, log_fn=None):
    """Search Yelp for business listings."""
    query = f"{keyword} {location}"
    url = f"https://www.yelp.com/search?find_desc={requests.utils.quote(keyword)}&find_loc={requests.utils.quote(location)}"

    if log_fn:
        log_fn(f"Searching Yelp: {keyword} in {location}...")

    results = []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(resp.text, "html.parser")

        for biz in soup.select("div[data-testid='serp-ia-card'], li.regular-search-result"):
            name_el = biz.select_one("a.css-19v1rkv, h3 a, a[name]")
            phone_el = biz.select_one("p.css-1p9ibgf, p[class*='phone']")
            addr_el = biz.select_one("address, p.css-qyp8bo")

            if not name_el:
                continue

            name = name_el.get_text(strip=True)
            phone = phone_el.get_text(strip=True) if phone_el else ""
            address = addr_el.get_text(strip=True) if addr_el else ""
            href = name_el.get("href", "")
            website = f"https://www.yelp.com{href}" if href.startswith("/") else href

            results.append({
                "name": name,
                "website": website,
                "phone": phone,
                "email": "",
                "address": address,
                "source": "Yelp",
            })

    except Exception as e:
        if log_fn:
            log_fn(f"Yelp error: {e}")

    return results


def enrich_leads(leads, log_fn=None, progress_fn=None, stop_event=None):
    """Visit each website and extract emails/phones."""
    enriched = []
    for i, lead in enumerate(leads):
        if stop_event and stop_event.is_set():
            break

        if progress_fn:
            progress_fn(i + 1, len(leads))

        website = lead.get("website", "")
        if website and website.startswith("http") and "yelp.com" not in website:
            if log_fn:
                log_fn(f"  [{i+1}/{len(leads)}] Extracting from: {website[:50]}...")
            contact, _ = scrape_website_contact(website)
            lead["email"] = contact.get("emails", lead.get("email", ""))
            if not lead.get("phone"):
                lead["phone"] = contact.get("phones", "")
            time.sleep(0.5)

        enriched.append(lead)

    return enriched


# ─── Export ─────────────────────────────────────

def export_excel(leads, filepath):
    """Export leads to a formatted Excel file."""
    if not HAS_OPENPYXL:
        export_csv(leads, filepath.replace(".xlsx", ".csv"))
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads"

    headers = ["#", "Business Name", "Phone", "Email", "Website", "Address", "Source"]
    colors = {
        "header_fill": "1A5276",
        "alt_fill": "EBF5FB",
    }

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor=colors["header_fill"])
    alt_fill = PatternFill("solid", fgColor=colors["alt_fill"])
    center = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side("thin", color="CCCCCC"),
        right=Side("thin", color="CCCCCC"),
        top=Side("thin", color="CCCCCC"),
        bottom=Side("thin", color="CCCCCC"),
    )

    # Header
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    ws.row_dimensions[1].height = 30

    # Data
    for row_idx, lead in enumerate(leads, 2):
        row_data = [
            row_idx - 1,
            lead.get("name", ""),
            lead.get("phone", ""),
            lead.get("email", ""),
            lead.get("website", ""),
            lead.get("address", ""),
            lead.get("source", ""),
        ]
        fill = alt_fill if row_idx % 2 == 0 else None
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if fill:
                cell.fill = fill
            # Make website clickable
            if col_idx == 5 and val and val.startswith("http"):
                cell.hyperlink = val
                cell.font = Font(color="1F618D", underline="single")

    # Column widths
    widths = [5, 35, 18, 35, 40, 30, 15]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Summary row
    ws.append([])
    summary_row = ws.max_row + 1
    ws.cell(row=summary_row, column=1, value=f"Total: {len(leads)} leads")
    ws.cell(row=summary_row, column=1).font = Font(bold=True, color="1A5276")

    # Freeze header
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{len(leads)+1}"

    wb.save(filepath)


def export_csv(leads, filepath):
    """Export leads to CSV."""
    if not leads:
        return
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["name", "phone", "email", "website", "address", "source"])
        writer.writeheader()
        writer.writerows(leads)


# ─── GUI ────────────────────────────────────────

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class LeadScraperApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Lead Generator Pro — أداة توليد العملاء")
        self.geometry("900x700")
        self.minsize(800, 600)
        self.resizable(True, True)

        self.leads = []
        self.is_running = False
        self.stop_event = threading.Event()

        self._build_ui()

    def _build_ui(self):
        """Build the main UI."""

        # ── Header ──
        header = ctk.CTkFrame(self, fg_color="#1A5276", corner_radius=0, height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="🔍  Lead Generator Pro",
            font=ctk.CTkFont(family="Helvetica", size=22, weight="bold"),
            text_color="white"
        ).pack(side="left", padx=24, pady=16)

        ctk.CTkLabel(
            header,
            text="أداة توليد العملاء المحتملين",
            font=ctk.CTkFont(size=13),
            text_color="#AED6F1"
        ).pack(side="right", padx=24)

        # ── Main content ──
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=16)

        # Left panel — Settings
        left = ctk.CTkFrame(main, width=320, corner_radius=12)
        left.pack(side="left", fill="y", padx=(0, 12))
        left.pack_propagate(False)

        ctk.CTkLabel(left, text="Search Settings", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(16, 4), padx=16, anchor="w")
        ctk.CTkLabel(left, text="─" * 36, text_color="#555").pack(padx=16)

        # Keyword
        ctk.CTkLabel(left, text="Business Type / Keyword", font=ctk.CTkFont(size=12)).pack(pady=(12, 2), padx=16, anchor="w")
        self.keyword_entry = ctk.CTkEntry(left, placeholder_text="e.g. restaurants, dentists, lawyers", height=38)
        self.keyword_entry.pack(padx=16, fill="x")

        # Location
        ctk.CTkLabel(left, text="Location / City", font=ctk.CTkFont(size=12)).pack(pady=(10, 2), padx=16, anchor="w")
        self.location_entry = ctk.CTkEntry(left, placeholder_text="e.g. Cairo Egypt, Dubai UAE", height=38)
        self.location_entry.pack(padx=16, fill="x")

        # Number of results
        ctk.CTkLabel(left, text="Number of Results", font=ctk.CTkFont(size=12)).pack(pady=(10, 2), padx=16, anchor="w")
        self.num_var = ctk.StringVar(value="20")
        num_menu = ctk.CTkOptionMenu(left, values=["10", "20", "30", "50"], variable=self.num_var)
        num_menu.pack(padx=16, fill="x")

        # Source
        ctk.CTkLabel(left, text="Search Source", font=ctk.CTkFont(size=12)).pack(pady=(10, 2), padx=16, anchor="w")
        self.source_var = ctk.StringVar(value="Google")
        source_menu = ctk.CTkOptionMenu(left, values=["Google", "Yelp", "Both"], variable=self.source_var)
        source_menu.pack(padx=16, fill="x")

        # Enrich toggle
        self.enrich_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            left,
            text="Extract emails from websites\n(slower but more data)",
            variable=self.enrich_var,
            font=ctk.CTkFont(size=12),
        ).pack(pady=(12, 4), padx=16, anchor="w")

        # ── Separator ──
        ctk.CTkLabel(left, text="─" * 36, text_color="#555").pack(padx=16, pady=4)

        # Export format
        ctk.CTkLabel(left, text="Export Format", font=ctk.CTkFont(size=12)).pack(pady=(4, 2), padx=16, anchor="w")
        self.export_var = ctk.StringVar(value="Excel (.xlsx)")
        ctk.CTkOptionMenu(left, values=["Excel (.xlsx)", "CSV (.csv)"], variable=self.export_var).pack(padx=16, fill="x")

        # Buttons
        self.start_btn = ctk.CTkButton(
            left,
            text="▶  Start Scraping",
            command=self._start_scraping,
            height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#1A8F4A",
            hover_color="#157A3E",
        )
        self.start_btn.pack(padx=16, pady=(16, 6), fill="x")

        self.stop_btn = ctk.CTkButton(
            left,
            text="⏹  Stop",
            command=self._stop_scraping,
            height=36,
            font=ctk.CTkFont(size=13),
            fg_color="#922B21",
            hover_color="#7B241C",
            state="disabled",
        )
        self.stop_btn.pack(padx=16, pady=(0, 6), fill="x")

        self.export_btn = ctk.CTkButton(
            left,
            text="💾  Export Results",
            command=self._export_results,
            height=36,
            font=ctk.CTkFont(size=13),
            fg_color="#1A5276",
            hover_color="#154360",
            state="disabled",
        )
        self.export_btn.pack(padx=16, pady=(0, 16), fill="x")

        # ── Stats ──
        stats_frame = ctk.CTkFrame(left, fg_color="#1A3A5C", corner_radius=8)
        stats_frame.pack(padx=16, pady=(0, 16), fill="x")

        self.leads_count_label = ctk.CTkLabel(
            stats_frame, text="0 leads found",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#AED6F1"
        )
        self.leads_count_label.pack(pady=(12, 2))

        self.email_count_label = ctk.CTkLabel(
            stats_frame, text="0 with email",
            font=ctk.CTkFont(size=12),
            text_color="#7FB3D3"
        )
        self.email_count_label.pack(pady=(0, 12))

        # ── Right panel ──
        right = ctk.CTkFrame(main, corner_radius=12)
        right.pack(side="right", fill="both", expand=True)

        # Tabs
        self.tabview = ctk.CTkTabview(right)
        self.tabview.pack(fill="both", expand=True, padx=12, pady=12)

        self.tabview.add("Results")
        self.tabview.add("Log")

        # Results tab — scrollable table
        results_tab = self.tabview.tab("Results")

        col_headers = ["Business Name", "Phone", "Email", "Website"]
        col_widths = [200, 130, 180, 180]

        header_row = ctk.CTkFrame(results_tab, fg_color="#1A5276", corner_radius=6)
        header_row.pack(fill="x", pady=(0, 2))

        for h, w in zip(col_headers, col_widths):
            ctk.CTkLabel(
                header_row, text=h,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="white", width=w, anchor="w"
            ).pack(side="left", padx=8, pady=6)

        self.results_scroll = ctk.CTkScrollableFrame(results_tab, corner_radius=6)
        self.results_scroll.pack(fill="both", expand=True)

        # Log tab
        log_tab = self.tabview.tab("Log")
        self.log_text = ctk.CTkTextbox(log_tab, font=ctk.CTkFont(family="Courier", size=12))
        self.log_text.pack(fill="both", expand=True)

        # Progress bar
        self.progress = ctk.CTkProgressBar(right, height=6)
        self.progress.pack(fill="x", padx=12, pady=(0, 4))
        self.progress.set(0)

        self.status_label = ctk.CTkLabel(right, text="Ready", font=ctk.CTkFont(size=11), text_color="#7F8C8D")
        self.status_label.pack(pady=(0, 8))

    def _log(self, message):
        """Add message to log tab."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.status_label.configure(text=message[:80])

    def _update_progress(self, current, total):
        """Update progress bar."""
        self.progress.set(current / total if total > 0 else 0)

    def _add_result_row(self, lead, idx):
        """Add a row to the results table."""
        row_color = "#1C2833" if idx % 2 == 0 else "#212F3D"
        row = ctk.CTkFrame(self.results_scroll, fg_color=row_color, corner_radius=4)
        row.pack(fill="x", pady=1)

        values = [
            lead.get("name", "")[:35],
            lead.get("phone", "")[:18],
            lead.get("email", "")[:28],
            lead.get("website", "")[:28],
        ]
        widths = [200, 130, 180, 180]
        colors = ["white", "#A9CCE3", "#A9DFBF", "#AED6F1"]

        for val, w, col in zip(values, widths, colors):
            ctk.CTkLabel(
                row, text=val or "—",
                font=ctk.CTkFont(size=11),
                text_color=col if val else "#555",
                width=w, anchor="w"
            ).pack(side="left", padx=8, pady=4)

    def _update_stats(self):
        """Update lead count labels."""
        total = len(self.leads)
        with_email = sum(1 for l in self.leads if l.get("email"))
        self.leads_count_label.configure(text=f"{total} leads found")
        self.email_count_label.configure(text=f"{with_email} with email")

    def _start_scraping(self):
        """Start scraping in background thread."""
        keyword = self.keyword_entry.get().strip()
        location = self.location_entry.get().strip()

        if not keyword or not location:
            messagebox.showerror("Missing Input", "Please enter both keyword and location.")
            return

        if not HAS_REQUESTS:
            messagebox.showerror("Missing Library", "Please install: pip install requests beautifulsoup4")
            return

        # Clear previous results
        for widget in self.results_scroll.winfo_children():
            widget.destroy()
        self.leads = []
        self.log_text.delete("1.0", "end")
        self.stop_event.clear()
        self.is_running = True

        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.export_btn.configure(state="disabled")
        self.progress.set(0)

        thread = threading.Thread(target=self._scrape_worker, args=(keyword, location), daemon=True)
        thread.start()

    def _scrape_worker(self, keyword, location):
        """Background scraping worker."""
        num = int(self.num_var.get())
        source = self.source_var.get()
        enrich = self.enrich_var.get()

        raw_leads = []

        if source in ("Google", "Both"):
            results = search_google(keyword, location, num, log_fn=lambda m: self.after(0, self._log, m))
            raw_leads.extend(results)

        if source in ("Yelp", "Both"):
            results = search_yelp(keyword, location, num, log_fn=lambda m: self.after(0, self._log, m))
            raw_leads.extend(results)

        self.after(0, self._log, f"Found {len(raw_leads)} initial results")

        # Deduplicate
        seen = set()
        unique = []
        for lead in raw_leads:
            key = lead.get("name", "").lower().strip()
            if key and key not in seen:
                seen.add(key)
                unique.append(lead)

        raw_leads = unique
        self.after(0, self._log, f"After deduplication: {len(raw_leads)} leads")

        # Enrich
        if enrich and not self.stop_event.is_set():
            self.after(0, self._log, "Extracting contact info from websites...")
            raw_leads = enrich_leads(
                raw_leads,
                log_fn=lambda m: self.after(0, self._log, m),
                progress_fn=lambda c, t: self.after(0, self._update_progress, c, t),
                stop_event=self.stop_event,
            )

        self.leads = raw_leads

        # Update UI with results
        for i, lead in enumerate(self.leads):
            self.after(0, self._add_result_row, lead, i)

        self.after(0, self._scraping_done)

    def _scraping_done(self):
        """Called when scraping finishes."""
        self.is_running = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.export_btn.configure(state="normal" if self.leads else "disabled")
        self.progress.set(1)
        self._update_stats()
        self._log(f"✅ Done! {len(self.leads)} leads collected")

    def _stop_scraping(self):
        """Stop the scraping process."""
        self.stop_event.set()
        self._log("⏹ Stopping...")
        self.stop_btn.configure(state="disabled")

    def _export_results(self):
        """Export results to file."""
        if not self.leads:
            messagebox.showwarning("No Data", "No leads to export.")
            return

        fmt = self.export_var.get()
        is_excel = "xlsx" in fmt

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx" if is_excel else ".csv",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")] if is_excel else [("CSV files", "*.csv")],
            initialfile=f"leads_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
        )

        if not filepath:
            return

        if is_excel:
            export_excel(self.leads, filepath)
        else:
            export_csv(self.leads, filepath)

        self._log(f"💾 Exported {len(self.leads)} leads → {filepath}")
        messagebox.showinfo("Exported!", f"✅ {len(self.leads)} leads saved to:\n{filepath}")

        if messagebox.askyesno("Open File?", "Open the exported file now?"):
            webbrowser.open(filepath)


# ─── Main ───────────────────────────────────────

if __name__ == "__main__":
    app = LeadScraperApp()
    app.mainloop()
