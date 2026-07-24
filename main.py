#!/usr/bin/env python3
import os
import sys
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except (AttributeError, ValueError):
    pass
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import logging
import queue
from datetime import date, datetime, timedelta

import pandas as pd
from tkcalendar import Calendar

from ConfigManager import ConfigManager
from TenderScraper import TenderScraper
from BatchProcessor import (create_batch_folder, save_daily_file,
                            create_end_product, update_equation_file,
                            calculate_counts, update_power_bi_export,
                            update_master_tenders, deduplicate_tenders)
from TenderSummary import create_tender_summary
from TenderAnalysisGenerator import create_tender_analysis
from CybersecurityTenders import create_cybersecurity_tenders

# ── Amidel brand colours ──────────────────────────────────────────────────────
NAVY        = "#1C3880"
NAVY_LIGHT  = "#254499"   # hover / active for navy elements
ORANGE      = "#F5A000"
ORANGE_DARK = "#E09000"   # hover / active for orange buttons
OFF_WHITE   = "#F4F6FA"   # form background
SEPARATOR   = "#D0D8EA"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Amidel eTender Scraper")
        self.resizable(False, False)
        self.configure(bg=OFF_WHITE)
        try:
            self.iconbitmap("amidel.ico")
        except Exception:
            pass
        self.log_queue = queue.Queue()
        self._batch_anchor = date.today()
        self._build_scraper_view()
        self._build_done_view()
        self._poll_log()
        self.scraper_frame.pack(fill="both", expand=True)

    # ── Batch week helpers ────────────────────────────────────────────────────

    def _anchor_range(self):
        anchor = self._batch_anchor
        wd = anchor.weekday()
        if self.batch_type_var.get() == "T":
            monday = anchor - timedelta(days=wd)
            d_from, d_to = monday, monday + timedelta(days=2)
        else:
            days_since_thu = (wd - 3) % 7
            thursday = anchor - timedelta(days=days_since_thu)
            d_from, d_to = thursday, thursday + timedelta(days=3)
        return d_from.strftime("%Y-%m-%d"), d_to.strftime("%Y-%m-%d")

    def _update_week_label(self):
        df, dt = self._anchor_range()
        start = datetime.strptime(df, "%Y-%m-%d")
        end   = datetime.strptime(dt, "%Y-%m-%d")
        s_day, e_day = str(start.day), str(end.day)
        if start.month == end.month:
            label = f"{start.strftime('%a')} {s_day} – {end.strftime('%a')} {e_day} {end.strftime('%b %Y')}"
        else:
            label = (f"{start.strftime('%a')} {s_day} {start.strftime('%b')}"
                     f" – {end.strftime('%a')} {e_day} {end.strftime('%b %Y')}")
        self.week_label_var.set(label)

    def _prev_week(self):
        self._batch_anchor -= timedelta(days=7)
        self._update_week_label()

    def _next_week(self):
        self._batch_anchor += timedelta(days=7)
        self._update_week_label()

    def _pick_week(self):
        top = tk.Toplevel(self)
        top.title("Select batch week")
        top.resizable(False, False)
        top.grab_set()
        top.configure(bg="white")

        batch_type = self.batch_type_var.get()

        cal = Calendar(
            top, selectmode="day",
            year=self._batch_anchor.year,
            month=self._batch_anchor.month,
            day=self._batch_anchor.day,
            date_pattern="yyyy-mm-dd",
            font=("Segoe UI", 10),
            headersbackground=NAVY,
            headersforeground="white",
            selectbackground=ORANGE,
            selectforeground="white",
            normalbackground="white",
            weekendbackground="white",
            othermonthbackground="#f0f0f0",
            othermonthwebackground="#f0f0f0",
        )
        cal.pack(padx=12, pady=(12, 6))
        cal.tag_config("batch", background=ORANGE, foreground="white")

        range_var = tk.StringVar()
        tk.Label(top, textvariable=range_var, bg="white",
                 font=("Segoe UI", 10, "bold"), fg=NAVY).pack()

        hint = ("Click any day Mon – Wed in the target week"
                if batch_type == "T" else
                "Click any day Thu – Sun in the target week")
        tk.Label(top, text=hint, bg="white", fg="gray",
                 font=("Segoe UI", 8)).pack(pady=(2, 0))

        current_anchor = [self._batch_anchor]
        suppress = [False]

        def _compute_range(anchor):
            wd = anchor.weekday()
            if batch_type == "T":
                monday = anchor - timedelta(days=wd)
                return monday, monday + timedelta(days=2)
            else:
                thursday = anchor - timedelta(days=(wd - 3) % 7)
                return thursday, thursday + timedelta(days=3)

        def _highlight(anchor):
            cal.calevent_remove("all")
            d_from, d_to = _compute_range(anchor)
            d = d_from
            while d <= d_to:
                cal.calevent_create(d, "", "batch")
                d += timedelta(days=1)
            s, e = str(d_from.day), str(d_to.day)
            if d_from.month == d_to.month:
                label = (f"{d_from.strftime('%a')} {s} – "
                         f"{d_to.strftime('%a')} {e} {d_to.strftime('%b %Y')}")
            else:
                label = (f"{d_from.strftime('%a')} {s} {d_from.strftime('%b')} – "
                         f"{d_to.strftime('%a')} {e} {d_to.strftime('%b %Y')}")
            range_var.set(label)
            suppress[0] = True
            cal.selection_set(d_from)
            suppress[0] = False

        _highlight(self._batch_anchor)

        def _on_date_click(event):
            if suppress[0]:
                return
            selected = datetime.strptime(cal.get_date(), "%Y-%m-%d").date()
            current_anchor[0] = selected
            _highlight(selected)

        cal.bind("<<CalendarSelected>>", _on_date_click)

        def _on_confirm():
            self._batch_anchor = current_anchor[0]
            self._update_week_label()
            top.destroy()

        tk.Button(
            top, text="Confirm", command=_on_confirm,
            bg=ORANGE, fg="white", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=14, pady=6, cursor="hand2",
            activebackground=ORANGE_DARK, activeforeground="white",
        ).pack(pady=10)

    def _on_source_change(self):
        src = self.source_var.get()
        if src in ("etenders", "full_batch", "all_but_etenders"):
            self._lbl_batch_type.grid()
            self._radio_frame.grid()
            self._lbl_week.grid()
            self._nav_frame.grid()
            self._lbl_ecdpw_info.grid_remove()
            self._lbl_ecdpw_filter.grid_remove()
            self._ecdpw_filter_frame.grid_remove()
        else:  # ecdpw
            self._lbl_batch_type.grid_remove()
            self._radio_frame.grid_remove()
            self._lbl_week.grid_remove()
            self._nav_frame.grid_remove()
            self._lbl_ecdpw_info.grid()
            self._lbl_ecdpw_filter.grid()
            self._ecdpw_filter_frame.grid()

    # ── Header helper ─────────────────────────────────────────────────────────

    def _build_header(self, parent):
        hdr = tk.Frame(parent, bg=NAVY)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Amidel eTender Scraper", bg=NAVY, fg="white",
                 font=("Segoe UI", 15, "bold"), pady=10).pack()
        tk.Label(hdr, text="Beyond Technology", bg=NAVY, fg=ORANGE,
                 font=("Segoe UI", 8, "italic")).pack(pady=(0, 8))

    # ── Scraper view ──────────────────────────────────────────────────────────

    def _build_scraper_view(self):
        f = self.scraper_frame = tk.Frame(self, bg=OFF_WHITE)

        self._build_header(f)

        form = tk.Frame(f, bg=OFF_WHITE, padx=20, pady=16)
        form.pack(fill="x")

        today          = date.today()
        suggested_type = "M" if today.weekday() == 0 else "T"

        lbl_opts = dict(bg=OFF_WHITE, font=("Segoe UI", 10), fg="#222")

        # Source selector
        tk.Label(form, text="Source", **lbl_opts).grid(
            row=0, column=0, sticky="w")
        src_frame = tk.Frame(form, bg=OFF_WHITE)
        src_frame.grid(row=0, column=1, padx=(12, 0), pady=(0, 10), sticky="w")
        self.source_var = tk.StringVar(value="etenders")
        src_opts = dict(bg=OFF_WHITE, activebackground=OFF_WHITE,
                        selectcolor=ORANGE, font=("Segoe UI", 10),
                        variable=self.source_var, command=self._on_source_change)
        tk.Radiobutton(src_frame, text="eTenders", value="etenders",
                       **src_opts).pack(side="left")
        tk.Radiobutton(src_frame, text="EC DPW", value="ecdpw",
                       **src_opts).pack(side="left", padx=(14, 0))
        tk.Radiobutton(src_frame, text="Full Batch", value="full_batch",
                       **src_opts).pack(side="left", padx=(14, 0))
        tk.Radiobutton(src_frame, text="All but eTenders", value="all_but_etenders",
                       **src_opts).pack(side="left", padx=(14, 0))

        # Batch type (eTenders only)
        self._lbl_batch_type = tk.Label(form, text="Batch Type", **lbl_opts)
        self._lbl_batch_type.grid(row=1, column=0, sticky="w")
        radio_frame = tk.Frame(form, bg=OFF_WHITE)
        radio_frame.grid(row=1, column=1, padx=(12, 0), pady=(0, 10), sticky="w")
        self._radio_frame = radio_frame
        self.batch_type_var = tk.StringVar(value=suggested_type)
        radio_opts = dict(bg=OFF_WHITE, activebackground=OFF_WHITE,
                          selectcolor=ORANGE, font=("Segoe UI", 10),
                          variable=self.batch_type_var)
        tk.Radiobutton(radio_frame, text="(T) Thursday report",
                       value="T", command=self._update_week_label,
                       **radio_opts).pack(side="left")
        tk.Radiobutton(radio_frame, text="(M) Monday report",
                       value="M", command=self._update_week_label,
                       **radio_opts).pack(side="left", padx=(14, 0))

        # EC DPW info (shown when EC DPW selected, hidden otherwise)
        self._lbl_ecdpw_info = tk.Label(
            form,
            text="Scrapes all Open tenders with closing date ≥ 14 days from today",
            bg=OFF_WHITE, fg="#555", font=("Segoe UI", 9, "italic"),
        )
        self._lbl_ecdpw_info.grid(row=1, column=1, padx=(12, 0), sticky="w")
        self._lbl_ecdpw_info.grid_remove()

        # EC DPW Month/Year filter (EC DPW only)
        self._lbl_ecdpw_filter = tk.Label(form, text="Filter by", **lbl_opts)
        self._lbl_ecdpw_filter.grid(row=2, column=0, sticky="w")
        ecdpw_filter_frame = tk.Frame(form, bg=OFF_WHITE)
        ecdpw_filter_frame.grid(row=2, column=1, padx=(12, 0), pady=(0, 10), sticky="w")
        self._ecdpw_filter_frame = ecdpw_filter_frame

        _MONTHS = ["", "January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]
        self._ecdpw_months = _MONTHS
        self.ecdpw_month_var = tk.StringVar(value="")
        ttk.Combobox(ecdpw_filter_frame, textvariable=self.ecdpw_month_var,
                     values=_MONTHS, width=11, state="readonly").pack(side="left")
        tk.Label(ecdpw_filter_frame, text="Month", bg=OFF_WHITE,
                 font=("Segoe UI", 9), fg="#888").pack(side="left", padx=(3, 12))

        _cur_year = date.today().year
        _years = [""] + [str(y) for y in range(_cur_year, 2019, -1)]
        self.ecdpw_year_var = tk.StringVar(value="")
        ttk.Combobox(ecdpw_filter_frame, textvariable=self.ecdpw_year_var,
                     values=_years, width=7, state="readonly").pack(side="left")
        tk.Label(ecdpw_filter_frame, text="Year", bg=OFF_WHITE,
                 font=("Segoe UI", 9), fg="#888").pack(side="left", padx=(3, 0))

        self._lbl_ecdpw_filter.grid_remove()
        self._ecdpw_filter_frame.grid_remove()

        # Week navigator (eTenders only)
        self._lbl_week = tk.Label(form, text="Batch Week", **lbl_opts)
        self._lbl_week.grid(row=2, column=0, sticky="w")
        nav = tk.Frame(form, bg=OFF_WHITE)
        nav.grid(row=2, column=1, padx=(12, 0), sticky="w")
        self._nav_frame = nav

        nav_btn = dict(font=("Segoe UI", 10), width=3, relief="flat",
                       bg=NAVY, fg="white", cursor="hand2",
                       activebackground=NAVY_LIGHT, activeforeground="white")
        tk.Button(nav, text="◄", command=self._prev_week, **nav_btn).pack(side="left")

        self.week_label_var = tk.StringVar()
        tk.Label(nav, textvariable=self.week_label_var, bg=OFF_WHITE,
                 font=("Segoe UI", 10, "bold"), fg=NAVY,
                 width=28, anchor="center").pack(side="left", padx=8)

        tk.Button(nav, text="►", command=self._next_week, **nav_btn).pack(side="left")
        tk.Button(nav, text="📅", command=self._pick_week,
                  font=("Segoe UI", 11), relief="flat",
                  bg=NAVY, fg="white", cursor="hand2",
                  activebackground=NAVY_LIGHT).pack(side="left", padx=(8, 0))

        self._update_week_label()

        # Run button + status
        ctrl = tk.Frame(f, bg=OFF_WHITE, padx=20, pady=8)
        ctrl.pack(fill="x")
        self.run_btn = tk.Button(
            ctrl, text="Run Scraper", command=self._start,
            bg=ORANGE, fg="white", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=16, pady=7, cursor="hand2",
            activebackground=ORANGE_DARK, activeforeground="white",
        )
        self.run_btn.pack(side="left")
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(ctrl, textvariable=self.status_var, bg=OFF_WHITE,
                 fg="#666", font=("Segoe UI", 9)).pack(side="left", padx=14)

        # Separator + log panel
        tk.Frame(f, bg=SEPARATOR, height=1).pack(fill="x", padx=20, pady=(6, 0))
        self.log_box = scrolledtext.ScrolledText(
            f, state="disabled", height=18, width=72,
            font=("Consolas", 9), bg="#1a1a2e", fg="#c8d0e0",
            insertbackground="white", relief="flat", borderwidth=0,
        )
        self.log_box.pack(padx=20, pady=12)

    # ── Done view ─────────────────────────────────────────────────────────────

    def _build_done_view(self):
        f = self.done_frame = tk.Frame(self, bg="white")

        self._build_header(f)

        body = tk.Frame(f, bg="white", padx=40, pady=36)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="✓", bg="white", fg=ORANGE,
                 font=("Segoe UI", 72)).pack()
        tk.Label(body, text="Scraping Complete!", bg="white", fg=NAVY,
                 font=("Segoe UI", 22, "bold")).pack(pady=(0, 10))

        self.done_count_var = tk.StringVar()
        tk.Label(body, textvariable=self.done_count_var, bg="white",
                 font=("Segoe UI", 13), fg="#333").pack()

        self.done_dedup_var = tk.StringVar()
        tk.Label(body, textvariable=self.done_dedup_var, bg="white",
                 font=("Segoe UI", 10), fg="#888").pack()

        self.done_date_var = tk.StringVar()
        tk.Label(body, textvariable=self.done_date_var, bg="white",
                 font=("Segoe UI", 10), fg="#888").pack(pady=(4, 12))

        self.done_file_var = tk.StringVar()
        tk.Label(body, textvariable=self.done_file_var, bg="white",
                 font=("Segoe UI", 9), fg=NAVY).pack()

        self.done_equation_var = tk.StringVar()
        tk.Label(body, textvariable=self.done_equation_var, bg="white",
                 font=("Segoe UI", 9), fg=ORANGE).pack(pady=(2, 4))

        self.done_summary_var = tk.StringVar()
        tk.Label(body, textvariable=self.done_summary_var, bg="white",
                 font=("Segoe UI", 9), fg=NAVY).pack(pady=(0, 32))

        btn_row = tk.Frame(body, bg="white")
        btn_row.pack()
        tk.Button(
            btn_row, text="Scrape Again", command=self._scrape_again,
            bg=ORANGE, fg="white", font=("Segoe UI", 11, "bold"),
            relief="flat", padx=20, pady=10, cursor="hand2",
            activebackground=ORANGE_DARK, activeforeground="white",
        ).pack(side="left", padx=10)
        tk.Button(
            btn_row, text="Close", command=self.destroy,
            bg="#E0E4EE", fg=NAVY, font=("Segoe UI", 11),
            relief="flat", padx=20, pady=10, cursor="hand2",
            activebackground="#C8CEDF", activeforeground=NAVY,
        ).pack(side="left", padx=10)

    # ── Shared logic ──────────────────────────────────────────────────────────

    def _poll_log(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_box.configure(state="normal")
                self.log_box.insert(tk.END, msg + "\n")
                self.log_box.see(tk.END)
                self.log_box.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_log)

    def _start(self):
        self.run_btn.configure(state="disabled")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", tk.END)
        self.log_box.configure(state="disabled")
        if self.source_var.get() == "ecdpw":
            month_str = self.ecdpw_month_var.get()
            year_str  = self.ecdpw_year_var.get()
            if month_str and not year_str:
                from tkinter import messagebox
                self.run_btn.configure(state="normal")
                messagebox.showerror("Filter Error",
                    "Please also select a Year when filtering by Month.")
                return
            pub_date_from = None
            filter_label  = None
            if year_str:
                year  = int(year_str)
                if month_str:
                    month        = self._ecdpw_months.index(month_str)
                    pub_date_from = date(year, month, 1)
                    filter_label  = f"{year}-{month:02d}"
                else:
                    pub_date_from = date(year, 1, 1)
                    filter_label  = str(year)
            self.status_var.set("Scraping EC DPW tenders…")
            threading.Thread(target=self._run_ecdpw, args=(pub_date_from, filter_label),
                             daemon=True).start()
        elif self.source_var.get() == "full_batch":
            date_from, date_to = self._anchor_range()
            self.status_var.set(f"Scraping Full Batch {date_from} → {date_to}…")
            threading.Thread(target=self._run_watchlist,
                             args=(date_from, date_to),
                             kwargs={"skip_etenders": False,
                                     "folder_name": "full_batch",
                                     "filter_to_watchlist": True,
                                     "dedupe_cross_source": True},
                             daemon=True).start()
        elif self.source_var.get() == "all_but_etenders":
            date_from, date_to = self._anchor_range()
            self.status_var.set(f"Scraping All but eTenders {date_from} → {date_to}…")
            threading.Thread(target=self._run_watchlist,
                             args=(date_from, date_to),
                             kwargs={"skip_etenders": True, "folder_name": "all_but_etenders"},
                             daemon=True).start()
        else:
            date_from, date_to = self._anchor_range()
            self.status_var.set(f"Scraping {date_from} → {date_to}…")
            threading.Thread(target=self._run, args=(date_from, date_to), daemon=True).start()

    @staticmethod
    def _closing_not_expired(tender: dict, date_from: str) -> bool:
        """Return False if the tender has a closing date that predates the batch start."""
        cd = str(tender.get("CLOSING_DATE") or "").strip()
        if not cd:
            return True
        try:
            cd_date = datetime.strptime(cd, "%Y/%m/%d").date()
            return cd_date >= datetime.strptime(date_from, "%Y-%m-%d").date()
        except ValueError:
            return True

    @staticmethod
    def _publication_in_range(tender: dict, date_from: str, date_to: str) -> bool:
        """Return False if the tender's publication date falls outside the batch window.
        Blanks and unparseable dates are kept — many watchlist sources omit or malform
        this field, and dropping them silently would cause missed tenders."""
        pd = str(tender.get("PUBLICATION_DATE") or "").strip()
        if not pd:
            return True
        try:
            pd_date = datetime.strptime(pd, "%Y/%m/%d").date()
        except ValueError:
            return True
        df = datetime.strptime(date_from, "%Y-%m-%d").date()
        dt = datetime.strptime(date_to,   "%Y-%m-%d").date()
        return df <= pd_date <= dt

    def _run(self, date_from, date_to):
        end_product_path = None
        equation_updated = False
        summaries_count  = 0
        all_tenders      = []
        batch_type       = self.batch_type_var.get()
        report_date      = datetime.strptime(date_to, "%Y-%m-%d")

        start = datetime.strptime(date_from, "%Y-%m-%d")
        end   = datetime.strptime(date_to,   "%Y-%m-%d")
        days, current = [], start
        while current <= end:
            days.append(current.strftime("%Y-%m-%d"))
            current += timedelta(days=1)

        try:
            batch_folder = create_batch_folder(date_from, date_to, batch_type)

            for i, day in enumerate(days, 1):
                self.after(0, lambda msg=f"Scraping day {i} of {len(days)}: {day}…":
                           self.status_var.set(msg))

                cm = ConfigManager()
                cm.updateConfig({"scraping": {"dateFrom": day, "dateTo": day}})

                scraper = TenderScraper(log_queue=self.log_queue)
                scraper.run(export=False)

                try:
                    save_daily_file(scraper.tenderData or [], day, batch_folder)
                except Exception as e:
                    logging.error(f"Could not save daily file for {day}: {e}")
                if scraper.tenderData:
                    all_tenders.extend(scraper.tenderData)

            raw_tenders     = list(all_tenders)
            after_expired   = [t for t in all_tenders if self._closing_not_expired(t, date_from)]
            expired_removed = len(raw_tenders) - len(after_expired)
            after_pub       = [t for t in after_expired if self._publication_in_range(t, date_from, date_to)]
            pub_removed     = len(after_expired) - len(after_pub)
            if pub_removed:
                logging.info(f"Removed {pub_removed} tender(s) with publication date outside {date_from}..{date_to}")
            all_tenders     = deduplicate_tenders(after_pub)
            dupes_removed   = len(after_pub) - len(all_tenders)

            # DUPLICATED column defaults to 0 for eTenders mode (single source).
            for t in all_tenders:
                t.setdefault("DUPLICATED", 0)
            for t in raw_tenders:
                t.setdefault("DUPLICATED", 0)

            if all_tenders:
                df = pd.DataFrame(all_tenders)
                report_date_str = datetime.strptime(date_to, "%Y-%m-%d").strftime("%d %b %Y").lstrip("0")
                try:
                    end_product_path = create_end_product(
                        df, date_from, date_to, batch_type, report_date, batch_folder,
                        raw_df=pd.DataFrame(raw_tenders)
                    )
                    counts = calculate_counts(df)
                    update_equation_file(counts, batch_type, report_date, batch_folder)
                    update_power_bi_export(batch_folder, date_from, date_to, batch_type)
                    update_master_tenders(batch_folder)
                    equation_updated = True
                except Exception as e:
                    logging.error(f"Batch processing error: {e}")

                try:
                    summaries_count = create_tender_summary(df, batch_folder)
                except Exception as e:
                    logging.error(f"Tender Summary creation error: {e}")

                try:
                    create_tender_analysis(df, batch_folder, report_date_str)
                except Exception as e:
                    logging.error(f"Tender Analysis creation error: {e}")

                try:
                    create_cybersecurity_tenders(df, batch_folder)
                except Exception as e:
                    logging.error(f"Cybersecurity Tenders creation error: {e}")

        except Exception as e:
            self.after(0, self._on_error, str(e))
            return

        self.after(0, self._show_done, len(all_tenders), date_from, date_to,
                   end_product_path, equation_updated, summaries_count,
                   len(raw_tenders), expired_removed, dupes_removed, pub_removed)

    # ── Watchlist helpers ─────────────────────────────────────────────────────

    _WATCHLIST_FILE = (
        r"C:\Users\CheltonGraham\OneDrive - Amidel (Pty) Ltd"
        r"\Documents\Sales\Sales Auto Hub\Scraping and Reports\Websites.xlsx"
    )

    def _load_watchlist(self) -> set:
        """Return the set of source names from Websites.xlsx."""
        import openpyxl
        try:
            wb = openpyxl.load_workbook(self._WATCHLIST_FILE)
            ws = wb.active
            skip = {"Source", "Report Date", None}
            return {
                str(row[0]).strip()
                for row in ws.iter_rows(min_row=2, values_only=True)
                if row[0] and str(row[0]).strip() not in skip
            }
        except Exception as e:
            logging.warning(f"Could not load Websites.xlsx: {e}")
            return set()

    def _run_watchlist(self, date_from: str, date_to: str,
                       skip_etenders: bool = False,
                       folder_name: str = "All_Tenders",
                       filter_to_watchlist: bool = False,
                       dedupe_cross_source: bool = False):
        all_tenders      = []
        end_product_path = None
        equation_updated = False
        summaries_count  = 0
        batch_type       = self.batch_type_var.get()
        report_date      = datetime.strptime(date_to, "%Y-%m-%d")
        root_dir         = os.path.join("data", folder_name)

        start = datetime.strptime(date_from, "%Y-%m-%d")
        end   = datetime.strptime(date_to,   "%Y-%m-%d")
        days, current = [], start
        while current <= end:
            days.append(current.strftime("%Y-%m-%d"))
            current += timedelta(days=1)

        watchlist_sources = self._load_watchlist()

        try:
            batch_folder = create_batch_folder(date_from, date_to, batch_type,
                                               root_dir=root_dir)

            # ── eTenders (day-by-day) ─────────────────────────────────────────
            # Saved to daily batch files and Power BI only — NOT included in
            # the Checker/end-product, which shows watchlist tenders exclusively.
            # Skipped entirely when running the "All but eTenders" mode.
            if not skip_etenders:
                for i, day in enumerate(days, 1):
                    self.after(0, lambda msg=f"eTenders: day {i}/{len(days)}: {day}…":
                               self.status_var.set(msg))
                    cm = ConfigManager()
                    cm.updateConfig({"scraping": {"dateFrom": day, "dateTo": day}})
                    scraper = TenderScraper(log_queue=self.log_queue)
                    scraper.run(export=False)
                    try:
                        save_daily_file(scraper.tenderData or [], day, batch_folder)
                    except Exception as e:
                        logging.error(f"Could not save daily file for {day}: {e}")
                    # In Full Batch mode, include eTenders portal tenders in the
                    # combined all_tenders list so cross-source dedupe against the
                    # watchlist scrapes can happen. Tag them so we know their origin.
                    if dedupe_cross_source and scraper.tenderData:
                        for t in scraper.tenderData:
                            t["_from_etenders"] = True
                        all_tenders.extend(scraper.tenderData)
            else:
                logging.info("All but eTenders mode: skipping eTenders.gov.za portal scrape")

            # ── EC DPW ────────────────────────────────────────────────────────
            self.after(0, lambda: self.status_var.set("Scraping EC DPW tenders…"))
            try:
                from ECDPWScraper import ECDPWScraper
                pub_from = datetime.strptime(date_from, "%Y-%m-%d").date()
                ec = ECDPWScraper(log_queue=self.log_queue, pub_date_from=pub_from)
                ec.run(status_callback=lambda msg: self.after(0, lambda m=msg: self.status_var.set(m)))
                if ec.tenderData:
                    all_tenders.extend(ec.tenderData)
                    logging.info(f"EC DPW: added {len(ec.tenderData)} tender(s)")
            except Exception as e:
                logging.error(f"EC DPW scraping error: {e}")

            # ── JPC ───────────────────────────────────────────────────────────
            if "JPC" in watchlist_sources:
                self.after(0, lambda: self.status_var.set("Scraping JPC tenders…"))
                try:
                    tenders = self._scrape_jpc(date_from, date_to)
                    all_tenders.extend(tenders)
                    logging.info(f"JPC: {len(tenders)} tender(s)")
                except Exception as e:
                    logging.error(f"JPC scraping error: {e}")

            # ── Raymond Mhlaba ────────────────────────────────────────────────
            if "Raymond Mhlaba LM" in watchlist_sources:
                self.after(0, lambda: self.status_var.set("Scraping Raymond Mhlaba tenders…"))
                try:
                    tenders = self._scrape_raymond_mhlaba(date_from, date_to)
                    all_tenders.extend(tenders)
                    logging.info(f"Raymond Mhlaba: {len(tenders)} tender(s)")
                except Exception as e:
                    logging.error(f"Raymond Mhlaba scraping error: {e}")

            # ── All other watchlist sites (static HTML) ───────────────────────
            self.after(0, lambda: self.status_var.set("Scraping watchlist sites…"))
            try:
                from WatchlistScrapers import run_watchlist_scrapers
                wl_tenders = run_watchlist_scrapers(
                    date_from, date_to, watchlist_sources, self.log_queue
                )
                all_tenders.extend(wl_tenders)
            except Exception as e:
                logging.error(f"Watchlist scrapers error: {e}")

            # ── Selenium-rendered watchlist sites ─────────────────────────────
            self.after(0, lambda: self.status_var.set("Scraping JS-rendered watchlist sites…"))
            try:
                from SeleniumWatchlistScrapers import run_selenium_watchlist_scrapers
                sel_tenders = run_selenium_watchlist_scrapers(
                    date_from, date_to, watchlist_sources, self.log_queue
                )
                all_tenders.extend(sel_tenders)
            except Exception as e:
                logging.error(f"Selenium watchlist scrapers error: {e}")

            raw_tenders     = list(all_tenders)
            after_expired   = [t for t in all_tenders if self._closing_not_expired(t, date_from)]
            expired_removed = len(raw_tenders) - len(after_expired)
            after_pub       = [t for t in after_expired if self._publication_in_range(t, date_from, date_to)]
            pub_removed     = len(after_expired) - len(after_pub)
            if pub_removed:
                logging.info(f"Removed {pub_removed} tender(s) with publication date outside {date_from}..{date_to}")
            all_tenders     = deduplicate_tenders(after_pub)
            dupes_removed   = len(after_pub) - len(all_tenders)

            # Full Batch mode: cross-source dedupe only. We deliberately do NOT
            # filter to the watchlist here — all scraped tenders are kept in the
            # end product; the Display Equation's TOTAL WATCHLIST row shows the
            # watchlist-matched subset alongside the full total.
            if dedupe_cross_source:
                from BatchProcessor import merge_and_flag_duplicates
                before = len(all_tenders)
                all_tenders = merge_and_flag_duplicates(all_tenders, prefer_etenders=True)
                num_dupes = sum(1 for t in all_tenders if t.get("DUPLICATED") == 1)
                logging.info(f"Cross-source dedupe: {before} -> {len(all_tenders)} unique, "
                             f"{num_dupes} flagged as duplicated across sources")
            else:
                # Every downstream sheet expects the column to exist. Default to 0.
                for t in all_tenders:
                    t.setdefault("DUPLICATED", 0)
                for t in raw_tenders:
                    t.setdefault("DUPLICATED", 0)

            # When skipping eTenders, write per-day batch files from the watchlist
            # tenders themselves — split by PUBLICATION_DATE — so batches/ isn't empty.
            if skip_etenders:
                for day in days:
                    day_slash = day.replace("-", "/")
                    day_tenders = [t for t in all_tenders
                                   if str(t.get("PUBLICATION_DATE") or "").strip() == day_slash]
                    try:
                        save_daily_file(day_tenders, day, batch_folder)
                    except Exception as e:
                        logging.error(f"Could not save daily file for {day}: {e}")

            if all_tenders:
                df = pd.DataFrame(all_tenders)
                report_date_str = datetime.strptime(date_to, "%Y-%m-%d").strftime("%d %b %Y").lstrip("0")
                try:
                    end_product_path = create_end_product(
                        df, date_from, date_to, batch_type, report_date, batch_folder,
                        raw_df=pd.DataFrame(raw_tenders),
                        no_etenders=skip_etenders,
                    )
                    counts = calculate_counts(df)
                    update_equation_file(counts, batch_type, report_date, batch_folder,
                                         no_etenders=skip_etenders)
                    update_power_bi_export(batch_folder, date_from, date_to, batch_type)
                    update_master_tenders(batch_folder)
                    equation_updated = True
                except Exception as e:
                    logging.error(f"Watchlist batch processing error: {e}")

                try:
                    summaries_count = create_tender_summary(df, batch_folder)
                except Exception as e:
                    logging.error(f"Watchlist Tender Summary error: {e}")

                try:
                    create_tender_analysis(df, batch_folder, report_date_str)
                except Exception as e:
                    logging.error(f"Watchlist Tender Analysis creation error: {e}")

                try:
                    create_cybersecurity_tenders(df, batch_folder)
                except Exception as e:
                    logging.error(f"Watchlist Cybersecurity Tenders creation error: {e}")

        except Exception as e:
            self.after(0, self._on_error, str(e))
            return

        self.after(0, self._show_done, len(all_tenders), date_from, date_to,
                   end_product_path, equation_updated, summaries_count,
                   len(raw_tenders), expired_removed, dupes_removed, pub_removed)

    def _scrape_jpc(self, date_from: str, date_to: str) -> list:
        from JPCScraper import JPCScraper
        s = JPCScraper(date_from=date_from, date_to=date_to, log_queue=self.log_queue)
        return s.run()

    def _scrape_raymond_mhlaba(self, date_from: str, date_to: str) -> list:
        from RaymondMhlabaScraper import RaymondMhlabaScraper
        s = RaymondMhlabaScraper(date_to=date_to, log_queue=self.log_queue)
        return s.run()

    def _run_ecdpw(self, pub_date_from=None, filter_label=None):
        from ECDPWScraper import ECDPWScraper
        try:
            scraper = ECDPWScraper(log_queue=self.log_queue, pub_date_from=pub_date_from,
                                   filter_label=filter_label)
            filepath = scraper.run(
                status_callback=lambda msg: self.after(0, lambda m=msg: self.status_var.set(m))
            )
            count = len(scraper.tenderData)
            self.after(0, self._show_done, count, "EC DPW", "", filepath, False, 0)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _show_done(self, count, date_from, date_to, end_product_path, equation_updated,
                   summaries_count=0, raw_count=None, expired_removed=0, dupes_removed=0,
                   pub_removed=0):
        self.done_count_var.set(
            f"{count} tender{'s' if count != 1 else ''} scraped"
        )
        parts = []
        if expired_removed:
            parts.append(f"{expired_removed} expired removed")
        if pub_removed:
            parts.append(f"{pub_removed} out-of-range removed")
        if dupes_removed:
            parts.append(f"{dupes_removed} duplicate{'s' if dupes_removed != 1 else ''} removed")
        if parts and raw_count:
            parts.append(f"{raw_count} total found")
        self.done_dedup_var.set("  ·  ".join(parts) if parts else "")
        if date_from and date_to:
            self.done_date_var.set(f"{date_from}  →  {date_to}")
        else:
            self.done_date_var.set(date_from or "")

        is_etenders = equation_updated or (date_from != "EC DPW")
        file_label  = "End product" if is_etenders else "Output file"
        if end_product_path:
            self.done_file_var.set(f"{file_label}: {os.path.basename(end_product_path)}")
        elif count == 0:
            self.done_file_var.set(f"No output created — 0 tenders found")
        else:
            self.done_file_var.set(f"{file_label} creation failed — check log")

        self.done_equation_var.set(
            "Equation file updated" if equation_updated else ""
        )

        if summaries_count > 0:
            self.done_summary_var.set(
                f"Tender Summary created — {summaries_count} tab{'s' if summaries_count != 1 else ''}"
            )
        else:
            self.done_summary_var.set("")

        self.scraper_frame.pack_forget()
        self.done_frame.pack(fill="both", expand=True)

    def _scrape_again(self):
        self.done_frame.pack_forget()
        self.run_btn.configure(state="normal")
        self.status_var.set("Ready")
        self.scraper_frame.pack(fill="both", expand=True)

    def _on_error(self, message):
        logging.error(f"Fatal scraper error: {message}")
        self.run_btn.configure(state="normal")
        self.status_var.set("Error — see log for details")
        from tkinter import messagebox
        messagebox.showerror("Scraper error", message)


if __name__ == "__main__":
    App().mainloop()
