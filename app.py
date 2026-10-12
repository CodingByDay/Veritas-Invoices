"""Veritas Software invoice desktop app."""

from __future__ import annotations

import ctypes
import os
from datetime import date
from decimal import Decimal
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from export_invoice import write_pdf
from invoice_model import (
    OUTPUT_DIR,
    PROFILES,
    Party,
    Invoice,
    Line,
    default_due_date,
    default_issue_date,
    english_stored_path,
    format_date,
    load_state,
    money_with_currency,
    output_pdf_path,
    parse_date,
    parse_money,
    period_label,
    remember_invoice,
    save_state,
    upsert_customer,
)

BG = "#eef2f6"
WHITE = "#ffffff"
NAVY = "#0e2a47"
NAVY_SOFT = "#163a63"
INK = "#1c2834"
MUTED = "#5c6b7a"
LINE = "#d5dde6"
FIELD = "#f8fafc"
TOTAL_BG = "#e6e8ff"
PLACEHOLDER = "Service name"


def _enable_dpi() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class ItemRow(tk.Frame):
    def __init__(self, master, on_change, on_remove):
        super().__init__(master, bg=WHITE)
        self.on_remove = on_remove
        entry = dict(
            relief="flat",
            bd=0,
            bg=FIELD,
            fg=INK,
            insertbackground=INK,
            highlightthickness=1,
            highlightbackground=LINE,
            highlightcolor=NAVY,
            font=("Segoe UI", 13),
        )
        self.name = tk.Entry(self, **entry)
        self.qty = tk.Entry(self, width=6, justify="center", **entry)
        self.price = tk.Entry(self, width=10, justify="right", **entry)
        self.qty.insert(0, "1")
        self.total = tk.Label(self, text="", bg=WHITE, fg=NAVY, font=("Segoe UI", 13, "bold"), anchor="e")
        self.remove_btn = tk.Button(
            self,
            text="✕",
            command=lambda: on_remove(self),
            bg=WHITE,
            fg=MUTED,
            activebackground="#fde8e8",
            activeforeground="#9b1c1c",
            relief="flat",
            bd=0,
            font=("Segoe UI", 11),
            cursor="hand2",
            width=3,
        )
        self.name.grid(row=0, column=0, sticky="ew", padx=(0, 8), ipady=8)
        self.qty.grid(row=0, column=1, sticky="ew", padx=(0, 8), ipady=8)
        self.price.grid(row=0, column=2, sticky="ew", padx=(0, 8), ipady=8)
        self.total.grid(row=0, column=3, sticky="ew", padx=(0, 4))
        self.remove_btn.grid(row=0, column=4)
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, minsize=88)
        self.columnconfigure(2, minsize=130)
        self.columnconfigure(3, minsize=120)
        self.columnconfigure(4, minsize=36)
        for widget in (self.name, self.qty, self.price):
            widget.bind("<KeyRelease>", lambda _event: on_change())
        self.name.bind("<FocusIn>", self._clear_placeholder)
        self.name.bind("<FocusOut>", self._restore_placeholder)

    def set_placeholder(self) -> None:
        if not self.name.get():
            self.name.insert(0, PLACEHOLDER)
            self.name.config(fg="#98a2b3")

    def _clear_placeholder(self, _event=None) -> None:
        if self.name.get() == PLACEHOLDER and self.name.cget("fg") == "#98a2b3":
            self.name.delete(0, "end")
            self.name.config(fg=INK)

    def _restore_placeholder(self, _event=None) -> None:
        if not self.name.get().strip():
            self.set_placeholder()

    def service_name(self) -> str:
        if self.name.cget("fg") == "#98a2b3":
            return ""
        text = self.name.get().strip()
        return "" if text == PLACEHOLDER else text


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Veritas Invoices")
        self.configure(bg=BG)
        self.minsize(900, 640)
        self.state = load_state()
        self.country = tk.StringVar(value=self.state.get("poslednji_format", "slovenija"))
        self.due_manual = False
        self._suspend = False
        self.customers: list[dict] = list(self.state.get("kupci_lista") or [])
        self.total_var = tk.StringVar(value="0.00")
        self.meta_var = tk.StringVar(value="")
        self.bank_var = tk.StringVar(value="")
        self.words_var = tk.StringVar(value="")
        self.customer_var = tk.StringVar()
        self.item_rows: list[ItemRow] = []
        self.history_paths: dict[str, str] = {}
        self.history_index: dict[str, int] = {}
        self.selected_history: str | None = None
        self.country_buttons: dict[str, tk.Button] = {}
        self._build()
        self._apply_country()
        self._select_customer_for_country(self.country.get())
        self._apply_default_dates()
        self._refresh_history()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._center(1000, 780)

    def _center(self, width: int, height: int) -> None:
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(width, screen_w - 40)
        height = min(height, screen_h - 80)
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _build(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "History.Treeview",
            background=WHITE,
            fieldbackground=WHITE,
            foreground=INK,
            rowheight=36,
            borderwidth=0,
            font=("Segoe UI", 13),
        )
        style.configure(
            "History.Treeview.Heading",
            background="#e7edf4",
            foreground=NAVY,
            font=("Segoe UI", 12, "bold"),
            relief="flat",
        )
        style.map("History.Treeview", background=[("selected", "#d9e4f5")], foreground=[("selected", NAVY)])
        style.configure(
            "Customer.TCombobox",
            fieldbackground=FIELD,
            background=WHITE,
            foreground=INK,
            padding=6,
            font=("Segoe UI", 13),
        )

        footer = tk.Frame(self, bg=WHITE, highlightthickness=1, highlightbackground=LINE)
        footer.pack(side="bottom", fill="x")
        footer_inner = tk.Frame(footer, bg=WHITE, padx=22, pady=12)
        footer_inner.pack(fill="x")
        self.generate_btn = tk.Button(
            footer_inner,
            text="Create PDF",
            command=self._generate,
            bg=NAVY,
            fg=WHITE,
            activebackground=NAVY_SOFT,
            activeforeground=WHITE,
            relief="flat",
            bd=0,
            padx=18,
            pady=10,
            font=("Segoe UI", 14, "bold"),
            cursor="hand2",
        )
        self.generate_btn.pack(side="left")
        tk.Button(
            footer_inner,
            text="Open folder",
            command=self._open_folder,
            bg="#e7edf4",
            fg=NAVY,
            activebackground="#d5dee8",
            relief="flat",
            bd=0,
            padx=14,
            pady=10,
            font=("Segoe UI", 13, "bold"),
            cursor="hand2",
        ).pack(side="left", padx=(8, 0))
        self.status = tk.Label(footer_inner, text="", bg=WHITE, fg=MUTED, font=("Segoe UI", 12), anchor="w")
        self.status.pack(side="left", padx=14, fill="x", expand=True)
        total_box = tk.Frame(footer_inner, bg=TOTAL_BG, padx=14, pady=4)
        total_box.pack(side="right")
        tk.Label(total_box, text="Amount due", bg=TOTAL_BG, fg=MUTED, font=("Segoe UI", 12)).pack(anchor="e")
        tk.Label(
            total_box,
            textvariable=self.total_var,
            bg=TOTAL_BG,
            fg=NAVY,
            font=("Segoe UI", 20, "bold"),
        ).pack(anchor="e")
        tk.Label(
            footer_inner,
            textvariable=self.words_var,
            bg=WHITE,
            fg=MUTED,
            font=("Segoe UI", 12, "italic"),
            wraplength=240,
            justify="right",
        ).pack(side="right", padx=(0, 12))

        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=BG)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.canvas_window, width=e.width))
        self.bind_all("<MouseWheel>", self._on_mousewheel)

        wrap = tk.Frame(self.body, bg=BG)
        wrap.pack(fill="both", expand=True, padx=22, pady=16)

        header = tk.Frame(wrap, bg=BG)
        header.pack(fill="x")
        tk.Label(header, text="VERITAS INVOICES", bg=BG, fg=NAVY, font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(
            header,
            text="New invoice",
            bg=BG,
            fg=INK,
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Choose the format and customer, then enter the services and prices.",
            bg=BG,
            fg=MUTED,
            font=("Segoe UI", 13),
        ).pack(anchor="w", pady=(2, 12))

        switch = tk.Frame(wrap, bg=BG)
        switch.pack(fill="x", pady=(0, 12))
        switch.columnconfigure(0, weight=1)
        switch.columnconfigure(1, weight=1)
        for col, key in enumerate(("slovenija", "srbija")):
            button = tk.Button(
                switch,
                text="Slovenia" if key == "slovenija" else "Serbia",
                command=lambda value=key: self._select_country(value),
                relief="flat",
                bd=0,
                cursor="hand2",
                font=("Segoe UI", 14, "bold"),
                justify="left",
                anchor="w",
                padx=16,
                pady=12,
            )
            button.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 8, 0))
            self.country_buttons[key] = button

        details, details_inner = self._card(wrap)
        details.pack(fill="x", pady=(0, 12))
        for col in range(3):
            details_inner.columnconfigure(col, weight=1)
        self.number_entry = self._field(details_inner, "Invoice number", 0, 0)
        self.number_entry.insert(0, str(self.state["sledeci_broj"]))
        self.number_entry.config(font=("Segoe UI", 16, "bold"))
        self.date_entry = self._field(details_inner, "Invoice date", 0, 1)
        self.date_entry.insert(0, format_date(default_issue_date()))
        self.date_entry.bind("<KeyRelease>", lambda _e: self._on_date_changed())
        self.due_entry = self._field(details_inner, "Payment deadline", 0, 2)
        self.due_entry.insert(0, format_date(default_due_date(default_issue_date())))
        self.due_entry.bind("<KeyRelease>", lambda _e: self._mark_due_manual())
        self._text_button(details_inner, "Save number", self._remember_number).grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )
        self._text_button(details_inner, "Default dates", self._apply_default_dates).grid(
            row=1, column=1, sticky="w", pady=(8, 0)
        )
        self.place_entry = self._field(details_inner, "Place of supply", 2, 0)
        self.period_entry = self._field(details_inner, "Service month", 2, 1)
        self.place_entry.bind("<KeyRelease>", lambda _e: self._update_meta())
        self.period_entry.bind("<KeyRelease>", lambda _e: self._update_meta())
        tk.Label(
            details_inner,
            textvariable=self.meta_var,
            bg=WHITE,
            fg=NAVY,
            font=("Segoe UI", 12, "italic"),
            anchor="w",
        ).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        tk.Label(
            details_inner,
            text="The invoice date defaults to the last day of this month, and the deadline to the 10th of the next month. You can change both dates and the invoice number.",
            bg=WHITE,
            fg=MUTED,
            font=("Segoe UI", 12),
            anchor="w",
            justify="left",
            wraplength=860,
        ).grid(row=4, column=0, columnspan=3, sticky="ew", pady=(6, 0))

        client_card, client = self._card(wrap)
        client_card.pack(fill="x", pady=(0, 12))
        client.columnconfigure(0, weight=1)
        client.columnconfigure(1, weight=1)
        client_head = tk.Frame(client, bg=WHITE)
        client_head.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(client_head, text="CUSTOMER", bg=WHITE, fg=NAVY, font=("Segoe UI", 14, "bold")).pack(side="left")
        tk.Label(
            client_head,
            text="Saved customers",
            bg=WHITE,
            fg=MUTED,
            font=("Segoe UI", 12),
        ).pack(side="right")
        self.customer_combo = ttk.Combobox(
            client,
            textvariable=self.customer_var,
            state="readonly",
            style="Customer.TCombobox",
            font=("Segoe UI", 13),
        )
        self.customer_combo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.customer_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_customer_picked())
        client_actions = tk.Frame(client, bg=WHITE)
        client_actions.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self._text_button(client_actions, "Save customer", self._save_customer).pack(side="left")
        self._text_button(client_actions, "Delete customer", self._delete_customer).pack(side="left", padx=(8, 0))
        self._text_button(client_actions, "New customer", self._new_customer).pack(side="left", padx=(8, 0))
        self.buyer_name = self._field(client, "Company name", 3, 0, columnspan=2)
        self.buyer_address = self._field(client, "Address", 5, 0, columnspan=2)
        self.buyer_mb = self._field(client, "Registration number", 7, 0)
        self.buyer_pib = self._field(client, "Tax number", 7, 1)
        tk.Label(
            client,
            textvariable=self.bank_var,
            bg=WHITE,
            fg=MUTED,
            font=("Segoe UI", 12),
            anchor="w",
            justify="left",
            wraplength=860,
        ).grid(row=9, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        items_card, items = self._card(wrap)
        items_card.pack(fill="x", pady=(0, 12))
        items.columnconfigure(0, weight=1)
        tk.Label(items, text="SERVICES", bg=WHITE, fg=NAVY, font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        heads = tk.Frame(items, bg=WHITE)
        heads.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        heads.columnconfigure(0, weight=1)
        heads.columnconfigure(1, minsize=88)
        heads.columnconfigure(2, minsize=130)
        heads.columnconfigure(3, minsize=120)
        heads.columnconfigure(4, minsize=36)
        self.name_header = tk.Label(heads, text="Service", bg=WHITE, fg=MUTED, font=("Segoe UI", 12), anchor="w")
        self.qty_header = tk.Label(heads, text="Qty", bg=WHITE, fg=MUTED, font=("Segoe UI", 12), anchor="center")
        self.price_header = tk.Label(heads, text="Price", bg=WHITE, fg=MUTED, font=("Segoe UI", 12), anchor="e")
        self.line_header = tk.Label(heads, text="Total", bg=WHITE, fg=MUTED, font=("Segoe UI", 12), anchor="e")
        self.name_header.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.qty_header.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.price_header.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.line_header.grid(row=0, column=3, sticky="ew", padx=(0, 4))

        self.rows_frame = tk.Frame(items, bg=WHITE)
        self.rows_frame.grid(row=2, column=0, sticky="ew")
        self._add_item(placeholder=True)

        actions = tk.Frame(items, bg=WHITE)
        actions.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        add = tk.Button(
            actions,
            text="+ Add service",
            command=lambda: self._add_item(placeholder=False),
            bg="#eef3f8",
            fg=NAVY,
            activebackground="#e1e9f2",
            relief="flat",
            bd=0,
            padx=12,
            pady=6,
            font=("Segoe UI", 13, "bold"),
            cursor="hand2",
        )
        add.pack(side="left")

        history_card, history = self._card(wrap)
        history_card.pack(fill="x")
        top = tk.Frame(history, bg=WHITE)
        top.pack(fill="x")
        tk.Label(top, text="SAVED INVOICES", bg=WHITE, fg=NAVY, font=("Segoe UI", 14, "bold")).pack(side="left")
        tk.Button(
            top,
            text="Delete invoice",
            command=self._delete_selected,
            bg=WHITE,
            fg="#9b1c1c",
            activeforeground="#9b1c1c",
            relief="flat",
            bd=0,
            font=("Segoe UI", 12, "bold"),
            cursor="hand2",
        ).pack(side="right")
        tk.Button(
            top,
            text="Open PDF",
            command=self._open_selected,
            bg=WHITE,
            fg=NAVY,
            relief="flat",
            bd=0,
            font=("Segoe UI", 12, "bold"),
            cursor="hand2",
        ).pack(side="right", padx=(0, 12))
        tk.Button(
            top,
            text="Refresh",
            command=self._reload_history,
            bg=WHITE,
            fg=NAVY,
            relief="flat",
            bd=0,
            font=("Segoe UI", 12, "bold"),
            cursor="hand2",
        ).pack(side="right")
        tk.Label(
            history,
            text="The invoice number is stored in the program and in Next bill number.txt.",
            bg=WHITE,
            fg=MUTED,
            font=("Segoe UI", 12),
        ).pack(anchor="w", pady=(2, 8))
        columns = ("broj", "datum", "format", "kupac", "iznos")
        self.tree = ttk.Treeview(
            history,
            columns=columns,
            show="headings",
            style="History.Treeview",
            height=4,
            selectmode="browse",
        )
        headings = {
            "broj": ("Number", 90),
            "datum": ("Date", 130),
            "format": ("Format", 120),
            "kupac": ("Customer", 280),
            "iznos": ("Amount", 160),
        }
        for key, (title, width) in headings.items():
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor="w" if key in ("kupac", "format") else "center")
        self.tree.column("iznos", anchor="e")
        self.tree.pack(fill="x")
        self.tree.bind("<Button-1>", self._remember_history_click, add="+")
        self.tree.bind("<<TreeviewSelect>>", self._remember_history_selection)
        self.tree.bind("<Double-1>", lambda _e: self._open_selected())
        self.tree.bind("<Delete>", lambda _e: self._delete_selected())
        self.tree.tag_configure("odd", background="#f7f9fc")

    def _card(self, parent) -> tuple[tk.Frame, tk.Frame]:
        outer = tk.Frame(parent, bg=LINE, padx=1, pady=1)
        inner = tk.Frame(outer, bg=WHITE, padx=16, pady=14)
        inner.pack(fill="both", expand=True)
        return outer, inner

    def _field(self, parent, label: str, row: int, column: int, columnspan: int = 1) -> tk.Entry:
        box = tk.Frame(parent, bg=WHITE)
        box.grid(row=row, column=column, columnspan=columnspan, sticky="ew", padx=(0, 10), pady=(8, 0))
        tk.Label(box, text=label, bg=WHITE, fg=MUTED, font=("Segoe UI", 11)).pack(anchor="w")
        entry = tk.Entry(
            box,
            relief="flat",
            bd=0,
            bg=FIELD,
            fg=INK,
            insertbackground=INK,
            highlightthickness=1,
            highlightbackground=LINE,
            highlightcolor=NAVY,
            font=("Segoe UI", 13),
        )
        entry.pack(fill="x", ipady=6)
        return entry

    def _text_button(self, parent, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg="#eef3f8",
            fg=NAVY,
            activebackground="#e1e9f2",
            activeforeground=NAVY,
            relief="flat",
            bd=0,
            padx=12,
            pady=6,
            font=("Segoe UI", 12, "bold"),
            cursor="hand2",
        )

    def _on_mousewheel(self, event) -> None:
        if isinstance(event.widget, ttk.Treeview):
            return
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _select_country(self, key: str) -> None:
        if self._suspend:
            return
        if self.country.get() == key:
            self._apply_country()
            return
        self.country.set(key)
        self._apply_country()
        self._select_customer_for_country(key)

    def _apply_country(self) -> None:
        profile = PROFILES[self.country.get()]
        for key, button in self.country_buttons.items():
            selected = key == profile.key
            button.config(
                bg=NAVY if selected else WHITE,
                fg=WHITE if selected else NAVY,
                activebackground=NAVY_SOFT if selected else "#f3f6f9",
                activeforeground=WHITE if selected else NAVY,
            )
        self.price_header.config(text="Price")
        self.bank_var.set(profile.payment_hint)
        self.words_var.set(
            "Amount in words is on the invoice."
            if profile.use_words
            else "No amount in words on this invoice."
        )
        self._update_meta()
        self._recalc()

    def _customer_label(self, customer: dict) -> str:
        country = "Slovenia" if customer.get("format") == "slovenija" else "Serbia"
        return f"{customer.get('naziv', '')}   ·   {country}"

    def _refresh_customer_combo(self) -> None:
        labels = [self._customer_label(customer) for customer in self.customers]
        self.customer_combo["values"] = labels

    def _set_combo_to(self, customer_id: str | None) -> None:
        self._suspend = True
        try:
            self._refresh_customer_combo()
            if not customer_id:
                self.customer_var.set("")
                return
            for customer in self.customers:
                if customer.get("id") == customer_id:
                    self.customer_var.set(self._customer_label(customer))
                    return
            self.customer_var.set("")
        finally:
            self._suspend = False

    def _current_customer(self) -> dict | None:
        label = self.customer_var.get().strip()
        if not label:
            return None
        for customer in self.customers:
            if self._customer_label(customer) == label:
                return customer
        return None

    def _fill_customer(self, customer: dict) -> None:
        self._set_entry(self.buyer_name, customer.get("naziv", ""))
        self._set_entry(self.buyer_address, customer.get("adresa", ""))
        self._set_entry(self.buyer_mb, customer.get("mb", ""))
        self._set_entry(self.buyer_pib, customer.get("pib", ""))
        self._set_entry(self.place_entry, customer.get("mesto", ""))
        self._update_meta()

    def _select_customer_for_country(self, key: str) -> None:
        preferred = "insist" if key == "slovenija" else "landfreight"
        chosen = next((item for item in self.customers if item.get("id") == preferred), None)
        if chosen is None:
            chosen = next((item for item in self.customers if item.get("format") == key), None)
        if chosen is None:
            profile = PROFILES[key]
            self._set_combo_to(None)
            self._fill_customer(
                {
                    "naziv": profile.buyer.name,
                    "adresa": profile.buyer.address,
                    "mb": profile.buyer.registration,
                    "pib": profile.buyer.tax,
                    "mesto": profile.place,
                }
            )
            return
        self._set_combo_to(chosen.get("id"))
        self._fill_customer(chosen)

    def _on_customer_picked(self) -> None:
        if self._suspend:
            return
        customer = self._current_customer()
        if customer is None:
            return
        country = customer.get("format", self.country.get())
        if country not in PROFILES:
            country = self.country.get()
        self._suspend = True
        try:
            self.country.set(country)
            self._apply_country()
            self._fill_customer(customer)
        finally:
            self._suspend = False

    def _save_customer(self) -> None:
        data = self._read_client_form()
        if not data["naziv"]:
            messagebox.showerror("Customer", "Enter the company name before saving.")
            return
        if not data["mesto"]:
            data["mesto"] = PROFILES[self.country.get()].place
        selected = self._current_customer()
        record = {
            "id": selected.get("id") if selected else f"k{len(self.customers) + 1}",
            "naziv": data["naziv"],
            "adresa": data["adresa"],
            "mb": data["mb"],
            "pib": data["pib"],
            "mesto": data["mesto"],
            "format": self.country.get(),
        }
        upsert_customer(self.customers, record)
        self._persist_customers()
        saved = next(
            (
                item
                for item in self.customers
                if item.get("naziv", "").strip().lower() == data["naziv"].strip().lower()
                and item.get("format") == self.country.get()
            ),
            None,
        )
        self._set_combo_to(saved.get("id") if saved else None)
        self.status.config(text=f"Saved customer: {data['naziv']}", fg="#0f7b4c")

    def _delete_customer(self) -> None:
        customer = self._current_customer()
        if customer is None:
            messagebox.showinfo("Customer", "Select a saved customer.")
            return
        confirmed = messagebox.askyesno("Customer", f"Delete {customer.get('naziv', '')}?")
        if not confirmed:
            return
        self.customers = [item for item in self.customers if item.get("id") != customer.get("id")]
        self._persist_customers()
        self._select_customer_for_country(self.country.get())
        self.status.config(text="Customer deleted.", fg=MUTED)

    def _new_customer(self) -> None:
        profile = PROFILES[self.country.get()]
        self._set_combo_to(None)
        self._fill_customer({"naziv": "", "adresa": "", "mb": "", "pib": "", "mesto": profile.place})
        self.buyer_name.focus_set()
        self.status.config(text="Enter the new company, then click Save customer.", fg=MUTED)

    def _persist_customers(self) -> None:
        state = load_state()
        state["kupci_lista"] = self.customers
        state["poslednji_format"] = self.country.get()
        save_state(state)
        self.state = state

    def _remember_number(self) -> None:
        raw = self.number_entry.get().strip()
        if not raw.isdigit() or int(raw) < 1:
            messagebox.showerror("Invoice number", "The invoice number must be a positive whole number.")
            return
        state = load_state()
        state["sledeci_broj"] = int(raw)
        save_state(state)
        self.state = state
        self.status.config(text=f"Next invoice number is {raw}.", fg="#0f7b4c")

    def _apply_default_dates(self) -> None:
        issued = default_issue_date()
        self.due_manual = False
        self._set_entry(self.date_entry, format_date(issued))
        self._set_entry(self.due_entry, format_date(default_due_date(issued)))
        self._set_entry(self.period_entry, period_label(issued))
        self._update_meta()

    def _mark_due_manual(self) -> None:
        self.due_manual = True

    def _set_entry(self, entry: tk.Entry, value: str) -> None:
        state = entry.cget("state")
        entry.config(state="normal")
        entry.delete(0, "end")
        entry.insert(0, value)
        entry.config(state=state)

    def _read_client_form(self) -> dict:
        return {
            "naziv": self.buyer_name.get().strip(),
            "adresa": self.buyer_address.get().strip(),
            "mb": self.buyer_mb.get().strip(),
            "pib": self.buyer_pib.get().strip(),
            "mesto": self.place_entry.get().strip(),
        }

    def _on_date_changed(self) -> None:
        try:
            issued = parse_date(self.date_entry.get())
        except ValueError:
            return
        if not self.due_manual:
            self._set_entry(self.due_entry, format_date(default_due_date(issued)))
        if self.focus_get() is not self.period_entry:
            self._set_entry(self.period_entry, period_label(issued))
        self._update_meta()

    def _update_meta(self) -> None:
        period = self.period_entry.get().strip()
        place = self.place_entry.get().strip()
        if period or place:
            self.meta_var.set(f"Place and month: {period}, {place}")
        else:
            self.meta_var.set("")

    def _add_item(self, placeholder: bool) -> None:
        if len(self.item_rows) >= 20:
            messagebox.showinfo("Services", "An invoice can have at most 20 services.")
            return
        row = ItemRow(self.rows_frame, self._recalc, self._remove_item)
        row.pack(fill="x", pady=3)
        if placeholder:
            row.set_placeholder()
        self.item_rows.append(row)
        self._sync_remove_buttons()
        self._recalc()

    def _remove_item(self, row: ItemRow) -> None:
        if len(self.item_rows) == 1:
            return
        self.item_rows.remove(row)
        row.destroy()
        self._sync_remove_buttons()
        self._recalc()

    def _sync_remove_buttons(self) -> None:
        state = "disabled" if len(self.item_rows) == 1 else "normal"
        for row in self.item_rows:
            row.remove_btn.config(state=state)

    def _recalc(self) -> None:
        profile = PROFILES[self.country.get()]
        total = Decimal("0.00")
        for row in self.item_rows:
            price_text = row.price.get().strip()
            qty_text = row.qty.get().strip() or "1"
            if not price_text:
                row.total.config(text="")
                continue
            try:
                line_total = (parse_money(qty_text) * parse_money(price_text)).quantize(Decimal("0.01"))
            except ValueError:
                row.total.config(text="—")
                continue
            row.total.config(text=f"{line_total:,.2f}")
            total += line_total
        self.total_var.set(money_with_currency(total, profile))

    def _collect(self) -> Invoice:
        raw_number = self.number_entry.get().strip()
        if not raw_number.isdigit() or int(raw_number) < 1:
            raise ValueError("The invoice number must be a positive whole number.")
        issued = parse_date(self.date_entry.get())
        due = parse_date(self.due_entry.get())
        client = self._read_client_form()
        if not client["naziv"] or not client["adresa"]:
            raise ValueError("Enter the customer name and address.")
        if not client["mesto"]:
            raise ValueError("Enter the place of supply.")
        period = self.period_entry.get().strip()
        if not period:
            raise ValueError("Enter the service month.")
        lines: list[Line] = []
        for row in self.item_rows:
            name = row.service_name()
            qty_text = row.qty.get().strip()
            price_text = row.price.get().strip()
            blank_qty = qty_text in ("", "1")
            if not name and not price_text and blank_qty:
                continue
            if not name:
                raise ValueError("Every service needs a name.")
            if not price_text:
                raise ValueError(f"Enter a price for: {name}")
            if not qty_text:
                raise ValueError(f"Enter a quantity for: {name}")
            qty = parse_money(qty_text)
            price = parse_money(price_text)
            if qty <= 0:
                raise ValueError("Quantity must be greater than zero.")
            if price <= 0:
                raise ValueError("Price must be greater than zero.")
            lines.append(Line(name=name, qty=qty, price=price))
        if not lines:
            raise ValueError("Enter at least one service and price.")
        return Invoice(
            number=int(raw_number),
            country=self.country.get(),
            issue_date=issued,
            due_date=due,
            place=client["mesto"],
            period=period,
            buyer=Party(
                name=client["naziv"],
                address=client["adresa"],
                registration=client["mb"],
                tax=client["pib"],
            ),
            lines=lines,
        )

    def _generate(self) -> None:
        self.generate_btn.config(state="disabled")
        self.update_idletasks()
        try:
            invoice = self._collect()
            known = [
                item
                for item in load_state()["istorija"]
                if item.get("broj") == invoice.number and item.get("zemlja") == invoice.country
            ]
            if known:
                country_name = "Slovenia" if invoice.country == "slovenija" else "Serbia"
                again = messagebox.askyesno(
                    "Invoice number",
                    f"Invoice {invoice.number} for {country_name} is already in the list.\n\nCreate another file with the same number?",
                )
                if not again:
                    return
            pdf_path = output_pdf_path(invoice)
            write_pdf(invoice, pdf_path)
            state = remember_invoice(invoice, pdf_path)
            self.state = state
            self.customers = list(state.get("kupci_lista") or [])
            self._set_entry(self.number_entry, str(state["sledeci_broj"]))
            self._refresh_customer_combo()
            self._refresh_history()
            self.status.config(text=f"Saved invoice {invoice.number}", fg="#0f7b4c")
            messagebox.showinfo(
                "Invoice saved",
                f"{pdf_path}\n\n"
                f"Place and month: {invoice.period}, {invoice.place}\n"
                f"Total: {money_with_currency(invoice.total, invoice.profile)}",
            )
            os.startfile(pdf_path)  # noqa: S606 - open the invoice the user just created
        except ValueError as exc:
            self.status.config(text=str(exc), fg="#9b1c1c")
            messagebox.showerror("Check the form", str(exc))
        except Exception as exc:
            self.status.config(text="The invoice could not be saved.", fg="#9b1c1c")
            messagebox.showerror("Error", str(exc))
        finally:
            self.generate_btn.config(state="normal")

    def _refresh_history(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.history_paths.clear()
        self.history_index.clear()
        self.selected_history = None
        for index, item in enumerate(self.state.get("istorija", [])):
            amount = item.get("iznos", "")
            try:
                pretty = f"{Decimal(amount):,.2f} {item.get('valuta', '')}"
            except Exception:
                pretty = f"{amount} {item.get('valuta', '')}"
            country = item.get("zemlja") or item.get("format") or ""
            if country in ("slovenija", "Slovenija", "Slovenia"):
                country_label = "Slovenia"
            elif country in ("srbija", "Srbija", "Serbia"):
                country_label = "Serbia"
            else:
                country_label = str(country)
            iid = self.tree.insert(
                "",
                "end",
                values=(
                    item.get("broj", ""),
                    _iso_to_display(item.get("datum", "")),
                    country_label,
                    item.get("kupac", ""),
                    pretty.strip(),
                ),
                tags=("odd",) if index % 2 else (),
            )
            self.history_paths[iid] = item.get("pdf", "")
            self.history_index[iid] = index

    def _remember_history_click(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.selected_history = row

    def _remember_history_selection(self, _event=None) -> None:
        selected = self.tree.selection()
        if selected:
            self.selected_history = selected[0]

    def _history_selection(self) -> str | None:
        selected = self.tree.selection()
        if selected:
            return selected[0]
        if self.selected_history and self.tree.exists(self.selected_history):
            return self.selected_history
        return None

    def _reload_history(self) -> None:
        state = load_state()
        self.state["istorija"] = list(state.get("istorija") or [])
        self._refresh_history()
        self.status.config(text="Saved invoices refreshed.", fg=MUTED)

    def _open_selected(self) -> None:
        selected = self._history_selection()
        if not selected:
            messagebox.showinfo("Invoices", "Select an invoice from the list.", parent=self)
            return
        path = _existing_invoice_file(self.history_paths.get(selected, ""))
        if path is None:
            messagebox.showerror("PDF", "That PDF file is no longer at the saved location.", parent=self)
            self._reload_history()
            return
        os.startfile(path)  # noqa: S606

    def _delete_selected(self) -> None:
        iid = self._history_selection()
        if not iid:
            messagebox.showinfo("Invoices", "Select an invoice from the list.", parent=self)
            return
        index = self.history_index.get(iid)
        history = list(self.state.get("istorija") or [])
        if index is None or not (0 <= index < len(history)):
            self._reload_history()
            messagebox.showinfo(
                "Invoices",
                "That invoice is no longer in the list. The list was refreshed.",
                parent=self,
            )
            return
        item = history[index]
        number = item.get("broj", "")
        customer = item.get("kupac", "")
        if not messagebox.askyesno(
            "Delete invoice",
            f"Delete invoice {number} for {customer}?\n\nThe saved PDF will be removed.",
            parent=self,
        ):
            return
        history.pop(index)
        state = load_state()
        state["istorija"] = history
        save_state(state)
        self.state["istorija"] = history
        problems = _delete_saved_files(item)
        self._refresh_history()
        self.update_idletasks()
        if problems:
            self.status.config(text=f"Invoice {number} was removed from the list.", fg="#9b1c1c")
            messagebox.showwarning(
                "Delete invoice",
                "The invoice was removed from the list, but a file is still open and could not be deleted:\n\n"
                + "\n".join(problems),
                parent=self,
            )
            return
        self.status.config(text=f"Deleted invoice {number}.", fg="#0f7b4c")

    def _open_folder(self) -> None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(OUTPUT_DIR)  # noqa: S606

    def _on_close(self) -> None:
        try:
            state = load_state()
            state["kupci_lista"] = self.customers
            state["poslednji_format"] = self.country.get()
            save_state(state)
        finally:
            self.destroy()


def _invoice_file_candidates(raw: str) -> list[Path]:
    if not raw:
        return []
    paths = [Path(raw)]
    updated = english_stored_path(raw)
    if os.path.normcase(updated) != os.path.normcase(raw):
        paths.append(Path(updated))
    return paths


def _existing_invoice_file(raw: str) -> Path | None:
    for path in _invoice_file_candidates(raw):
        if path.is_file():
            return path
    return None


def _delete_saved_files(item: dict) -> list[str]:
    problems: list[str] = []
    candidates: list[Path] = []
    for key in ("pdf", "excel"):
        raw = str(item.get(key) or "")
        candidates.extend(_invoice_file_candidates(raw))
        pdf = Path(raw) if raw else None
        if pdf is not None and pdf.suffix.lower() == ".pdf":
            candidates.append(pdf.with_suffix(".xlsx"))
            updated = english_stored_path(raw)
            candidates.append(Path(updated).with_suffix(".xlsx"))
    seen: set[str] = set()
    for path in candidates:
        if not path.suffix:
            continue
        key = os.path.normcase(str(path))
        if key in seen or not path.is_file():
            continue
        seen.add(key)
        try:
            path.unlink()
        except OSError as exc:
            problems.append(f"{path.name} ({exc.strerror or exc})")
    return problems


def _iso_to_display(value: str) -> str:
    try:
        return format_date(date.fromisoformat(value))
    except ValueError:
        return value


def main() -> None:
    _enable_dpi()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
