"""
widgets.py  –  Komponen UI reusable
- DatePickerWidget  : input tanggal dengan popup kalender
- PeriodeSelector   : dropdown pilih periode aktif
- MonthYearPicker   : pilih bulan & tahun
"""
import tkinter as tk
from tkinter import ttk
from datetime import date, datetime
import calendar
from helpers import BULAN_NAMA, this_year, this_month

C_ERR  = "#EF4444"
C_OK   = "#059669"

CAL_HDR    = "#1E3A5F"
CAL_TODAY  = "#2B6CB0"
CAL_SEL    = "#1E3A5F"
CAL_HOVER  = "#DBEAFE"
CAL_WK     = "#94A3B8"
CAL_WEEKEND= "#EF4444"
CAL_BG     = "#FFFFFF"
CAL_BORDER = "#E2E8F0"


class CalendarPopup(tk.Toplevel):
    """
    Popup kalender - style gambar 1:
    Title bar 'Pilih Tanggal' + nav bar ◀▶ + dropdown tahun + footer Hari ini.
    """
    NAV_BG   = "#2B6CB0"
    HEAD_BG  = "#EBF8FF"
    SEL_BG   = "#2B6CB0"
    SEL_FG   = "white"
    TODAY_FG = "#C05621"
    DAY_FG   = "#1E293B"
    HOVER_BG = "#EBF8FF"

    def __init__(self, parent, current_date=None, on_select=None):
        super().__init__(parent)
        self.title("Pilih Tanggal")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.configure(bg="white")
        self._on_select = on_select
        self._alive     = True

        today = date.today()
        if current_date:
            try:
                self._cur = datetime.strptime(current_date, "%Y-%m-%d").date()
            except Exception:
                self._cur = today
        else:
            self._cur = today

        self._view_year  = self._cur.year
        self._view_month = self._cur.month
        self._today      = today
        self._yv         = tk.StringVar(value=str(self._view_year))

        self._main = tk.Frame(self, bg="white", padx=6, pady=6)
        self._main.pack(fill="both", expand=True)

        self._draw()
        self.bind("<Escape>", lambda e: self._close())
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _fmt(self, d):
        return f"{d.day:02d} {BULAN_NAMA[d.month]} {d.year}"

    def _draw(self):
        for w in self._main.winfo_children():
            w.destroy()

        # Nav bar
        nav = tk.Frame(self._main, bg=self.NAV_BG)
        nav.pack(fill="x", pady=(0, 4))
        tk.Button(nav, text="◀", bg=self.NAV_BG, fg="white",
                  font=("Arial", 10, "bold"), relief="flat", bd=0,
                  cursor="hand2", activebackground="#1A4F8A",
                  command=self._prev_month).pack(side="left", padx=4, pady=4)
        tk.Label(nav, text=f"{BULAN_NAMA[self._view_month]} {self._view_year}",
                 font=("Arial", 10, "bold"), bg=self.NAV_BG,
                 fg="white").pack(side="left", expand=True)
        tk.Button(nav, text="▶", bg=self.NAV_BG, fg="white",
                  font=("Arial", 10, "bold"), relief="flat", bd=0,
                  cursor="hand2", activebackground="#1A4F8A",
                  command=self._next_month).pack(side="right", padx=4, pady=4)

        # Dropdown tahun
        yf = tk.Frame(self._main, bg="white")
        yf.pack(fill="x", pady=(0, 4))
        tk.Label(yf, text="Tahun:", font=("Arial", 8),
                 bg="white", fg="#64748B").pack(side="left")
        self._yv.set(str(self._view_year))
        years = [str(y) for y in range(self._today.year - 10,
                                       self._today.year + 11)]
        cb = ttk.Combobox(yf, textvariable=self._yv, values=years,
                          width=6, state="readonly", font=("Arial", 8))
        cb.pack(side="left", padx=4)
        cb.bind("<<ComboboxSelected>>", self._on_year_change)

        # Header hari
        hdr_f = tk.Frame(self._main, bg=self.HEAD_BG)
        hdr_f.pack(fill="x")
        for i, d in enumerate(["Sen","Sel","Rab","Kam","Jum","Sab","Min"]):
            fg = "#C05621" if i == 6 else "#1E293B"
            tk.Label(hdr_f, text=d, font=("Arial", 8, "bold"),
                     bg=self.HEAD_BG, fg=fg, width=4,
                     anchor="center").pack(side="left", padx=1)

        # Grid hari
        grid_f = tk.Frame(self._main, bg="white")
        grid_f.pack()
        for week in calendar.monthcalendar(self._view_year, self._view_month):
            row_f = tk.Frame(grid_f, bg="white")
            row_f.pack()
            for col_i, day in enumerate(week):
                if day == 0:
                    tk.Label(row_f, text="", width=4, font=("Arial", 9),
                             bg="white").pack(side="left", padx=1, pady=1)
                    continue
                d = date(self._view_year, self._view_month, day)
                is_sel   = (d == self._cur)
                is_today = (d == self._today)
                is_sun   = (col_i == 6)
                bg  = self.SEL_BG if is_sel else "white"
                fg  = self.SEL_FG if is_sel else                       (self.TODAY_FG if is_today else
                       ("#C05621" if is_sun else self.DAY_FG))
                fnt = ("Arial", 9, "bold") if (is_sel or is_today)                       else ("Arial", 9)
                btn = tk.Button(row_f, text=str(day), width=3, font=fnt,
                                bg=bg, fg=fg, relief="flat", bd=0,
                                cursor="hand2",
                                command=lambda dd=d: self._select_date(dd))
                btn.pack(side="left", padx=1, pady=1)
                if not is_sel:
                    btn.bind("<Enter>",
                             lambda e, b=btn: b.config(bg=self.HOVER_BG))
                    btn.bind("<Leave>",
                             lambda e, b=btn, bg_=bg: b.config(bg=bg_))

        # Footer
        tk.Button(self._main,
                  text=f"📅 Hari ini: {self._fmt(self._today)}",
                  font=("Arial", 8), bg="#F0FFF4", fg="#276749",
                  relief="flat", cursor="hand2",
                  command=lambda: self._select_date(self._today)
                  ).pack(fill="x", pady=(6, 0))

    def _prev_month(self):
        if self._view_month == 1:
            self._view_month, self._view_year = 12, self._view_year - 1
        else:
            self._view_month -= 1
        self._draw()

    def _next_month(self):
        if self._view_month == 12:
            self._view_month, self._view_year = 1, self._view_year + 1
        else:
            self._view_month += 1
        self._draw()

    def _on_year_change(self, _=None):
        try:
            self._view_year = int(self._yv.get())
            self._draw()
        except ValueError:
            pass

    def _select_date(self, d):
        self._cur = d
        if self._on_select:
            self._on_select(d)
        self._close()

    # Compat aliases
    def _go_today(self):      self._select_date(date.today())
    def _select(self, day):   self._select_date(date(self._view_year, self._view_month, day))
    def _on_month_change(self, _=None): pass

    def _close(self):
        if not self._alive:
            return
        self._alive = False
        try:
            self.grab_release()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    def position_near(self, widget):
        self.update_idletasks()
        wx = widget.winfo_rootx()
        wy = widget.winfo_rooty() + widget.winfo_height() + 2
        sw = self.winfo_screenwidth()
        pw = self.winfo_reqwidth()
        if wx + pw > sw:
            wx = sw - pw - 4
        self.geometry(f"+{wx}+{wy}")
        self.lift()
        self.focus_set()
        self.after(50, self._activate_grab)

    def _activate_grab(self):
        if not self._alive:
            return
        try:
            self.grab_set()
        except Exception:
            pass
        self.bind("<ButtonPress-1>", self._check_outside)

    def _check_outside(self, event):
        if not self._alive:
            return
        try:
            w = event.widget
            while w is not None:
                if str(w) == str(self):
                    return
                try:
                    w = w.nametowidget(w.winfo_parent())
                except Exception:
                    break
        except Exception:
            pass
        self._close()


class DatePickerWidget(tk.Frame):
    """
    Input tanggal dengan tombol kalender popup.
    Interface sama persis dengan versi lama:
      .get()       → 'YYYY-MM-DD'
      .set(str)    → dari 'YYYY-MM-DD'
      ._set_today()
    """
    def __init__(self, parent, label="Tanggal", default=None, **kw):
        bg = kw.pop("bg", parent.cget("bg"))
        super().__init__(parent, bg=bg, **kw)
        self._popup = None

        if label:
            tk.Label(self, text=label, font=("Arial", 9),
                     bg=bg, fg="#64748B").grid(row=0, column=0, columnspan=2,
                                               sticky="w", pady=(2, 1))

        today = date.today()
        self._date = today

        self._var = tk.StringVar(value=self._fmt_display(today))

        self._entry = ttk.Entry(self, textvariable=self._var, width=16,
                                state="readonly", font=("Arial", 9))
        self._entry.grid(row=1, column=0, padx=(0, 2))

        self._btn = tk.Button(self, text="📅", font=("Arial", 10),
                               bg="#EBF8FF", fg="#2B6CB0", bd=0,
                               padx=4, pady=0, cursor="hand2",
                               activebackground="#DBEAFE",
                               command=self._open_calendar)
        self._btn.grid(row=1, column=1)

        if default:
            self.set(default)

    def _fmt_display(self, d):
        return f"{d.day:02d} {BULAN_NAMA[d.month]} {d.year}"

    def _open_calendar(self):
        if self._popup is not None:
            try:
                if self._popup._alive:
                    self._popup._close()
                    self._popup = None
                    return
            except Exception:
                pass
            self._popup = None

        popup = CalendarPopup(self, current_date=self.get() or None,
                              on_select=self._from_calendar)
        popup.position_near(self._btn)
        self._popup = popup

    def _from_calendar(self, d: date):
        self._date = d
        self._var.set(self._fmt_display(d))

    def _set_today(self):
        t = date.today()
        self._date = t
        self._var.set(self._fmt_display(t))

    def get(self):
        return self._date.strftime("%Y-%m-%d") if self._date else ""

    def set(self, date_str):
        if not date_str:
            return
        try:
            if isinstance(date_str, str):
                d = datetime.strptime(date_str, "%Y-%m-%d").date()
            elif isinstance(date_str, date):
                d = date_str
            else:
                return
            self._date = d
            self._var.set(self._fmt_display(d))
        except Exception:
            pass

    def is_valid(self):
        return self._date is not None


class MonthYearPicker(tk.Frame):
    def __init__(self, parent, label="Periode", default_year=None,
                 default_month=None, show_all=True, **kw):
        bg = kw.pop("bg", parent.cget("bg"))
        super().__init__(parent, bg=bg, **kw)
        tk.Label(self, text=label, font=("Arial", 9), bg=bg,
                 fg="#64748B").grid(row=0, column=0, columnspan=2, sticky="w", pady=(2, 1))

        months = (["Semua Bulan"] if show_all else []) + [BULAN_NAMA[i] for i in range(1, 13)]
        years  = [str(y) for y in range(2020, this_year() + 3)]

        self._mv = tk.StringVar()
        self._yv = tk.StringVar()

        self._cm = ttk.Combobox(self, textvariable=self._mv, values=months, width=14, state="readonly")
        self._cy = ttk.Combobox(self, textvariable=self._yv, values=years,  width=6,  state="readonly")
        self._cm.grid(row=1, column=0, padx=(0, 6))
        self._cy.grid(row=1, column=1)

        m = default_month or this_month()
        y = default_year  or this_year()
        self._mv.set("Semua Bulan" if show_all else BULAN_NAMA[m])
        self._yv.set(str(y))

    def get_bulan(self):
        v = self._mv.get()
        for k, n in BULAN_NAMA.items():
            if n == v:
                return k
        return None

    def get_tahun(self):
        try:
            return int(self._yv.get())
        except Exception:
            return this_year()

    def set(self, bulan, tahun):
        self._yv.set(str(tahun))
        self._mv.set(BULAN_NAMA.get(bulan, "Semua Bulan"))


class PeriodeSelector(tk.Frame):
    def __init__(self, parent, conn, label="Periode", include_all=True, **kw):
        bg = kw.pop("bg", parent.cget("bg"))
        super().__init__(parent, bg=bg, **kw)
        tk.Label(self, text=label, font=("Arial", 9), bg=bg,
                 fg="#64748B").grid(row=0, column=0, sticky="w", pady=(2, 1))
        self._var = tk.StringVar()
        self._map = {}
        rows = conn.execute("SELECT id,nama,tahun FROM periode ORDER BY tahun DESC,id DESC").fetchall()
        opts = (["— Semua Periode —"] if include_all else []) + \
               [f"{r['nama']} ({r['tahun']})" for r in rows]
        for r in rows:
            self._map[f"{r['nama']} ({r['tahun']})"] = r["id"]
        self._cb = ttk.Combobox(self, textvariable=self._var, values=opts,
                                width=22, state="readonly")
        self._cb.grid(row=1, column=0)
        if opts:
            self._cb.current(0)

    def get_id(self):
        return self._map.get(self._var.get(), None)

    def bind_change(self, fn):
        self._var.trace_add("write", lambda *_: fn())