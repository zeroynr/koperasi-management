"""
rekap.py
Halaman rekap & export Excel — dengan filter periode, jenis simpanan,
dan tampilan export yang minimalis & rapi.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from database import get_conn
from helpers import fmt_rp, JENIS_LABEL, JENIS_LIST, BULAN_NAMA, this_year
from export_excel import (export_rekap_lengkap, export_simpanan,
                          export_simpanan_perbulan,
                          export_pinjaman_angsuran,
                          export_anggota, export_neraca,
                          export_kartu_pinjaman)
from pages.base_page import BasePage, C_BG, C_WHITE, C_BLUE, C_DARK, C_GRAY

# ── Warna kartu export ────────────────────────────────────
CARD_COLORS = {
    "lengkap":      ("#1E3A5F", "#EBF8FF", "📊"),
    "simpanan":     ("#276749", "#F0FFF4", "💰"),
    "pin_angsuran": ("#C05621", "#FFFAF0", "📋"),
    "kartu_pin":    ("#2B6CB0", "#EBF8FF", "🪪"),
    "anggota":      ("#553C9A", "#F3E8FF", "👥"),
    "neraca":       ("#1A365D", "#E0F2FE", "🏦"),
}


class RekapPage(BasePage):
    def __init__(self, parent):
        super().__init__(parent)
        self._build()
        self.refresh()

    # ─────────────────────────────────────────────────────
    def _build(self):
        # ── Judul ─────────────────────────────────────────
        hdr = tk.Frame(self, bg=C_BG)
        hdr.pack(fill="x", padx=16, pady=(14, 6))
        tk.Label(hdr, text="Rekap & Export Excel",
                 font=("Arial", 13, "bold"), bg=C_BG, fg=C_DARK).pack(anchor="w")
        tk.Label(hdr, text="Pilih filter lalu klik tombol export yang diinginkan",
                 font=("Arial", 9), bg=C_BG, fg=C_GRAY).pack(anchor="w")

        # ── Panel filter ──────────────────────────────────
        filter_outer = tk.Frame(self, bg=C_BG)
        filter_outer.pack(fill="x", padx=16, pady=(0, 8))

        filter_card = tk.Frame(filter_outer, bg=C_WHITE,
                               highlightbackground="#E2E8F0", highlightthickness=1)
        filter_card.pack(fill="x")

        # Judul filter
        f_title = tk.Frame(filter_card, bg="#F8FAFC")
        f_title.pack(fill="x")
        tk.Label(f_title, text="🔍  Filter Export", font=("Arial", 9, "bold"),
                 bg="#F8FAFC", fg=C_DARK).pack(side="left", padx=14, pady=8)
        tk.Frame(filter_card, bg="#E2E8F0", height=1).pack(fill="x")

        # Baris filter
        frow = tk.Frame(filter_card, bg=C_WHITE, pady=10)
        frow.pack(fill="x", padx=14)

        # Filter Periode
        fp = tk.Frame(frow, bg=C_WHITE)
        fp.pack(side="left", padx=(0, 20))
        tk.Label(fp, text="Periode", font=("Arial", 8, "bold"),
                 bg=C_WHITE, fg="#64748B").pack(anchor="w")
        self._v_periode = tk.StringVar(value="— Semua Periode —")
        self._periode_map = {}
        self._cb_periode = ttk.Combobox(fp, textvariable=self._v_periode,
                                         width=22, state="readonly")
        self._cb_periode.pack(anchor="w", pady=(2, 0))
        self._cb_periode.bind("<<ComboboxSelected>>", lambda _: self._refresh_rekap())

        # Filter Tahun
        ft = tk.Frame(frow, bg=C_WHITE)
        ft.pack(side="left", padx=(0, 20))
        tk.Label(ft, text="Tahun", font=("Arial", 8, "bold"),
                 bg=C_WHITE, fg="#64748B").pack(anchor="w")
        self._v_tahun = tk.StringVar(value=str(this_year()))
        years = ["Semua"] + [str(y) for y in range(2020, this_year() + 3)]
        cb_tahun = ttk.Combobox(ft, textvariable=self._v_tahun, values=years,
                     width=8, state="readonly")
        cb_tahun.pack(anchor="w", pady=(2, 0))
        cb_tahun.bind("<<ComboboxSelected>>", lambda _: self._refresh_rekap())

        # Filter Jenis Simpanan
        fj = tk.Frame(frow, bg=C_WHITE)
        fj.pack(side="left", padx=(0, 20))
        tk.Label(fj, text="Jenis Simpanan (untuk export simpanan)",
                 font=("Arial", 8, "bold"), bg=C_WHITE, fg="#64748B").pack(anchor="w")
        self._v_jenis = tk.StringVar(value="Semua Jenis")
        jenis_opts = ["Semua Jenis"] + [JENIS_LABEL[j] for j in JENIS_LIST]
        ttk.Combobox(fj, textvariable=self._v_jenis, values=jenis_opts,
                     width=20, state="readonly").pack(anchor="w", pady=(2, 0))

        # Tombol reset filter
        fr = tk.Frame(frow, bg=C_WHITE)
        fr.pack(side="left", padx=(0, 0))
        tk.Label(fr, text=" ", font=("Arial", 8), bg=C_WHITE).pack()
        tk.Button(fr, text="↺ Reset", font=("Arial", 8),
                  bg="#F1F5F9", fg="#64748B", bd=0, padx=8, pady=4,
                  cursor="hand2", activebackground="#E2E8F0",
                  command=self._reset_filter).pack(pady=(2, 0))

        # ── Kartu export ──────────────────────────────────
        exp_frame = tk.Frame(self, bg=C_BG)
        exp_frame.pack(fill="x", padx=16, pady=(0, 10))

        tk.Label(exp_frame, text="Pilih Jenis Export", font=("Arial", 9, "bold"),
                 bg=C_BG, fg=C_DARK).pack(anchor="w", pady=(0, 6))

        cards_row = tk.Frame(exp_frame, bg=C_BG)
        cards_row.pack(fill="x")

        exports = [
            ("lengkap",      "Rekap Lengkap",        "Semua data\n(Anggota, Simpanan,\nPinjaman, Angsuran, Neraca)", self._exp_lengkap),
            ("simpanan",     "Simpanan",              "Matriks per-bulan\n+ Detail list\n(sesuai filter jenis)", self._exp_simpanan),
            ("pin_angsuran", "Pinjaman &\nAngsuran", "4 sheet:\nDaftar Pinjaman\nRiwayat Angsuran\nMatriks Per Bulan\nRekap Per Tahun", self._exp_pin_angsuran),
            ("kartu_pin",    "Kartu Pinjaman",        "Sheet per anggota\nRiwayat angsuran\nper bulan (Jan-Des)\nHistoris merah", self._exp_kartu_pin),
            ("anggota",      "Data Anggota",          "Daftar lengkap\nanggota koperasi", self._exp_anggota),
            ("neraca",       "Neraca",                "Ringkasan\nkeuangan\n& aset", self._exp_neraca),
        ]

        for key, title, desc, cmd in exports:
            fg_c, bg_c, icon = CARD_COLORS[key]
            self._make_export_card(cards_row, icon, title, desc, cmd, fg_c, bg_c)

        # ── Rekap statistik bawah ──────────────────────────
        mid = tk.Frame(self, bg=C_BG)
        mid.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        mid.columnconfigure(0, weight=1)
        mid.columnconfigure(1, weight=1)

        self._smp_card = tk.Frame(mid, bg=C_WHITE,
                                   highlightbackground="#E2E8F0", highlightthickness=1)
        self._smp_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=4)

        self._pin_card = tk.Frame(mid, bg=C_WHITE,
                                   highlightbackground="#E2E8F0", highlightthickness=1)
        self._pin_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=4)

    def _make_export_card(self, parent, icon, title, desc, cmd, fg_c, bg_c):
        """Buat kartu export yang minimalis."""
        card = tk.Frame(parent, bg=bg_c, cursor="hand2",
                        highlightbackground=fg_c, highlightthickness=1)
        card.pack(side="left", padx=(0, 8), ipadx=8, ipady=8, fill="y")

        # Icon
        tk.Label(card, text=icon, font=("Arial", 20), bg=bg_c).pack(pady=(8, 2))

        # Judul
        tk.Label(card, text=title, font=("Arial", 9, "bold"),
                 bg=bg_c, fg=fg_c, justify="center").pack(padx=12)

        # Deskripsi
        tk.Label(card, text=desc, font=("Arial", 7),
                 bg=bg_c, fg="#64748B", justify="center").pack(padx=10, pady=(2, 6))

        # Tombol export
        btn = tk.Button(card, text="⬇ Export", font=("Arial", 8, "bold"),
                        bg=fg_c, fg="white", bd=0, padx=10, pady=4,
                        cursor="hand2", relief="flat", command=cmd,
                        activebackground="#1A365D", activeforeground="white")
        btn.pack(pady=(0, 8), padx=12, fill="x")

        # Hover efek
        for widget in [card] + list(card.winfo_children()):
            widget.bind("<Enter>", lambda e, c=card, bg=bg_c, fc=fg_c:
                        self._card_hover(c, True, bg, fc))
            widget.bind("<Leave>", lambda e, c=card, bg=bg_c, fc=fg_c:
                        self._card_hover(c, False, bg, fc))

    def _card_hover(self, card, enter, bg_c, fg_c):
        try:
            thickness = 2 if enter else 1
            card.config(highlightthickness=thickness)
        except Exception:
            pass

    # ── Data & refresh ────────────────────────────────────
    def refresh(self):
        self._load_periode_opts()
        self._refresh_rekap()

    def _load_periode_opts(self):
        conn = get_conn()
        try:
            rows = conn.execute(
                "SELECT id, nama, tahun, bulan_mulai, bulan_akhir, tgl_mulai, tgl_akhir "
                "FROM periode ORDER BY tahun DESC, id DESC"
            ).fetchall()
        finally:
            conn.close()
        self._periode_map = {"— Semua Periode —": None}
        opts = ["— Semua Periode —"]
        for r in rows:
            from calendar import monthrange
            from datetime import datetime as _dt
            from helpers import BULAN_SHORT

            def _to_label_date(tgl_str, fallback_day, fallback_m, fallback_y):
                if tgl_str:
                    try:
                        d = _dt.strptime(tgl_str, "%Y-%m-%d")
                        short = BULAN_SHORT.get(d.month, str(d.month))
                        return f"{d.day:02d} {short} {d.year}"
                    except Exception:
                        pass
                short = BULAN_SHORT.get(fallback_m, str(fallback_m))
                return f"{fallback_day:02d} {short} {fallback_y}"

            hari_akhir = monthrange(r["tahun"], r["bulan_akhir"])[1]
            tgl_m = _to_label_date(r["tgl_mulai"],    1,          r["bulan_mulai"], r["tahun"])
            tgl_a = _to_label_date(r["tgl_akhir"],  hari_akhir,   r["bulan_akhir"], r["tahun"])
            label = f"{r['nama']}  [{tgl_m} – {tgl_a}]"
            opts.append(label)
            self._periode_map[label] = r["id"]
        self._cb_periode["values"] = opts
        if self._v_periode.get() not in opts:
            self._v_periode.set("— Semua Periode —")

    def _refresh_rekap(self):
        self._build_smp()
        self._build_pin()

    def _reset_filter(self):
        self._v_periode.set("— Semua Periode —")
        self._v_tahun.set(str(this_year()))
        self._v_jenis.set("Semua Jenis")
        self._refresh_rekap()

    def _get_filter(self):
        """Return (periode_id, tahun, jenis_key)."""
        pid   = self._periode_map.get(self._v_periode.get(), None)
        tahun_str = self._v_tahun.get()
        tahun = None if tahun_str == "Semua" else int(tahun_str)
        jenis_label = self._v_jenis.get()
        jenis = next((k for k, v in JENIS_LABEL.items() if v == jenis_label), None)
        return pid, tahun, jenis

    def _build_smp(self):
        for w in self._smp_card.winfo_children():
            w.destroy()
        pid, tahun, _ = self._get_filter()
        filter_txt = self._filter_label(pid, tahun)

        tk.Label(self._smp_card, text="Rekap Simpanan per Jenis",
                 font=("Arial", 10, "bold"), bg=C_WHITE, fg=C_DARK).pack(
            anchor="w", padx=12, pady=(10, 2))
        if filter_txt:
            tk.Label(self._smp_card, text=filter_txt, font=("Arial", 8),
                     bg=C_WHITE, fg="#64748B").pack(anchor="w", padx=12, pady=(0, 4))

        conn = get_conn()
        try:
            anggota = conn.execute("SELECT * FROM anggota ORDER BY id").fetchall()
            cols = ("No. Anggota", "Nama") + tuple(JENIS_LABEL[j] for j in JENIS_LIST) + ("Total",)
            tv = ttk.Treeview(self._smp_card, columns=cols, show="headings", height=11)
            widths = [90, 150] + [90]*5 + [100]
            for col, w in zip(cols, widths):
                tv.heading(col, text=col)
                tv.column(col, width=w,
                          anchor="w" if col in ("No. Anggota","Nama") else "e")
            tv.tag_configure("even", background="#EBF8FF")
            tv.tag_configure("odd",  background="white")

            grand_total = 0
            for i, a in enumerate(anggota):
                vals = {}
                for j in JENIS_LIST:
                    q = "SELECT COALESCE(SUM(jumlah),0) FROM simpanan WHERE anggota_id=? AND jenis=?"
                    params = [a["id"], j]
                    if pid:
                        q += " AND periode_id=?"; params.append(pid)
                    elif tahun:
                        q += " AND tahun=?"; params.append(tahun)
                    vals[j] = conn.execute(q, params).fetchone()[0]
                tot = sum(vals.values())
                grand_total += tot
                tag = "even" if i % 2 == 0 else "odd"
                tv.insert("", "end", tags=(tag,), values=(
                    a["no_anggota"], a["nama"],
                    fmt_rp(vals["pokok"]), fmt_rp(vals["wajib"]),
                    fmt_rp(vals["sukarela"]), fmt_rp(vals["khusus"]),
                    fmt_rp(vals["hariraya"]), fmt_rp(tot)
                ))
            # Total row
            total_vals = []
            for j in JENIS_LIST:
                q = "SELECT COALESCE(SUM(jumlah),0) FROM simpanan WHERE jenis=?"
                params = [j]
                if pid:
                    q += " AND periode_id=?"; params.append(pid)
                elif tahun:
                    q += " AND tahun=?"; params.append(tahun)
                total_vals.append(conn.execute(q, params).fetchone()[0])
            tv.insert("", "end", values=(
                "", "TOTAL",
                *[fmt_rp(v) for v in total_vals],
                fmt_rp(sum(total_vals))
            ))
        finally:
            conn.close()

        vsb = ttk.Scrollbar(self._smp_card, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", padx=(0,4))
        tv.pack(fill="both", expand=True, padx=8, pady=(0,10))

    def _build_pin(self):
        for w in self._pin_card.winfo_children():
            w.destroy()
        pid, tahun, _ = self._get_filter()
        filter_txt = self._filter_label(pid, tahun)

        # Header
        hdr_f = tk.Frame(self._pin_card, bg=C_WHITE)
        hdr_f.pack(fill="x", padx=12, pady=(10, 0))
        tk.Label(hdr_f, text="Rekap Pinjaman",
                 font=("Arial", 10, "bold"), bg=C_WHITE, fg=C_DARK).pack(side="left")
        if filter_txt:
            tk.Label(self._pin_card, text=filter_txt, font=("Arial", 8),
                     bg=C_WHITE, fg="#64748B").pack(anchor="w", padx=12, pady=(0, 2))

        conn = get_conn()
        try:
            # Query realtime: selalu ambil data fresh sesuai filter saat ini
            conditions = []
            params = []
            if pid:
                conditions.append("p.periode_id=?")
                params.append(pid)
            if tahun:
                conditions.append("p.tahun=?")
                params.append(tahun)
            where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

            q = f"""
                SELECT a.no_anggota, a.nama, p.jumlah, p.jangka, p.bunga,
                       p.tgl, p.status,
                       (SELECT COALESCE(SUM(ag.jumlah),0)
                        FROM angsuran ag WHERE ag.pinjaman_id=p.id) as bayar,
                       p.id as pin_id
                FROM pinjaman p JOIN anggota a ON p.anggota_id=a.id
                {where} ORDER BY p.tgl DESC
            """
            rows = conn.execute(q, params).fetchall()

            # Hitung ringkasan
            total_pinjaman = sum(r["jumlah"] for r in rows)
            total_bayar    = sum(r["bayar"]  for r in rows)
            total_sisa     = sum(max(0, r["jumlah"] - r["bayar"]) for r in rows)
            jml_aktif      = sum(1 for r in rows if r["status"] == "aktif")
            jml_lunas      = sum(1 for r in rows if r["status"] == "lunas")
        finally:
            conn.close()

        # Info ringkasan singkat
        info_f = tk.Frame(self._pin_card, bg="#FFF7ED")
        info_f.pack(fill="x", padx=8, pady=(2, 4))
        for lbl, val, clr in [
            ("Total Pinjaman", fmt_rp(total_pinjaman), "#92400E"),
            ("Terbayar",       fmt_rp(total_bayar),    "#276749"),
            ("Sisa",           fmt_rp(total_sisa),     "#C05621"),
            (f"Aktif: {jml_aktif} | Lunas: {jml_lunas}", "", "#64748B"),
        ]:
            tk.Label(info_f, text=f"{lbl}  {val}" if val else lbl,
                     font=("Arial", 8), bg="#FFF7ED", fg=clr).pack(
                side="left", padx=8, pady=3)

        # Treeview
        cols = ("No. Anggota","Nama","Jml Pinjaman","Terbayar","Sisa","Status")
        tv = ttk.Treeview(self._pin_card, columns=cols, show="headings", height=9)
        for col, w, a in zip(cols, [90,150,110,110,110,70],
                              ["center","w","e","e","e","center"]):
            tv.heading(col, text=col, command=lambda c=col: self._sort_pin(tv, c))
            tv.column(col, width=w, anchor=a)
        tv.tag_configure("even",  background="#EBF8FF")
        tv.tag_configure("odd",   background="white")
        tv.tag_configure("lunas", foreground="#276749", font=("Arial", 9, "italic"))
        tv.tag_configure("aktif", foreground="#C05621")

        for i, r in enumerate(rows):
            sisa = max(0, r["jumlah"] - r["bayar"])
            base_tag = "even" if i % 2 == 0 else "odd"
            st_tag   = "lunas" if r["status"] == "lunas" else "aktif"
            tv.insert("", "end", tags=(base_tag, st_tag), values=(
                r["no_anggota"], r["nama"],
                fmt_rp(r["jumlah"]), fmt_rp(r["bayar"]),
                fmt_rp(sisa), r["status"].upper()
            ))

        # Baris total
        tv.insert("", "end", tags=("total",), values=(
            "", "TOTAL",
            fmt_rp(total_pinjaman), fmt_rp(total_bayar),
            fmt_rp(total_sisa), f"{len(rows)} pinjaman"
        ))
        tv.tag_configure("total", background="#FED7AA",
                         font=("Arial", 9, "bold"), foreground="#7B341E")

        vsb = ttk.Scrollbar(self._pin_card, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", padx=(0, 4))
        tv.pack(fill="both", expand=True, padx=8, pady=(0, 10))

    def _sort_pin(self, tv, col):
        """Sort treeview pinjaman by column."""

    def _filter_label(self, pid, tahun):
        parts = []
        if pid:
            for k, v in self._periode_map.items():
                if v == pid:
                    parts.append(f"Periode: {k}")
        if tahun:
            parts.append(f"Tahun: {tahun}")
        return "  |  ".join(parts) if parts else ""

    # ── Export helpers ────────────────────────────────────
    def _ask_save(self, default_name):
        return filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files","*.xlsx")],
            initialfile=default_name,
            title="Simpan File Excel"
        )

    def _do_export(self, fn, default_name, **kwargs):
        path = self._ask_save(default_name)
        if not path:
            return
        conn = get_conn()
        try:
            fn(conn, path, **kwargs)
            messagebox.showinfo("✅ Berhasil",
                                f"File berhasil disimpan:\n{os.path.basename(path)}")
            try:
                os.startfile(path)
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Gagal Export", str(e))
        finally:
            conn.close()

    def _exp_lengkap(self):
        pid, tahun, _ = self._get_filter()
        self._do_export(export_rekap_lengkap, "Rekap_Koperasi_Lengkap.xlsx",
                        periode_id=pid, tahun=tahun)

    def _exp_simpanan(self):
        pid, tahun, jenis = self._get_filter()
        jenis_name = JENIS_LABEL.get(jenis, "Semua") if jenis else "Semua"
        tahun_name = str(tahun) if tahun else str(__import__("helpers").this_year())
        fname = f"Rekap_Simpanan_{jenis_name}_{tahun_name}.xlsx".replace(" ","_")
        self._do_export(export_simpanan, fname,
                        jenis=jenis, periode_id=pid, tahun=tahun)

    def _exp_pin_angsuran(self):
        pid, tahun, _ = self._get_filter()
        from helpers import this_year
        tahun_name = str(tahun) if tahun else str(this_year())
        fname = f"Rekap_Pinjaman_Angsuran_{tahun_name}.xlsx"
        self._do_export(export_pinjaman_angsuran, fname,
                        periode_id=pid, tahun=tahun)

    def _exp_anggota(self):
        self._do_export(export_anggota, "Data_Anggota.xlsx")

    def _exp_kartu_pin(self):
        from helpers import this_year
        _, tahun, _ = self._get_filter()
        thn = tahun if tahun else this_year()
        fname = f"Kartu_Pinjaman_{thn}.xlsx"
        self._do_export(export_kartu_pinjaman, fname, tahun=thn)

    def _exp_neraca(self):
        pid, tahun, _ = self._get_filter()
        self._do_export(export_neraca, "Neraca_Koperasi.xlsx",
                        periode_id=pid, tahun=tahun)