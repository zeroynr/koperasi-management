"""
kas.py  — Buku Kas Masuk (READ-ONLY + Rekap Terintegrasi)
Data otomatis masuk dari modul Simpanan dan Angsuran.
3 tab:
  Tab 1 : Buku Kas Masuk  (semua transaksi masuk per bulan)
  Tab 2 : Rekap Simpanan  (SW / SK / Pokok yang masuk bulan ini)
  Tab 3 : Rekap Angsuran  (siapa yang bayar angsuran bulan ini)
"""
import tkinter as tk
from tkinter import ttk, messagebox
from database import get_conn
from helpers import fmt_rp, today_str, BULAN_NAMA
from pages.base_page import BasePage, C_BG, C_WHITE, C_BLUE, C_DARK, C_GREEN

# BULAN_NAMA dari helpers adalah dict {1:'Januari',...}
BULAN_LIST = [BULAN_NAMA[i] for i in range(1, 13)]

JENIS_LBL = {
    "wajib":    "Sim. Wajib",
    "sukarela": "Sim. Sukarela",
    "pokok":    "Sim. Pokok",
    "sihara":   "Sihara",
    "simkus":   "Simkus",
}


class KasPage(BasePage):
    def __init__(self, parent):
        super().__init__(parent)
        self._build()
        self.refresh()

    # ─────────────────────────────────────────────────────────────────
    def _build(self):
        from datetime import date
        now = date.today()

        # ── Filter bar ───────────────────────────────────────────────
        fbar = tk.Frame(self, bg=C_BG)
        fbar.pack(fill="x", padx=12, pady=(8, 4))

        tk.Label(fbar, text="Bulan:", font=("Arial", 9), bg=C_BG).pack(side="left")
        self._v_bulan = tk.StringVar(value=BULAN_NAMA[now.month])
        cb_bln = ttk.Combobox(fbar, textvariable=self._v_bulan,
                               values=BULAN_LIST, width=12, state="readonly")
        cb_bln.pack(side="left", padx=4)
        cb_bln.bind("<<ComboboxSelected>>", lambda _: self.refresh())

        tk.Label(fbar, text="Tahun:", font=("Arial", 9), bg=C_BG).pack(side="left", padx=(8, 0))
        self._v_tahun = tk.StringVar(value=str(now.year))
        ttk.Combobox(fbar, textvariable=self._v_tahun,
                     values=[str(y) for y in range(2020, 2031)],
                     width=6, state="readonly").pack(side="left", padx=4)
        self._v_tahun.trace_add("write", lambda *_: self.refresh())

        self.btn(fbar, "🔍 Tampilkan",   self.refresh, C_BLUE).pack(side="left", padx=8)
        self.btn(fbar, "📥 Export Excel", self._export, C_GREEN).pack(side="left", padx=4)

        tk.Label(fbar,
                 text="ℹ  Data otomatis tercatat dari modul Simpanan & Angsuran",
                 font=("Arial", 8, "italic"), bg=C_BG, fg="#64748B").pack(side="left", padx=12)

        # ── Notebook ─────────────────────────────────────────────────
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=4)

        self._tab_kas  = tk.Frame(nb, bg=C_BG)
        self._tab_simp = tk.Frame(nb, bg=C_BG)
        self._tab_ags  = tk.Frame(nb, bg=C_BG)

        nb.add(self._tab_kas,  text="📋 Buku Kas Masuk")
        nb.add(self._tab_simp, text="💰 Rekap Simpanan")
        nb.add(self._tab_ags,  text="🔄 Rekap Angsuran")

        self._build_tab_kas()
        self._build_tab_simpanan()
        self._build_tab_angsuran()

    # ── TAB 1 ────────────────────────────────────────────────────────
    def _build_tab_kas(self):
        tc = tk.Frame(self._tab_kas, bg=C_WHITE,
                      highlightbackground="#E2E8F0", highlightthickness=1)
        tc.pack(fill="both", expand=True, padx=4, pady=4)

        tk.Label(tc, text="Buku Kas Masuk", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).pack(anchor="w", padx=12, pady=(8, 2))
        self._lbl_tot_kas = tk.Label(tc, text="", font=("Arial", 8, "bold"),
                                      bg=C_WHITE, fg=C_BLUE)
        self._lbl_tot_kas.pack(anchor="w", padx=12, pady=(0, 4))

        cols = ("ID", "Tgl", "Uraian", "Ket/Ke", "KAS",
                "Piutang", "Jasa", "Sim.Wajib", "Sihara",
                "Sim.SK", "Simkus", "Sim.Pokok", "Lain")
        self._tv_kas = self.make_tree(tc, cols, height=20)
        widths  = [0, 90, 200, 80, 110, 110, 90, 90, 80, 90, 70, 90, 80]
        anchors = ["c","c","w","c","e","e","e","e","e","e","e","e","e"]
        for col, w, a in zip(cols, widths, anchors):
            self._tv_kas.heading(col, text=col)
            self._tv_kas.column(col, width=w, anchor=a)
        self._tv_kas.column("ID", width=0, stretch=False)

    # ── TAB 2 ────────────────────────────────────────────────────────
    def _build_tab_simpanan(self):
        tc = tk.Frame(self._tab_simp, bg=C_WHITE,
                      highlightbackground="#E2E8F0", highlightthickness=1)
        tc.pack(fill="both", expand=True, padx=4, pady=4)

        tk.Label(tc, text="Rekap Simpanan Bulan Ini", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).pack(anchor="w", padx=12, pady=(8, 2))
        self._lbl_tot_simp = tk.Label(tc, text="", font=("Arial", 8, "bold"),
                                       bg=C_WHITE, fg=C_BLUE)
        self._lbl_tot_simp.pack(anchor="w", padx=12, pady=(0, 4))

        cols = ("ID", "No.KOP", "Nama", "Jenis", "Jumlah", "Tanggal", "Keterangan")
        self._tv_simp = self.make_tree(tc, cols, height=22)
        for col, w, a in zip(cols,
                              [0, 80, 210, 110, 120, 90, 260],
                              ["c","c","w","c","e","c","w"]):
            self._tv_simp.heading(col, text=col)
            self._tv_simp.column(col, width=w, anchor=a)
        self._tv_simp.column("ID", width=0, stretch=False)

    # ── TAB 3 ────────────────────────────────────────────────────────
    def _build_tab_angsuran(self):
        tc = tk.Frame(self._tab_ags, bg=C_WHITE,
                      highlightbackground="#E2E8F0", highlightthickness=1)
        tc.pack(fill="both", expand=True, padx=4, pady=4)

        tk.Label(tc, text="Rekap Angsuran Bulan Ini", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).pack(anchor="w", padx=12, pady=(8, 2))
        self._lbl_tot_ags = tk.Label(tc, text="", font=("Arial", 8, "bold"),
                                      bg=C_WHITE, fg=C_BLUE)
        self._lbl_tot_ags.pack(anchor="w", padx=12, pady=(0, 4))

        cols = ("ID", "No.KOP", "Nama", "Ke", "Jangka",
                "Total Bayar", "Pokok", "Jasa 1%", "Tanggal", "Info Pinjaman")
        self._tv_ags = self.make_tree(tc, cols, height=22)
        for col, w, a in zip(cols,
                              [0, 80, 200, 40, 60, 110, 110, 90, 90, 230],
                              ["c","c","w","c","c","e","e","e","c","w"]):
            self._tv_ags.heading(col, text=col)
            self._tv_ags.column(col, width=w, anchor=a)
        self._tv_ags.column("ID", width=0, stretch=False)

    # ── Refresh semua tab ────────────────────────────────────────────
    def refresh(self, *_):
        bulan = BULAN_LIST.index(self._v_bulan.get()) + 1
        tahun = int(self._v_tahun.get())
        self._load_kas(bulan, tahun)
        self._load_simpanan(bulan, tahun)
        self._load_angsuran(bulan, tahun)

    def _load_kas(self, bulan, tahun):
        conn = get_conn()
        try:
            rows = conn.execute("""
                SELECT id, tgl, uraian, keterangan,
                       kas, piutang, jasa, sim_wajib, sihara,
                       sim_sukarela, simkus, sim_pokok, lain_lain
                FROM kas
                WHERE bulan=? AND tahun=? AND jenis='masuk'
                ORDER BY tgl, id
            """, (bulan, tahun)).fetchall()
            tot = conn.execute("""
                SELECT SUM(kas), SUM(piutang), SUM(jasa),
                       SUM(sim_wajib), SUM(sim_sukarela), SUM(lain_lain)
                FROM kas WHERE bulan=? AND tahun=? AND jenis='masuk'
            """, (bulan, tahun)).fetchone()
        finally:
            conn.close()

        self.tv_clear(self._tv_kas)
        for r in rows:
            f = lambda n: fmt_rp(n) if n else ""
            self.tv_insert(self._tv_kas, (
                r[0], r[1], r[2], r[3] or "",
                f(r[4]), f(r[5]), f(r[6]),
                f(r[7]), f(r[8]), f(r[9]),
                f(r[10]), f(r[11]), f(r[12]),
            ))
        self._lbl_tot_kas.config(
            text=(f"Total KAS: {fmt_rp(tot[0] or 0)}  |  "
                  f"Piutang: {fmt_rp(tot[1] or 0)}  |  "
                  f"Jasa: {fmt_rp(tot[2] or 0)}  |  "
                  f"Sim.Wajib: {fmt_rp(tot[3] or 0)}  |  "
                  f"Sukarela: {fmt_rp(tot[4] or 0)}  |  "
                  f"Lain: {fmt_rp(tot[5] or 0)}")
        )

    def _load_simpanan(self, bulan, tahun):
        conn = get_conn()
        try:
            rows = conn.execute("""
                SELECT s.id, a.no_anggota, a.nama, s.jenis, s.jumlah, s.tgl, s.keterangan
                FROM simpanan s
                JOIN anggota a ON s.anggota_id = a.id
                WHERE s.bulan=? AND s.tahun=?
                ORDER BY CAST(SUBSTR(a.no_anggota,5) AS INTEGER), s.jenis
            """, (bulan, tahun)).fetchall()
            tot = conn.execute(
                "SELECT SUM(jumlah) FROM simpanan WHERE bulan=? AND tahun=?",
                (bulan, tahun)
            ).fetchone()[0]
        finally:
            conn.close()

        self.tv_clear(self._tv_simp)
        for r in rows:
            self.tv_insert(self._tv_simp, (
                r[0], r[1], r[2],
                JENIS_LBL.get(r[3], r[3]),
                fmt_rp(r[4]), r[5], r[6] or ""
            ))
        self._lbl_tot_simp.config(
            text=f"Total Simpanan: {fmt_rp(tot or 0)}  |  {len(rows)} transaksi"
        )

    def _load_angsuran(self, bulan, tahun):
        conn = get_conn()
        try:
            rows = conn.execute("""
                SELECT ag.id, a.no_anggota, a.nama,
                       ag.ke, p.jangka, ag.jumlah,
                       ag.tgl, p.keterangan, p.jumlah as saldo, p.bunga
                FROM angsuran ag
                JOIN pinjaman p ON ag.pinjaman_id = p.id
                JOIN anggota  a ON p.anggota_id   = a.id
                WHERE ag.bulan=? AND ag.tahun=?
                ORDER BY CAST(SUBSTR(a.no_anggota,5) AS INTEGER), ag.ke
            """, (bulan, tahun)).fetchall()
            tot = conn.execute("""
                SELECT SUM(ag.jumlah)
                FROM angsuran ag
                JOIN pinjaman p ON ag.pinjaman_id=p.id
                WHERE ag.bulan=? AND ag.tahun=?
            """, (bulan, tahun)).fetchone()[0]
        finally:
            conn.close()

        self.tv_clear(self._tv_ags)
        for r in rows:
            jasa  = round(float(r[8] or 0) * float(r[9] or 1.0) / 100)
            pokok = max(0, float(r[5] or 0) - jasa)
            self.tv_insert(self._tv_ags, (
                r[0], r[1], r[2],
                r[3], r[4],
                fmt_rp(r[5]),
                fmt_rp(pokok),
                fmt_rp(jasa),
                r[6],
                r[7] or ""
            ))
        self._lbl_tot_ags.config(
            text=(f"Total Bayar: {fmt_rp(tot or 0)}  |  "
                  f"{len(rows)} transaksi angsuran")
        )

    # ── Export Excel ─────────────────────────────────────────────────
    def _export(self):
        from export_excel import export_kas_bulan
        bulan = BULAN_LIST.index(self._v_bulan.get()) + 1
        tahun = int(self._v_tahun.get())
        export_kas_bulan(bulan, tahun)