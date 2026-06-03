"""
angsuran.py  –  fix: loop ke-, validasi 1 tgl per pinjaman per hari
FIX:
1. _update_info: next_ke tidak di-clamp ke jangka bila sudah lunas semua
   → bila bayar_kali >= jangka, form dikunci & info tampil "LUNAS"
2. _save INSERT: cek duplikat tanggal per pinjaman_id (1 tgl = 1 angsuran)
3. _save INSERT: cek duplikat ke- lebih ketat (semua status, bukan hanya 'lunas')
"""
import tkinter as tk
from tkinter import ttk, messagebox
from database import get_conn
from helpers import fmt_rp, today_str, cek_periode_aktif, hitung_angsuran, BULAN_NAMA, this_year
from pages.widgets import DatePickerWidget
from pages.base_page import BasePage, C_BG, C_WHITE, C_BLUE, C_DARK, C_GRAY, C_RED

C_GREEN  = "#276749"
C_ORANGE = "#C05621"


def _fmt_angka(val):
    digits = "".join(c for c in str(val) if c.isdigit())
    if not digits:
        return ""
    return f"{int(digits):,}".replace(",", ".")


def _parse_tgl(tgl_str):
    from datetime import datetime as _dt
    try:
        d = _dt.strptime(tgl_str, "%Y-%m-%d")
        return d.month, d.year
    except Exception:
        return None, None


def _hitung_wajib(pin):
    """Total rupiah yang wajib dibayar untuk 1 pinjaman (flat)."""
    jangka    = pin["jangka"]
    jumlah    = pin["jumlah"]
    bunga_pct = pin.get("bunga") or 1.5
    pokok     = round(jumlah / jangka) if jangka > 0 else 0
    bunga     = round(jumlah * bunga_pct / 100)
    ang       = pokok + bunga
    return {
        "pokok_nom":   pokok,
        "bunga_nom":   bunga,
        "ang":         ang,
        "total_wajib": ang * jangka,
    }


class AngsuranPage(BasePage):
    def __init__(self, parent):
        super().__init__(parent)
        self._sel_id       = None
        self._pinjaman_map = {}
        self._build()
        self.refresh()

    # ─────────────────────────────────────────────────────────────
    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self._tab_catat = tk.Frame(nb, bg=C_BG)
        self._tab_rekap = tk.Frame(nb, bg=C_BG)
        nb.add(self._tab_catat, text="📝  Catat Angsuran")
        nb.add(self._tab_rekap, text="📊  Rekap Angsuran")
        self._build_catat(self._tab_catat)
        self._build_rekap(self._tab_rekap)

    # ── Tab Catat ────────────────────────────────────────────────
    def _build_catat(self, parent):
        top = tk.Frame(parent, bg=C_BG)
        top.pack(fill="both", expand=True, padx=4, pady=4)
        top.columnconfigure(1, weight=1)

        form_card = tk.Frame(top, bg=C_WHITE,
                             highlightbackground="#E2E8F0", highlightthickness=1)
        form_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        tk.Label(form_card, text="Form Angsuran", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).grid(row=0, column=0, columnspan=2,
                                              sticky="w", padx=12, pady=(12, 6))

        self._v_pinjaman = tk.StringVar()
        self._v_ke       = tk.StringVar()
        self._v_jumlah   = tk.StringVar()

        self._cb_pin = self.lbl_combo(form_card, "Pinjaman *", 1, self._v_pinjaman, [])
        self._cb_pin.bind("<<ComboboxSelected>>", self._on_pinjaman_change)
        self.lbl_entry(form_card, "Angsuran ke- *", 2, self._v_ke, width=8)

        tk.Label(form_card, text="Jumlah Bayar (Rp)*", font=("Arial", 9),
                 bg=C_WHITE, fg=C_DARK, width=16, anchor="w").grid(
            row=3, column=0, sticky="w", padx=(12, 4), pady=4)
        self._e_jumlah = ttk.Entry(form_card, textvariable=self._v_jumlah, width=20)
        self._e_jumlah.grid(row=3, column=1, sticky="ew", padx=(0, 12), pady=4)
        self._v_jumlah.trace_add("write", self._on_jumlah_change)
        self._jumlah_updating = False

        tk.Label(form_card, text="Tanggal *", font=("Arial", 9),
                 bg=C_WHITE, fg=C_DARK, width=16, anchor="w").grid(
            row=4, column=0, sticky="w", padx=(12, 4), pady=4)
        self._dp_tgl = DatePickerWidget(form_card, label="", bg=C_WHITE)
        self._dp_tgl.grid(row=4, column=1, sticky="w", padx=(0, 12), pady=4)

        # Info pinjaman
        info = tk.Frame(form_card, bg="#EBF8FF", padx=10, pady=8)
        info.grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 0))
        self._v_info_angsuran = tk.StringVar(value="")
        self._v_info_sisa     = tk.StringVar(value="")
        self._v_info_total    = tk.StringVar(value="")
        tk.Label(info, text="Angsuran/bln (flat):", font=("Arial", 8),
                 bg="#EBF8FF", fg=C_GRAY).grid(row=0, column=0, sticky="w")
        tk.Label(info, textvariable=self._v_info_angsuran,
                 font=("Arial", 9, "bold"), bg="#EBF8FF", fg=C_BLUE).grid(
            row=0, column=1, sticky="w", padx=8)
        tk.Label(info, text="Sudah bayar:", font=("Arial", 8),
                 bg="#EBF8FF", fg=C_GRAY).grid(row=1, column=0, sticky="w")
        tk.Label(info, textvariable=self._v_info_sisa,
                 font=("Arial", 9, "bold"), bg="#EBF8FF", fg=C_ORANGE).grid(
            row=1, column=1, sticky="w", padx=8)
        tk.Label(info, text="Sisa lunas:", font=("Arial", 8),
                 bg="#EBF8FF", fg=C_GRAY).grid(row=2, column=0, sticky="w")
        tk.Label(info, textvariable=self._v_info_total,
                 font=("Arial", 9, "bold"), bg="#EBF8FF", fg=C_RED).grid(
            row=2, column=1, sticky="w", padx=8)

        btn_row = tk.Frame(form_card, bg=C_WHITE)
        btn_row.grid(row=6, column=0, columnspan=2, pady=10, padx=12, sticky="ew")
        self.btn(btn_row, "💾 Simpan",         self._save,            C_BLUE).pack(side="left", padx=(0, 4))
        self.btn(btn_row, "✅ Lunas Sekaligus", self._lunas_sekaligus, C_GREEN).pack(side="left", padx=4)
        self.btn(btn_row, "🗑 Hapus",           self._delete,          C_RED).pack(side="left", padx=4)
        self.btn(btn_row, "✖ Bersihkan",       self._clear,           "#718096", fg="white").pack(side="left", padx=4)

        tv_card = tk.Frame(top, bg=C_WHITE,
                           highlightbackground="#E2E8F0", highlightthickness=1)
        tv_card.grid(row=0, column=1, sticky="nsew")

        tk.Label(tv_card, text="Riwayat Angsuran", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).pack(anchor="w", padx=12, pady=(10, 4))

        cols = ("ID", "No. Anggota", "Nama", "Jml Pinjaman", "Ke-", "Jml Bayar", "Tanggal", "Status")
        self._tv = self.make_tree(tv_card, cols, height=20)
        for col, w, a in zip(cols,
                              [40, 100, 160, 120, 50, 120, 100, 80],
                              ["center", "center", "w", "e", "center", "e", "center", "center"]):
            self._tv.heading(col, text=col)
            self._tv.column(col, width=w, anchor=a)
        self._tv.column("ID", width=0, stretch=False)
        self._tv.tag_configure("lunas_semua", background="#F0FFF4", foreground="#276749")
        self._tv.bind("<<TreeviewSelect>>", self._on_select)

    # ── Tab Rekap ────────────────────────────────────────────────
    def _build_rekap(self, parent):
        ctrl = tk.Frame(parent, bg=C_BG)
        ctrl.pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(ctrl, text="Tahun:", font=("Arial", 9), bg=C_BG, fg=C_DARK).pack(side="left")
        self._v_rek_tahun = tk.StringVar(value=str(this_year()))
        years = [str(y) for y in range(2020, this_year() + 3)]
        ttk.Combobox(ctrl, textvariable=self._v_rek_tahun, values=years,
                     width=7, state="readonly").pack(side="left", padx=(4, 20))
        self._v_rek_tahun.trace_add("write", lambda *_: self._load_rekap())

        self._rek_nb = ttk.Notebook(parent)
        self._rek_nb.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self._rek_bulan_frame = tk.Frame(self._rek_nb, bg=C_WHITE)
        self._rek_tahun_frame = tk.Frame(self._rek_nb, bg=C_WHITE)
        self._rek_nb.add(self._rek_bulan_frame, text="📅  Per Bulan (1 Tahun)")
        self._rek_nb.add(self._rek_tahun_frame, text="📆  Per Tahun (Semua)")
        self._rek_nb.bind("<<NotebookTabChanged>>", lambda _: self._load_rekap())
        self._build_rekap_bulan(self._rek_bulan_frame)
        self._build_rekap_tahun(self._rek_tahun_frame)

    def _build_rekap_bulan(self, parent):
        cols = ("No", "No. Anggota", "Nama", "Jml Pinjaman",
                "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
                "Jul", "Agu", "Sep", "Okt", "Nov", "Des", "Total")
        self._tv_rek_bulan = self.make_tree(parent, cols, height=22)
        widths  = [35, 95, 150, 110] + [65]*12 + [90]
        anchors = ["center", "center", "w", "e"] + ["e"]*12 + ["e"]
        for col, w, a in zip(cols, widths, anchors):
            self._tv_rek_bulan.heading(col, text=col)
            self._tv_rek_bulan.column(col, width=w, anchor=a)
        self._tv_rek_bulan.column("No", width=0, stretch=False)
        self._tv_rek_bulan.tag_configure("even", background="#F0F7FF")
        self._tv_rek_bulan.tag_configure("total_row", background="#FED7AA",
                                          font=("Arial", 9, "bold"))
        hsb = ttk.Scrollbar(parent, orient="horizontal",
                             command=self._tv_rek_bulan.xview)
        self._tv_rek_bulan.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        self._tv_rek_bulan.pack(fill="both", expand=True)

    def _build_rekap_tahun(self, parent):
        cols = ("Tahun", "Jml Transaksi", "Total Dibayar", "Rata-rata/Transaksi",
                "Anggota Aktif", "Pinjaman Lunas")
        self._tv_rek_tahun = self.make_tree(parent, cols, height=22)
        widths  = [80, 110, 140, 150, 110, 120]
        anchors = ["center", "center", "e", "e", "center", "center"]
        for col, w, a in zip(cols, widths, anchors):
            self._tv_rek_tahun.heading(col, text=col)
            self._tv_rek_tahun.column(col, width=w, anchor=a)
        self._tv_rek_tahun.tag_configure("even", background="#F0F7FF")
        self._tv_rek_tahun.tag_configure("total_row", background="#FED7AA",
                                          font=("Arial", 9, "bold"))
        self._tv_rek_tahun.pack(fill="both", expand=True, padx=4, pady=4)

    # ── Format angka ─────────────────────────────────────────────
    def _on_jumlah_change(self, *_):
        if self._jumlah_updating:
            return
        self._jumlah_updating = True
        self._v_jumlah.set(_fmt_angka(self._v_jumlah.get()))
        self._jumlah_updating = False

    def _jumlah_raw(self):
        return self._v_jumlah.get().replace(".", "").replace(",", "")

    # ── Helper DB (pakai conn yang sudah ada) ─────────────────────
    def _total_bayar_dari_conn(self, conn, pin_id):
        return conn.execute(
            "SELECT COALESCE(SUM(jumlah), 0) FROM angsuran WHERE pinjaman_id=?",
            (pin_id,)
        ).fetchone()[0]

    def _bayar_kali_dari_conn(self, conn, pin_id):
        return conn.execute(
            "SELECT COUNT(*) FROM angsuran WHERE pinjaman_id=?",
            (pin_id,)
        ).fetchone()[0]

    # ── Data load ────────────────────────────────────────────────
    def refresh(self):
        self._auto_tutup_pinjaman_lunas()
        self._load_pinjaman()
        self._load_tv()
        self._load_rekap()

    def _auto_tutup_pinjaman_lunas(self):
        conn = get_conn()
        try:
            aktif = conn.execute(
                "SELECT id, jumlah, jangka, bunga FROM pinjaman WHERE status='aktif'"
            ).fetchall()
            updated = 0
            for p in aktif:
                w = _hitung_wajib(p)
                if self._total_bayar_dari_conn(conn, p["id"]) >= w["total_wajib"]:
                    conn.execute("UPDATE pinjaman SET status='lunas' WHERE id=?", (p["id"],))
                    updated += 1
            if updated:
                conn.commit()
        except Exception:
            pass
        finally:
            conn.close()

    def _load_pinjaman(self):
        conn = get_conn()
        try:
            rows = conn.execute("""
                SELECT p.id, a.no_anggota, a.nama, p.jumlah, p.jangka, p.bunga
                FROM pinjaman p JOIN anggota a ON p.anggota_id=a.id
                WHERE p.status='aktif' ORDER BY a.nama
            """).fetchall()
        finally:
            conn.close()

        self._pinjaman_map = {}
        for r in rows:
            key = f"{r['no_anggota']} – {r['nama']} ({fmt_rp(r['jumlah'])}, {r['jangka']}x)"
            self._pinjaman_map[key] = dict(r)

        names = list(self._pinjaman_map.keys())
        self._cb_pin["values"] = names
        if names and not self._v_pinjaman.get():
            self._v_pinjaman.set(names[0])
            self._update_info(names[0])

    def _update_info(self, key):
        pin = self._pinjaman_map.get(key)
        if not pin:
            self._v_info_angsuran.set("")
            self._v_info_sisa.set("")
            self._v_info_total.set("")
            return

        w = _hitung_wajib(pin)
        conn = get_conn()
        try:
            bayar_kali  = self._bayar_kali_dari_conn(conn, pin["id"])
            total_bayar = self._total_bayar_dari_conn(conn, pin["id"])
        finally:
            conn.close()

        sisa_kali   = max(0, pin["jangka"] - bayar_kali)
        sisa_rupiah = max(0, w["total_wajib"] - total_bayar)

        self._v_info_angsuran.set(
            f"{fmt_rp(w['ang'])}  "
            f"(pokok {fmt_rp(w['pokok_nom'])} + bunga {fmt_rp(w['bunga_nom'])})"
        )
        self._v_info_sisa.set(
            f"Sudah {bayar_kali}/{pin['jangka']} kali  –  sisa {sisa_kali} kali"
        )
        self._v_info_total.set(
            f"{fmt_rp(sisa_rupiah)}  (dari total {fmt_rp(w['total_wajib'])})"
        )

        # ── FIX BUG LOOP: bila semua kali sudah terbayar, kunci form ──────
        if sisa_kali <= 0 or sisa_rupiah <= 0:
            # Pinjaman ini seharusnya sudah lunas → tampilkan pesan, kosongkan ke-
            self._v_ke.set("–")
            self._jumlah_updating = True
            self._v_jumlah.set("0")
            self._jumlah_updating = False
            # Timpa info sisa supaya jelas
            self._v_info_total.set("✅ LUNAS")
            return

        # Auto-isi ke- berikutnya (tanpa clamp yang salah)
        # next_ke = jumlah record yang ADA + 1, tapi tidak boleh > jangka
        # Hanya masuk sini kalau sisa_kali > 0, jadi next_ke pasti valid
        next_ke = bayar_kali + 1   # tidak perlu clamp — sudah dijaga sisa_kali > 0
        self._v_ke.set(str(next_ke))

        # Auto-isi jumlah: pakai sisa_rupiah tepat bila kali terakhir
        self._jumlah_updating = True
        if sisa_kali == 1:
            self._v_jumlah.set(_fmt_angka(str(int(sisa_rupiah))))
        else:
            self._v_jumlah.set(_fmt_angka(str(int(w["ang"]))))
        self._jumlah_updating = False

    def _on_pinjaman_change(self, _=None):
        self._update_info(self._v_pinjaman.get())

    def _load_tv(self):
        conn = get_conn()
        try:
            rows = conn.execute("""
                SELECT ag.id, a.no_anggota, a.nama, p.jumlah,
                       ag.ke, ag.jumlah, ag.tgl, ag.status
                FROM angsuran ag
                JOIN pinjaman p ON ag.pinjaman_id=p.id
                JOIN anggota a  ON p.anggota_id=a.id
                ORDER BY ag.tgl DESC, ag.id DESC
            """).fetchall()
        finally:
            conn.close()
        self.tv_clear(self._tv)
        for r in rows:
            tag = ("lunas_semua",) if r[7] == "lunas_semua" else ()
            self.tv_insert(self._tv, (
                r[0], r[1], r[2], fmt_rp(r[3]),
                r[4], fmt_rp(r[5]), r[6],
                "LUNAS SEKALIGUS" if r[7] == "lunas_semua" else r[7].upper()
            ), tags=tag)

    def _load_rekap(self):
        try:
            tahun = int(self._v_rek_tahun.get())
        except Exception:
            tahun = this_year()
        self._load_rekap_bulan(tahun)
        self._load_rekap_tahun()

    def _load_rekap_bulan(self, tahun):
        conn = get_conn()
        try:
            pinjaman = conn.execute("""
                SELECT p.id, a.no_anggota, a.nama, p.jumlah
                FROM pinjaman p JOIN anggota a ON p.anggota_id=a.id
                ORDER BY a.nama, p.id
            """).fetchall()
            raw = conn.execute("""
                SELECT pinjaman_id, bulan, SUM(jumlah)
                FROM angsuran WHERE tahun=?
                GROUP BY pinjaman_id, bulan
            """, (tahun,)).fetchall()
        finally:
            conn.close()

        lookup = {(r[0], r[1]): r[2] for r in raw}
        self.tv_clear(self._tv_rek_bulan)
        col_totals  = {b: 0 for b in range(1, 13)}
        grand_total = 0

        for i, p in enumerate(pinjaman):
            row_total  = 0
            bulan_vals = []
            for b in range(1, 13):
                v = lookup.get((p[0], b), 0)
                bulan_vals.append(fmt_rp(v) if v else "-")
                row_total     += v
                col_totals[b] += v
            grand_total += row_total
            tag = ("even",) if i % 2 == 0 else ()
            self.tv_insert(self._tv_rek_bulan, (
                i + 1, p[1], p[2], fmt_rp(p[3]),
                *bulan_vals,
                fmt_rp(row_total) if row_total else "-"
            ), tags=tag)

        total_bulan = [fmt_rp(col_totals[b]) if col_totals[b] else "-" for b in range(1, 13)]
        self.tv_insert(self._tv_rek_bulan, (
            "", "", "TOTAL", "",
            *total_bulan,
            fmt_rp(grand_total)
        ), tags=("total_row",))

    def _load_rekap_tahun(self):
        conn = get_conn()
        try:
            rows = conn.execute("""
                SELECT ag.tahun,
                       COUNT(*)                                           AS jml,
                       SUM(ag.jumlah)                                    AS total,
                       COUNT(DISTINCT p.anggota_id)                      AS anggota,
                       SUM(CASE WHEN p.status='lunas' THEN 1 ELSE 0 END) AS pin_lunas
                FROM angsuran ag
                JOIN pinjaman p ON ag.pinjaman_id=p.id
                WHERE ag.tahun IS NOT NULL
                GROUP BY ag.tahun
                ORDER BY ag.tahun DESC
            """).fetchall()
        finally:
            conn.close()

        self.tv_clear(self._tv_rek_tahun)
        grand_jml = grand_total = 0
        for i, r in enumerate(rows):
            rata = r[2] / r[1] if r[1] else 0
            tag  = ("even",) if i % 2 == 0 else ()
            self.tv_insert(self._tv_rek_tahun, (
                r[0] or "-", r[1], fmt_rp(r[2]), fmt_rp(rata), r[3], r[4]
            ), tags=tag)
            grand_jml   += r[1]
            grand_total += r[2]

        if rows:
            rata_all = grand_total / grand_jml if grand_jml else 0
            self.tv_insert(self._tv_rek_tahun, (
                "TOTAL", grand_jml, fmt_rp(grand_total), fmt_rp(rata_all), "", ""
            ), tags=("total_row",))

    # ── Select treeview ──────────────────────────────────────────
    def _on_select(self, _=None):
        sel = self._tv.selection()
        if not sel:
            return
        vals = self._tv.item(sel[0], "values")
        self._sel_id = vals[0]
        conn = get_conn()
        try:
            ag = conn.execute("SELECT * FROM angsuran WHERE id=?", (self._sel_id,)).fetchone()
            p  = conn.execute("SELECT * FROM pinjaman WHERE id=?", (ag["pinjaman_id"],)).fetchone()
            a  = conn.execute("SELECT no_anggota,nama FROM anggota WHERE id=?",
                              (p["anggota_id"],)).fetchone()
        finally:
            conn.close()
        # Format key HARUS sama persis dengan _pinjaman_map ({jangka}x)
        key = f"{a['no_anggota']} – {a['nama']} ({fmt_rp(p['jumlah'])}, {p['jangka']}x)"
        self._v_pinjaman.set(key)
        self._v_ke.set(str(ag["ke"]))
        self._jumlah_updating = True
        self._v_jumlah.set(_fmt_angka(str(int(ag["jumlah"]))))
        self._jumlah_updating = False
        self._dp_tgl.set(ag["tgl"])

    # ── CRUD ─────────────────────────────────────────────────────
    def _save(self):
        if not cek_periode_aktif():
            messagebox.showwarning("Periode Tutup",
                "Tidak ada periode aktif!\n\nBuka atau tambah periode aktif "
                "di menu Periode sebelum mencatat angsuran.")
            return

        pin_key = self._v_pinjaman.get()
        ke_str  = self._v_ke.get().strip()
        jml_str = self._jumlah_raw()
        tgl     = self._dp_tgl.get()

        if not pin_key or not ke_str or not jml_str or not tgl:
            messagebox.showwarning("Perhatian", "Semua field wajib diisi!")
            return

        # Tolak jika ke- adalah "–" (pinjaman sudah lunas)
        if ke_str == "–":
            messagebox.showinfo("Info", "Pinjaman ini sudah lunas, tidak ada yang perlu dicatat.")
            return

        try:
            ke     = int(ke_str)
            jumlah = float(jml_str)
            if ke <= 0 or jumlah <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Perhatian", "Ke- dan Jumlah harus angka positif!")
            return

        pin = self._pinjaman_map.get(pin_key)
        if not pin:
            messagebox.showerror("Error",
                "Pinjaman tidak ditemukan di daftar aktif.\n"
                "Klik Bersihkan lalu pilih ulang pinjaman dari combobox.")
            return

        if ke > pin["jangka"]:
            messagebox.showwarning("Perhatian",
                f"Angsuran ke-{ke} melebihi jangka ({pin['jangka']} kali)!\n"
                "Gunakan 'Lunas Sekaligus' bila ingin melunasi semua sisa.")
            return

        _bulan, _tahun = _parse_tgl(tgl)
        w = _hitung_wajib(pin)

        conn = get_conn()
        try:
            total_bayar_lama = self._total_bayar_dari_conn(conn, pin["id"])
            sisa_rupiah_lama = max(0, w["total_wajib"] - total_bayar_lama)

            if self._sel_id:
                # ── Mode EDIT ────────────────────────────────────────────
                ag_lama      = conn.execute(
                    "SELECT jumlah, tgl FROM angsuran WHERE id=?", (self._sel_id,)
                ).fetchone()
                jumlah_lama  = ag_lama["jumlah"] if ag_lama else 0
                tgl_lama     = ag_lama["tgl"]    if ag_lama else ""

                # Validasi batas nominal (sisa + jumlah lama = ruang yg tersedia)
                ruang = sisa_rupiah_lama + jumlah_lama
                if jumlah > ruang:
                    messagebox.showwarning("Melebihi Tagihan",
                        f"Jumlah bayar melebihi tagihan!\n"
                        f"Maksimal yang bisa dicatat: {fmt_rp(ruang)}")
                    return

                # ── FIX: cek duplikat tgl bila tgl diubah ────────────────
                if tgl != tgl_lama:
                    tgl_bentrok = conn.execute(
                        "SELECT id FROM angsuran WHERE pinjaman_id=? AND tgl=? AND id!=?",
                        (pin["id"], tgl, self._sel_id)
                    ).fetchone()
                    if tgl_bentrok:
                        messagebox.showwarning("Tanggal Duplikat",
                            f"Sudah ada angsuran pada tanggal {tgl} untuk pinjaman ini.\n"
                            "Satu tanggal hanya boleh satu angsuran per pinjaman.")
                        return

                conn.execute(
                    "UPDATE angsuran SET ke=?, jumlah=?, tgl=? WHERE id=?",
                    (ke, jumlah, tgl, self._sel_id))

                total_baru = self._total_bayar_dari_conn(conn, pin["id"])
                sisa_baru  = max(0, w["total_wajib"] - total_baru)
                if sisa_baru <= 0:
                    conn.execute("UPDATE pinjaman SET status='lunas' WHERE id=?", (pin["id"],))
                    conn.commit()
                    messagebox.showinfo("Berhasil",
                        "Angsuran diperbarui.\n\n🎉 Pinjaman ini sudah LUNAS!")
                else:
                    conn.commit()
                    messagebox.showinfo("Berhasil", "Angsuran diperbarui.")

            else:
                # ── Mode INSERT ──────────────────────────────────────────

                # 1. Cegah pembayaran melebihi sisa tagihan
                if jumlah > sisa_rupiah_lama:
                    messagebox.showwarning("Melebihi Tagihan",
                        f"Jumlah bayar melebihi sisa tagihan!\n"
                        f"Sisa tagihan : {fmt_rp(sisa_rupiah_lama)}\n"
                        "Gunakan 'Lunas Sekaligus' untuk melunasi sekaligus.")
                    return

                # 2. ── FIX BUG DUPLIKAT TANGGAL: 1 tgl = 1 angsuran per pinjaman ──
                tgl_bentrok = conn.execute(
                    "SELECT id FROM angsuran WHERE pinjaman_id=? AND tgl=?",
                    (pin["id"], tgl)
                ).fetchone()
                if tgl_bentrok:
                    messagebox.showwarning("Tanggal Duplikat",
                        f"Sudah ada angsuran pada tanggal {tgl} untuk pinjaman ini.\n"
                        "Satu tanggal hanya boleh satu angsuran per pinjaman.")
                    return

                # 3. Cek duplikat ke- (semua status)
                ke_bentrok = conn.execute(
                    "SELECT id FROM angsuran WHERE pinjaman_id=? AND ke=?",
                    (pin["id"], ke)
                ).fetchone()
                if ke_bentrok:
                    messagebox.showwarning("Duplikat",
                        f"Angsuran ke-{ke} sudah pernah dicatat!\n"
                        "Pilih nomor angsuran yang belum dibayar.")
                    return

                conn.execute(
                    "INSERT INTO angsuran (pinjaman_id, ke, jumlah, tgl, bulan, tahun, status)"
                    " VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (pin["id"], ke, jumlah, tgl, _bulan, _tahun, "lunas"))

                # Cek lunas berdasarkan nominal (pakai conn yang sama → data segar)
                total_baru = self._total_bayar_dari_conn(conn, pin["id"])
                sisa_baru  = max(0, w["total_wajib"] - total_baru)
                bayar_kali = self._bayar_kali_dari_conn(conn, pin["id"])

                if sisa_baru <= 0:
                    conn.execute("UPDATE pinjaman SET status='lunas' WHERE id=?", (pin["id"],))
                    conn.commit()
                    messagebox.showinfo("Berhasil",
                        f"Angsuran ke-{ke} berhasil dicatat.\n\n"
                        f"🎉 Pinjaman ini sudah LUNAS!\n"
                        f"Total dibayar: {fmt_rp(total_baru)}")
                else:
                    sisa_kali = max(0, pin["jangka"] - bayar_kali)
                    conn.commit()
                    messagebox.showinfo("Berhasil",
                        f"Angsuran ke-{ke} berhasil dicatat.\n"
                        f"Sisa: {sisa_kali} kali  ({fmt_rp(sisa_baru)})")

        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()

        self._clear()
        self.refresh()

    def _lunas_sekaligus(self):
        if not cek_periode_aktif():
            messagebox.showwarning("Periode Tutup",
                "Tidak ada periode aktif!\n\nBuka periode aktif terlebih dahulu.")
            return

        pin_key = self._v_pinjaman.get()
        tgl     = self._dp_tgl.get()
        if not pin_key:
            messagebox.showwarning("Perhatian", "Pilih pinjaman terlebih dahulu!")
            return
        if not tgl:
            messagebox.showwarning("Perhatian", "Isi tanggal pembayaran!")
            return

        pin = self._pinjaman_map.get(pin_key)
        if not pin:
            messagebox.showerror("Error", "Pinjaman tidak ditemukan!")
            return

        conn = get_conn()
        try:
            w           = _hitung_wajib(pin)
            total_bayar = self._total_bayar_dari_conn(conn, pin["id"])
            bayar_kali  = self._bayar_kali_dari_conn(conn, pin["id"])
            sisa_rupiah = max(0, w["total_wajib"] - total_bayar)
            sisa_kali   = max(0, pin["jangka"] - bayar_kali)

            if sisa_rupiah <= 0:
                messagebox.showinfo("Info", "Pinjaman ini sudah lunas semua!")
                return

            # Cek duplikat tanggal
            tgl_bentrok = conn.execute(
                "SELECT id FROM angsuran WHERE pinjaman_id=? AND tgl=?",
                (pin["id"], tgl)
            ).fetchone()
            if tgl_bentrok:
                messagebox.showwarning("Tanggal Duplikat",
                    f"Sudah ada angsuran pada tanggal {tgl} untuk pinjaman ini.\n"
                    "Gunakan tanggal lain untuk pembayaran lunas sekaligus.")
                return

            konfirmasi = messagebox.askyesno(
                "Konfirmasi Lunas Sekaligus",
                f"Bayar LUNAS SEKALIGUS pinjaman ini?\n\n"
                f"Sisa angsuran : {sisa_kali} kali\n"
                f"Total dibayar : {fmt_rp(sisa_rupiah)}\n"
                f"Tanggal       : {tgl}\n\n"
                f"Semua sisa angsuran akan dicatat dalam 1 transaksi."
            )
            if not konfirmasi:
                return

            _bulan, _tahun = _parse_tgl(tgl)
            conn.execute(
                "INSERT INTO angsuran (pinjaman_id, ke, jumlah, tgl, bulan, tahun, status)"
                " VALUES (?, ?, ?, ?, ?, ?, ?)",
                (pin["id"], bayar_kali + 1, sisa_rupiah, tgl, _bulan, _tahun, "lunas_semua"))
            conn.execute("UPDATE pinjaman SET status='lunas' WHERE id=?", (pin["id"],))
            conn.commit()
            messagebox.showinfo("Berhasil",
                f"Pinjaman berhasil dilunasi!\n"
                f"Total dibayar: {fmt_rp(sisa_rupiah)}")

        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()

        self._clear()
        self.refresh()

    def _delete(self):
        if not self._sel_id:
            messagebox.showwarning("Perhatian", "Pilih angsuran yang ingin dihapus!")
            return
        if not messagebox.askyesno("Konfirmasi", "Yakin hapus angsuran ini?"):
            return
        conn = get_conn()
        try:
            ag = conn.execute("SELECT * FROM angsuran WHERE id=?", (self._sel_id,)).fetchone()
            conn.execute("DELETE FROM angsuran WHERE id=?", (self._sel_id,))
            conn.execute(
                "UPDATE pinjaman SET status='aktif' WHERE id=? AND status='lunas'",
                (ag["pinjaman_id"],))
            conn.commit()
            messagebox.showinfo("Berhasil", "Angsuran dihapus.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()
        self._clear()
        self.refresh()

    def _clear(self):
        self._sel_id = None
        self._v_ke.set("")
        self._jumlah_updating = True
        self._v_jumlah.set("")
        self._jumlah_updating = False
        self._dp_tgl._set_today()
        self._v_info_angsuran.set("")
        self._v_info_sisa.set("")
        self._v_info_total.set("")
        self._tv.selection_remove(self._tv.selection())