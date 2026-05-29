"""
angsuran.py
Halaman catat & kelola angsuran pinjaman.
- Bayar per angsuran biasa
- Bayar LUNAS SEKALIGUS (semua sisa)
- Rekap tab: Per Bulan (dalam 1 tahun) & Per Tahun (multi tahun)
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


class AngsuranPage(BasePage):
    def __init__(self, parent):
        super().__init__(parent)
        self._sel_id      = None
        self._pinjaman_map = {}
        self._build()
        self.refresh()

    # ─────────────────────────────────────────────────────────────
    def _build(self):
        # Notebook: Tab Catat | Tab Rekap
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

        # Form card
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

        # Tombol aksi
        btn_row = tk.Frame(form_card, bg=C_WHITE)
        btn_row.grid(row=6, column=0, columnspan=2, pady=10, padx=12, sticky="ew")
        self.btn(btn_row, "💾 Simpan",        self._save,         C_BLUE).pack(side="left", padx=(0, 4))
        self.btn(btn_row, "✅ Lunas Sekaligus", self._lunas_sekaligus, C_GREEN).pack(side="left", padx=4)
        self.btn(btn_row, "🗑 Hapus",          self._delete,       C_RED).pack(side="left", padx=4)
        self.btn(btn_row, "✖ Bersihkan",      self._clear,        "#718096", fg="white").pack(side="left", padx=4)

        # Treeview riwayat
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

        # Filter tahun
        tk.Label(ctrl, text="Tahun:", font=("Arial", 9), bg=C_BG, fg=C_DARK).pack(side="left")
        self._v_rek_tahun = tk.StringVar(value=str(this_year()))
        years = [str(y) for y in range(2020, this_year() + 3)]
        ttk.Combobox(ctrl, textvariable=self._v_rek_tahun, values=years,
                     width=7, state="readonly").pack(side="left", padx=(4, 20))
        self._v_rek_tahun.trace_add("write", lambda *_: self._load_rekap())

        # Sub-tab: Per Bulan | Per Tahun
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
        """Matriks: baris=anggota/pinjaman, kolom=Jan-Des untuk 1 tahun."""
        cols = ("No", "No. Anggota", "Nama", "Jml Pinjaman",
                "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
                "Jul", "Agu", "Sep", "Okt", "Nov", "Des", "Total")
        self._tv_rek_bulan = self.make_tree(parent, cols, height=22)
        widths = [35, 95, 150, 110] + [65]*12 + [90]
        anchors = ["center", "center", "w", "e"] + ["e"]*12 + ["e"]
        for col, w, a in zip(cols, widths, anchors):
            self._tv_rek_bulan.heading(col, text=col)
            self._tv_rek_bulan.column(col, width=w, anchor=a)
        self._tv_rek_bulan.column("No", width=0, stretch=False)
        self._tv_rek_bulan.tag_configure("even", background="#F0F7FF")
        self._tv_rek_bulan.tag_configure("total_row", background="#FED7AA",
                                          font=("Arial", 9, "bold"))
        # Scrollbar horizontal
        hsb = ttk.Scrollbar(parent, orient="horizontal",
                             command=self._tv_rek_bulan.xview)
        self._tv_rek_bulan.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        self._tv_rek_bulan.pack(fill="both", expand=True)

    def _build_rekap_tahun(self, parent):
        """Ringkasan per tahun: total bayar, jumlah transaksi, rata-rata."""
        cols = ("Tahun", "Jml Transaksi", "Total Dibayar", "Rata-rata/Transaksi",
                "Anggota Aktif", "Pinjaman Lunas")
        self._tv_rek_tahun = self.make_tree(parent, cols, height=22)
        widths = [80, 110, 140, 150, 110, 120]
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

    # ── Data load ────────────────────────────────────────────────
    def refresh(self):
        self._load_pinjaman()
        self._load_tv()
        self._load_rekap()

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
        pokok_nom = round(pin["jumlah"] / pin["jangka"]) if pin["jangka"] > 0 else 0
        bunga_nom = round(pin["jumlah"] * (pin["bunga"] or 1.5) / 100)
        ang       = pokok_nom + bunga_nom
        total_wajib = ang * pin["jangka"]   # total yg harus dibayar selama jangka

        conn = get_conn()
        try:
            # Jumlah kali angsuran sudah dibayar (setiap record = 1 kali bayar)
            bayar_kali = conn.execute(
                "SELECT COUNT(*) FROM angsuran WHERE pinjaman_id=? AND status='lunas'",
                (pin["id"],)
            ).fetchone()[0]
            # Total rupiah sudah dibayar
            total_bayar = conn.execute(
                "SELECT COALESCE(SUM(jumlah),0) FROM angsuran WHERE pinjaman_id=?",
                (pin["id"],)
            ).fetchone()[0]
        finally:
            conn.close()

        sisa_kali   = max(0, pin["jangka"] - bayar_kali)
        sisa_rupiah = max(0, total_wajib - total_bayar)

        self._v_info_angsuran.set(
            f"{fmt_rp(ang)}  (pokok {fmt_rp(pokok_nom)} + bunga {fmt_rp(bunga_nom)})"
        )
        self._v_info_sisa.set(
            f"Sudah {bayar_kali}/{pin['jangka']} kali  –  sisa {sisa_kali} kali"
        )
        self._v_info_total.set(
            f"{fmt_rp(sisa_rupiah)}  (dari total {fmt_rp(total_wajib)})"
        )
        # Auto-isi ke- berikutnya dan jumlah
        next_ke = bayar_kali + 1
        if next_ke > pin["jangka"]:
            next_ke = pin["jangka"]
        self._v_ke.set(str(next_ke))
        self._jumlah_updating = True
        self._v_jumlah.set(_fmt_angka(str(int(ang))))
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
        """Matriks per bulan untuk 1 tahun yang dipilih."""
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
            row_total = 0
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

        # Baris total
        total_bulan = [fmt_rp(col_totals[b]) if col_totals[b] else "-" for b in range(1, 13)]
        self.tv_insert(self._tv_rek_bulan, (
            "", "", "TOTAL", "",
            *total_bulan,
            fmt_rp(grand_total)
        ), tags=("total_row",))

    def _load_rekap_tahun(self):
        """Ringkasan per tahun dari semua data angsuran."""
        conn = get_conn()
        try:
            rows = conn.execute("""
                SELECT ag.tahun,
                       COUNT(*)             as jml,
                       SUM(ag.jumlah)       as total,
                       COUNT(DISTINCT p.anggota_id) as anggota,
                       SUM(CASE WHEN p.status='lunas' THEN 1 ELSE 0 END) as pin_lunas
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
                r[0] or "-",
                r[1],
                fmt_rp(r[2]),
                fmt_rp(rata),
                r[3],
                r[4]
            ), tags=tag)
            grand_jml   += r[1]
            grand_total += r[2]

        if rows:
            rata_all = grand_total / grand_jml if grand_jml else 0
            self.tv_insert(self._tv_rek_tahun, (
                "TOTAL", grand_jml, fmt_rp(grand_total),
                fmt_rp(rata_all), "", ""
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
        key = f"{a['no_anggota']} – {a['nama']} ({fmt_rp(p['jumlah'])}, {p['jangka']} bln)"
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
            messagebox.showerror("Error", "Pinjaman tidak ditemukan!")
            return

        # Validasi ke- tidak boleh melebihi jangka
        if ke > pin["jangka"]:
            messagebox.showwarning("Perhatian",
                f"Angsuran ke-{ke} melebihi jangka pinjaman ({pin['jangka']} kali)!\n"
                "Gunakan tombol 'Lunas Sekaligus' bila ingin melunasi semua sisa.")
            return

        conn = get_conn()
        try:
            if self._sel_id:
                # Update: boleh ubah jumlah & tgl, ke- tidak berubah
                conn.execute(
                    "UPDATE angsuran SET ke=?,jumlah=?,tgl=? WHERE id=?",
                    (ke, jumlah, tgl, self._sel_id))
                messagebox.showinfo("Berhasil", "Angsuran diperbarui.")
            else:
                # Cek apakah ke- ini sudah pernah dicatat
                ada = conn.execute(
                    "SELECT id FROM angsuran WHERE pinjaman_id=? AND ke=? AND status='lunas'",
                    (pin["id"], ke)
                ).fetchone()
                if ada:
                    messagebox.showwarning("Duplikat",
                        f"Angsuran ke-{ke} sudah pernah dicatat!\n"
                        "Pilih nomor angsuran yang belum dibayar.")
                    return
                from datetime import datetime as _dt
                try:
                    _d = _dt.strptime(tgl, "%Y-%m-%d")
                    _bulan, _tahun = _d.month, _d.year
                except Exception:
                    _bulan, _tahun = None, None
                conn.execute(
                    "INSERT INTO angsuran (pinjaman_id,ke,jumlah,tgl,bulan,tahun,status)"
                    " VALUES (?,?,?,?,?,?,?)",
                    (pin["id"], ke, jumlah, tgl, _bulan, _tahun, "lunas"))
                # Cek apakah sudah bayar semua kali
                bayar_kali = conn.execute(
                    "SELECT COUNT(*) FROM angsuran WHERE pinjaman_id=? AND status='lunas'",
                    (pin["id"],)
                ).fetchone()[0]
                if bayar_kali >= pin["jangka"]:
                    conn.execute("UPDATE pinjaman SET status='lunas' WHERE id=?", (pin["id"],))
                    messagebox.showinfo("Berhasil",
                        f"Angsuran ke-{ke} berhasil dicatat.\n\n"
                        f"🎉 Pinjaman ini sudah LUNAS! ({pin['jangka']} dari {pin['jangka']} kali)")
                else:
                    sisa = pin["jangka"] - bayar_kali
                    messagebox.showinfo("Berhasil",
                        f"Angsuran ke-{ke} berhasil dicatat.\n"
                        f"Sisa: {sisa} kali lagi.")
            conn.commit()
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()
        self._clear()
        self.refresh()

    def _lunas_sekaligus(self):
        """Bayar semua sisa angsuran sekaligus dalam 1 transaksi."""
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
            bayar = conn.execute(
                "SELECT COUNT(*) FROM angsuran WHERE pinjaman_id=?", (pin["id"],)
            ).fetchone()[0]
            total_bayar = conn.execute(
                "SELECT COALESCE(SUM(jumlah),0) FROM angsuran WHERE pinjaman_id=?", (pin["id"],)
            ).fetchone()[0]
        finally:
            conn.close()

        pokok_nom  = round(pin["jumlah"] / pin["jangka"]) if pin["jangka"] > 0 else 0
        bunga_nom  = round(pin["jumlah"] * (pin["bunga"] or 1.5) / 100)
        ang        = pokok_nom + bunga_nom
        total_jml  = ang * pin["jangka"]
        sisa_rupiah = max(0, total_jml - total_bayar)
        sisa_kali   = max(0, pin["jangka"] - bayar)

        if sisa_kali <= 0:
            messagebox.showinfo("Info", "Pinjaman ini sudah lunas semua!")
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

        from datetime import datetime as _dt
        try:
            _d = _dt.strptime(tgl, "%Y-%m-%d")
            _bulan, _tahun = _d.month, _d.year
        except Exception:
            _bulan, _tahun = None, None

        conn = get_conn()
        try:
            conn.execute(
                "INSERT INTO angsuran (pinjaman_id,ke,jumlah,tgl,bulan,tahun,status)"
                " VALUES (?,?,?,?,?,?,?)",
                (pin["id"], bayar + 1, sisa_rupiah, tgl, _bulan, _tahun, "lunas_semua"))
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
            # Jika pinjaman sudah lunas karena angsuran ini, kembalikan ke aktif
            ag = conn.execute("SELECT * FROM angsuran WHERE id=?", (self._sel_id,)).fetchone()
            conn.execute("DELETE FROM angsuran WHERE id=?", (self._sel_id,))
            conn.execute("UPDATE pinjaman SET status='aktif' WHERE id=? AND status='lunas'",
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