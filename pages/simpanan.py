"""
simpanan.py
Halaman kelola simpanan anggota (5 jenis).
"""
import tkinter as tk
from tkinter import ttk, messagebox
from database import get_conn
from helpers import fmt_rp, today_str, cek_periode_aktif, JENIS_LIST, JENIS_LABEL
from pages.widgets import DatePickerWidget
from pages.base_page import BasePage, C_BG, C_WHITE, C_BLUE, C_DARK, C_GRAY, C_RED


def _fmt_angka(val):
    """Format angka dengan titik ribuan."""
    digits = "".join(c for c in str(val) if c.isdigit())
    if not digits:
        return ""
    return f"{int(digits):,}".replace(",", ".")


class SimpananPage(BasePage):
    def __init__(self, parent):
        super().__init__(parent)
        self._sel_id = None
        self._anggota_map = {}   # nama → id
        self._build()
        self.refresh()

    def _build(self):
        top = tk.Frame(self, bg=C_BG)
        top.pack(fill="both", expand=True, padx=12, pady=8)
        top.columnconfigure(1, weight=1)

        # ── Form ─────────────────────────────────────────
        form_card = tk.Frame(top, bg=C_WHITE,
                             highlightbackground="#E2E8F0", highlightthickness=1)
        form_card.grid(row=0, column=0, sticky="ns", padx=(0, 8))

        tk.Label(form_card, text="Form Simpanan", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).grid(row=0, column=0, columnspan=2,
                                              sticky="w", padx=12, pady=(12, 6))

        self._v_anggota = tk.StringVar()
        self._v_jenis   = tk.StringVar(value=JENIS_LIST[0])
        self._v_jumlah  = tk.StringVar()
        self._v_ket     = tk.StringVar()

        self._cb_anggota = self.lbl_combo(form_card, "Anggota *", 1, self._v_anggota, [])
        self.lbl_combo(form_card, "Jenis Simpanan *", 2, self._v_jenis,
                       [JENIS_LABEL[j] for j in JENIS_LIST])
        self._v_jenis.set(JENIS_LABEL[JENIS_LIST[0]])
        # Jumlah dengan format titik ribuan
        tk.Label(form_card, text="Jumlah (Rp) *", font=("Arial",9),
                 bg=C_WHITE, fg=C_DARK, width=16, anchor="w").grid(
            row=3, column=0, sticky="w", padx=(12,4), pady=4)
        self._e_jumlah = ttk.Entry(form_card, textvariable=self._v_jumlah, width=20)
        self._e_jumlah.grid(row=3, column=1, sticky="ew", padx=(0,12), pady=4)
        self._jumlah_updating = False
        self._v_jumlah.trace_add("write", self._on_jumlah_change)
        tk.Label(form_card, text="Tanggal *", font=("Arial",9),
                 bg=C_WHITE, fg=C_DARK, width=16, anchor="w").grid(
            row=4, column=0, sticky="w", padx=(12,4), pady=4)
        self._dp_tgl = DatePickerWidget(form_card, label="", bg=C_WHITE)
        self._dp_tgl.grid(row=4, column=1, sticky="w", padx=(0,12), pady=4)
        self.lbl_entry(form_card, "Keterangan",       5, self._v_ket)

        btn_row = tk.Frame(form_card, bg=C_WHITE)
        btn_row.grid(row=6, column=0, columnspan=2, pady=10, padx=12, sticky="ew")
        self.btn(btn_row, "💾 Simpan",    self._save,   C_BLUE).pack(side="left", padx=(0, 4))
        self.btn(btn_row, "🗑 Hapus",     self._delete, C_RED).pack(side="left", padx=4)
        self.btn(btn_row, "✖ Bersihkan", self._clear,  "#718096", fg="white").pack(side="left", padx=4)

        # Ringkasan per jenis
        ring_card = tk.Frame(top, bg=C_WHITE,
                             highlightbackground="#E2E8F0", highlightthickness=1)
        ring_card.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(8, 0))
        tk.Label(ring_card, text="Ringkasan", font=("Arial", 9, "bold"),
                 bg=C_WHITE, fg=C_DARK).pack(anchor="w", padx=12, pady=(8, 4))
        self._ring_frame = tk.Frame(ring_card, bg=C_WHITE)
        self._ring_frame.pack(fill="x", padx=12, pady=(0, 8))

        # ── Treeview ─────────────────────────────────────
        tv_card = tk.Frame(top, bg=C_WHITE,
                           highlightbackground="#E2E8F0", highlightthickness=1)
        tv_card.grid(row=0, column=1, rowspan=2, sticky="nsew")

        tk.Label(tv_card, text="Riwayat Simpanan", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).pack(anchor="w", padx=12, pady=(10, 4))

        # Info periode aktif (otomatis)
        self._lbl_periode_info = tk.Label(tv_card, text="", font=("Arial", 8),
                                           bg=C_WHITE, fg="#64748B")
        self._lbl_periode_info.pack(anchor="w", padx=12, pady=(0, 2))

        # Filter Jenis
        fb = tk.Frame(tv_card, bg=C_WHITE)
        fb.pack(fill="x", padx=12, pady=(0, 6))
        tk.Label(fb, text="Filter Jenis:", font=("Arial", 9), bg=C_WHITE).pack(side="left")
        self._v_filter = tk.StringVar(value="Semua")
        cb_f = ttk.Combobox(fb, textvariable=self._v_filter,
                             values=["Semua"] + [JENIS_LABEL[j] for j in JENIS_LIST],
                             width=18, state="readonly")
        cb_f.pack(side="left", padx=6)
        cb_f.bind("<<ComboboxSelected>>", lambda _: self._load_tv())

        cols = ("ID", "No. Anggota", "Nama", "Jenis", "Jumlah", "Tanggal", "Periode", "Keterangan")
        self._tv = self.make_tree(tv_card, cols, height=18)
        for col, w, a in zip(cols, [40, 100, 160, 130, 110, 100, 120, 160],
                              ["center","center","w","center","e","center","center","w"]):
            self._tv.heading(col, text=col)
            self._tv.column(col, width=w, anchor=a)
        self._tv.column("ID", width=0, stretch=False)
        self._tv.bind("<<TreeviewSelect>>", self._on_select)

    def _on_jumlah_change(self, *_):
        if self._jumlah_updating:
            return
        self._jumlah_updating = True
        formatted = _fmt_angka(self._v_jumlah.get())
        self._v_jumlah.set(formatted)
        self._jumlah_updating = False

    def _jumlah_raw(self):
        return self._v_jumlah.get().replace(".", "").replace(",", "")

    def refresh(self):
        self._load_anggota()
        self._load_tv()
        self._load_ringkasan()

    def _load_anggota(self):
        conn = get_conn()
        try:
            rows = conn.execute("SELECT id,no_anggota,nama FROM anggota ORDER BY CAST(SUBSTR(no_anggota,5) AS INTEGER)").fetchall()
        finally:
            conn.close()
        self._anggota_map = {f"{r['no_anggota']} – {r['nama']}": r['id'] for r in rows}
        names = list(self._anggota_map.keys())
        self._cb_anggota["values"] = names
        if names and not self._v_anggota.get():
            self._v_anggota.set(names[0])

    def _load_tv(self):
        conn = get_conn()
        try:
            # Hanya tampilkan simpanan dari periode AKTIF (periode tutup otomatis disembunyikan)
            q = """
                SELECT s.id, a.no_anggota, a.nama, s.jenis, s.jumlah, s.tgl,
                       p.nama as periode_nama, s.keterangan
                FROM simpanan s
                JOIN anggota a ON s.anggota_id=a.id
                LEFT JOIN periode p ON s.periode_id=p.id
                WHERE (p.status='aktif' OR s.periode_id IS NULL)
            """
            params = []

            # Filter jenis
            f = self._v_filter.get()
            if f != "Semua":
                jenis_key = next((k for k, v in JENIS_LABEL.items() if v == f), None)
                if jenis_key:
                    q += " AND s.jenis=?"
                    params.append(jenis_key)

            q += " ORDER BY s.tgl DESC, s.id DESC"
            rows = conn.execute(q, params).fetchall()

            # Update label info periode aktif (hanya satu)
            aktif_row = conn.execute(
                "SELECT nama, tahun FROM periode WHERE status='aktif' LIMIT 1"
            ).fetchone()
            if aktif_row:
                self._lbl_periode_info.config(
                    text=f"🟢 Periode aktif: {aktif_row['nama']} ({aktif_row['tahun']})",
                    fg="#276749"
                )
            else:
                self._lbl_periode_info.config(
                    text="⚠ Tidak ada periode aktif — tambah simpanan tidak diperbolehkan",
                    fg="#C05621"
                )
        finally:
            conn.close()
        self.tv_clear(self._tv)
        for r in rows:
            self.tv_insert(self._tv, (
                r[0], r[1], r[2], JENIS_LABEL[r[3]],
                fmt_rp(r[4]), r[5], r[6] or "-", r[7] or "-"
            ))

    def _load_ringkasan(self):
        for w in self._ring_frame.winfo_children():
            w.destroy()
        conn = get_conn()
        try:
            for j in JENIS_LIST:
                total = conn.execute(
                    "SELECT COALESCE(SUM(jumlah),0) FROM simpanan WHERE jenis=?", (j,)
                ).fetchone()[0]
                row = tk.Frame(self._ring_frame, bg=C_WHITE)
                row.pack(fill="x", pady=1)
                tk.Label(row, text=JENIS_LABEL[j], font=("Arial", 8),
                         bg=C_WHITE, fg=C_DARK, width=18, anchor="w").pack(side="left")
                tk.Label(row, text=fmt_rp(total), font=("Arial", 8, "bold"),
                         bg=C_WHITE, fg=C_BLUE).pack(side="right")
        finally:
            conn.close()

    def _on_select(self, _=None):
        sel = self._tv.selection()
        if not sel:
            return
        vals = self._tv.item(sel[0], "values")
        self._sel_id = vals[0]
        # Isi balik form
        conn = get_conn()
        try:
            r = conn.execute("SELECT * FROM simpanan WHERE id=?", (self._sel_id,)).fetchone()
            a = conn.execute("SELECT no_anggota,nama FROM anggota WHERE id=?",
                             (r["anggota_id"],)).fetchone()
        finally:
            conn.close()
        key = f"{a['no_anggota']} – {a['nama']}"
        self._v_anggota.set(key)
        self._v_jenis.set(JENIS_LABEL[r["jenis"]])
        self._jumlah_updating = True
        self._v_jumlah.set(_fmt_angka(str(int(r["jumlah"]))))
        self._jumlah_updating = False
        self._dp_tgl.set(r["tgl"])
        self._v_ket.set(r["keterangan"] or "")

    def _jenis_key(self):
        label = self._v_jenis.get()
        return next((k for k, v in JENIS_LABEL.items() if v == label), None)

    def _save(self):
        # Cek periode aktif — wajib ada sebelum simpan
        from helpers import get_periode_aktif
        periode_aktif = get_periode_aktif()
        if periode_aktif is None:
            messagebox.showwarning(
                "Tidak Ada Periode Aktif",
                "Tidak ada periode yang sedang aktif!\n\n"
                "Silakan buka atau tambah periode aktif\n"
                "di menu Periode terlebih dahulu."
            )
            return
        anggota_key = self._v_anggota.get()
        jenis = self._jenis_key()
        jumlah_str = self._jumlah_raw()
        tgl = self._dp_tgl.get()
        ket = self._v_ket.get().strip()

        if not anggota_key or not jenis or not jumlah_str or not tgl:
            messagebox.showwarning("Perhatian", "Semua field wajib diisi!")
            return
        try:
            jumlah = float(jumlah_str)
            if jumlah <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Perhatian", "Jumlah harus angka positif!")
            return

        anggota_id = self._anggota_map.get(anggota_key)
        if not anggota_id:
            messagebox.showerror("Error", "Anggota tidak ditemukan!")
            return

        conn = get_conn()
        try:
            if self._sel_id:
                conn.execute(
                    "UPDATE simpanan SET anggota_id=?,jenis=?,jumlah=?,tgl=?,keterangan=? WHERE id=?",
                    (anggota_id, jenis, jumlah, tgl, ket or None, self._sel_id))
                messagebox.showinfo("Berhasil", "Data simpanan diperbarui.")
            else:
                from datetime import datetime as _dt
                try:
                    _d = _dt.strptime(tgl, "%Y-%m-%d")
                    _bulan, _tahun = _d.month, _d.year
                except Exception:
                    _bulan, _tahun = None, None
                # Gunakan periode aktif yang sudah divalidasi
                _periode_id = periode_aktif[0] if periode_aktif else None
                conn.execute(
                    "INSERT INTO simpanan (anggota_id,jenis,jumlah,tgl,bulan,tahun,periode_id,keterangan) VALUES (?,?,?,?,?,?,?,?)",
                    (anggota_id, jenis, jumlah, tgl, _bulan, _tahun, _periode_id, ket or None))
                messagebox.showinfo("Berhasil", "Simpanan berhasil dicatat.")
            conn.commit()
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()
        self._clear()
        self.refresh()

    def _delete(self):
        if not self._sel_id:
            messagebox.showwarning("Perhatian", "Pilih data simpanan yang ingin dihapus!")
            return
        if not messagebox.askyesno("Konfirmasi", "Yakin hapus data simpanan ini?"):
            return
        conn = get_conn()
        try:
            conn.execute("DELETE FROM simpanan WHERE id=?", (self._sel_id,))
            conn.commit()
            messagebox.showinfo("Berhasil", "Data simpanan dihapus.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()
        self._clear()
        self.refresh()

    def _clear(self):
        self._sel_id = None
        self._jumlah_updating = True
        self._v_jumlah.set("")
        self._jumlah_updating = False
        self._dp_tgl._set_today()
        self._v_ket.set("")
        self._tv.selection_remove(self._tv.selection())