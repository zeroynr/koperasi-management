"""
anggota.py
Halaman kelola data anggota koperasi.
"""
import tkinter as tk
from tkinter import ttk, messagebox
from database import get_conn
from helpers import today_str
from pages.widgets import DatePickerWidget
from pages.base_page import BasePage, C_BG, C_WHITE, C_BLUE, C_DARK, C_GRAY, C_RED


class AnggotaPage(BasePage):
    def __init__(self, parent):
        super().__init__(parent)
        self._sel_id = None
        self._build()
        self.refresh()

    # ── helper: generate no anggota berikutnya ────────────────────────
    def _next_no_anggota(self):
        conn = get_conn()
        try:
            row = conn.execute(
                """SELECT no_anggota FROM anggota
                   ORDER BY CAST(SUBSTR(no_anggota, 5) AS INTEGER) DESC
                   LIMIT 1"""
            ).fetchone()
        finally:
            conn.close()
        if row:
            try:
                num = int(row[0].split("-")[1]) + 1
            except Exception:
                num = 1
        else:
            num = 1
        return f"KOP-{num:03d}"

    def _build(self):
        # ── Top: form + tabel ─────────────────────────────
        top = tk.Frame(self, bg=C_BG)
        top.pack(fill="both", expand=True, padx=12, pady=8)
        top.columnconfigure(1, weight=1)

        # Form panel
        form_card = tk.Frame(top, bg=C_WHITE,
                             highlightbackground="#E2E8F0", highlightthickness=1)
        form_card.grid(row=0, column=0, sticky="ns", padx=(0, 8))

        tk.Label(form_card, text="Form Anggota", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).grid(row=0, column=0, columnspan=2,
                                              sticky="w", padx=12, pady=(12, 6))

        self._v_no   = tk.StringVar()
        self._v_nama = tk.StringVar()
        self._v_almt = tk.StringVar()
        self._v_hp   = tk.StringVar()

        self.lbl_entry(form_card, "No. Anggota *", 1, self._v_no)
        self.lbl_entry(form_card, "Nama Lengkap *", 2, self._v_nama)
        self.lbl_entry(form_card, "Alamat",         3, self._v_almt)
        self.lbl_entry(form_card, "No. HP",         4, self._v_hp)

        # Tanggal masuk dengan kalender
        tk.Label(form_card, text="Tgl Masuk *", font=("Arial", 9),
                 bg=C_WHITE, fg=C_DARK, width=16, anchor="w").grid(
            row=5, column=0, sticky="w", padx=(12, 4), pady=4)
        self._dp_tgl = DatePickerWidget(form_card, label="", bg=C_WHITE)
        self._dp_tgl.grid(row=5, column=1, sticky="w", padx=(0, 12), pady=4)

        btn_row = tk.Frame(form_card, bg=C_WHITE)
        btn_row.grid(row=6, column=0, columnspan=2, pady=10, padx=12, sticky="ew")

        self.btn(btn_row, "💾 Simpan",   self._save,   C_BLUE).pack(side="left", padx=(0, 4))
        self.btn(btn_row, "🗑 Hapus",    self._delete, C_RED).pack(side="left", padx=4)
        self.btn(btn_row, "✖ Bersihkan", self._clear,  "#718096", fg="white").pack(side="left", padx=4)

        # Treeview panel
        tv_card = tk.Frame(top, bg=C_WHITE,
                           highlightbackground="#E2E8F0", highlightthickness=1)
        tv_card.grid(row=0, column=1, sticky="nsew")

        tk.Label(tv_card, text="Daftar Anggota", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).pack(anchor="w", padx=12, pady=(10, 4))

        # Search bar
        sb = tk.Frame(tv_card, bg=C_WHITE)
        sb.pack(fill="x", padx=12, pady=(0, 6))
        tk.Label(sb, text="Cari:", font=("Arial", 9), bg=C_WHITE).pack(side="left")
        self._v_cari = tk.StringVar()
        self._v_cari.trace_add("write", lambda *_: self._filter())
        ttk.Entry(sb, textvariable=self._v_cari, width=28).pack(side="left", padx=6)

        cols = ("ID", "No. Anggota", "Nama", "Alamat", "No. HP", "Tgl Masuk")
        self._tv = self.make_tree(tv_card, cols, height=16)
        widths  = [40, 100, 180, 200, 120, 100]
        anchors = ["center", "center", "w", "w", "center", "center"]
        for col, w, a in zip(cols, widths, anchors):
            self._tv.heading(col, text=col)
            self._tv.column(col, width=w, anchor=a)
        self._tv.column("ID", width=0, stretch=False)  # Sembunyikan ID

        # Klik heading kolom untuk sort
        for col in cols[1:]:  # skip ID
            self._tv.heading(col, text=col,
                             command=lambda c=col: self._sort_by(c))

        self._tv.bind("<<TreeviewSelect>>", self._on_select)

    # ── Data ──────────────────────────────────────────────────────────
    def refresh(self):
        self._all_rows = []
        conn = get_conn()
        try:
            rows = conn.execute(
                """SELECT id, no_anggota, nama, alamat, no_hp, tgl_masuk
                   FROM anggota
                   ORDER BY CAST(SUBSTR(no_anggota, 5) AS INTEGER)"""
            ).fetchall()
            self._all_rows = [tuple(r) for r in rows]
        finally:
            conn.close()
        self._render(self._all_rows)

    def _render(self, rows):
        self.tv_clear(self._tv)
        for r in rows:
            self.tv_insert(self._tv, (r[0], r[1], r[2],
                                      r[3] or "-", r[4] or "-", r[5]))

    def _filter(self):
        q = self._v_cari.get().lower()
        if not q:
            self._render(self._all_rows)
            return
        filtered = [r for r in self._all_rows
                    if q in r[1].lower() or q in r[2].lower()]
        self._render(filtered)

    # ── Sort kolom ────────────────────────────────────────────────────
    _sort_asc = {}

    def _sort_by(self, col):
        col_map = {
            "No. Anggota": 1, "Nama": 2,
            "Alamat": 3, "No. HP": 4, "Tgl Masuk": 5
        }
        idx = col_map.get(col, 1)
        asc = not self._sort_asc.get(col, True)
        self._sort_asc[col] = asc

        def key_fn(r):
            val = r[idx] or ""
            # Urutan numerik untuk No. Anggota
            if col == "No. Anggota":
                try:
                    return int(val.split("-")[1])
                except Exception:
                    return 0
            return val.lower()

        sorted_rows = sorted(self._all_rows, key=key_fn, reverse=not asc)
        self._render(sorted_rows)

    # ── Pilih baris ───────────────────────────────────────────────────
    def _on_select(self, _=None):
        sel = self._tv.selection()
        if not sel:
            return
        vals = self._tv.item(sel[0], "values")
        self._sel_id = vals[0]
        self._v_no.set(vals[1])
        self._v_nama.set(vals[2])
        self._v_almt.set(vals[3] if vals[3] != "-" else "")
        self._v_hp.set(vals[4] if vals[4] != "-" else "")
        self._dp_tgl.set(vals[5])

    # ── CRUD ──────────────────────────────────────────────────────────
    def _save(self):
        no   = self._v_no.get().strip()
        nama = self._v_nama.get().strip()
        almt = self._v_almt.get().strip()
        hp   = self._v_hp.get().strip()
        tgl  = self._dp_tgl.get()
        if not no or not nama:
            messagebox.showwarning("Perhatian", "No. Anggota dan Nama wajib diisi!")
            return
        if not tgl:
            messagebox.showwarning("Perhatian", "Tanggal masuk tidak valid!")
            return
        conn = get_conn()
        try:
            if self._sel_id:
                conn.execute(
                    "UPDATE anggota SET no_anggota=?,nama=?,alamat=?,no_hp=?,tgl_masuk=? WHERE id=?",
                    (no, nama, almt or None, hp or None, tgl, self._sel_id))
                messagebox.showinfo("Berhasil", "Data anggota berhasil diperbarui.")
            else:
                conn.execute(
                    "INSERT INTO anggota (no_anggota,nama,alamat,no_hp,tgl_masuk) VALUES (?,?,?,?,?)",
                    (no, nama, almt or None, hp or None, tgl))
                messagebox.showinfo("Berhasil", "Anggota baru berhasil ditambahkan.")
            conn.commit()
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()
        self._clear()
        self.refresh()

    def _delete(self):
        if not self._sel_id:
            messagebox.showwarning("Perhatian", "Pilih anggota yang ingin dihapus!")
            return
        if not messagebox.askyesno("Konfirmasi", "Yakin ingin menghapus anggota ini?"):
            return
        conn = get_conn()
        try:
            conn.execute("DELETE FROM anggota WHERE id=?", (self._sel_id,))
            conn.commit()
            messagebox.showinfo("Berhasil", "Anggota berhasil dihapus.")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal hapus (mungkin ada transaksi terkait):\n{e}")
        finally:
            conn.close()
        self._clear()
        self.refresh()

    def _clear(self):
        self._sel_id = None
        # Auto-generate nomor KOP berikutnya untuk tambah baru
        self._v_no.set(self._next_no_anggota())
        self._v_nama.set("")
        self._v_almt.set("")
        self._v_hp.set("")
        self._dp_tgl._set_today()
        self._tv.selection_remove(self._tv.selection())