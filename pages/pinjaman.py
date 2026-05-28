"""
pinjaman.py
Halaman kelola pinjaman anggota.
"""
import tkinter as tk
from tkinter import ttk, messagebox
from database import get_conn
from helpers import fmt_rp, today_str, cek_periode_aktif, hitung_angsuran
from pages.widgets import DatePickerWidget
from pages.base_page import BasePage, C_BG, C_WHITE, C_BLUE, C_DARK, C_GRAY, C_RED


def _fmt_angka(val):
    """Format angka dengan titik ribuan."""
    digits = "".join(c for c in str(val) if c.isdigit())
    if not digits:
        return ""
    return f"{int(digits):,}".replace(",", ".")


class PinjamanPage(BasePage):
    def __init__(self, parent):
        super().__init__(parent)
        self._sel_id = None
        self._anggota_map = {}
        self._jumlah_updating = False   # ← flag anti-loop untuk format titik ribuan
        self._build()
        self.refresh()

    def _build(self):
        top = tk.Frame(self, bg=C_BG)
        top.pack(fill="both", expand=True, padx=12, pady=8)
        top.columnconfigure(1, weight=1)

        # ── Form ─────────────────────────────────────────
        form_card = tk.Frame(top, bg=C_WHITE,
                             highlightbackground="#E2E8F0", highlightthickness=1)
        form_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        tk.Label(form_card, text="Form Pinjaman", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).grid(row=0, column=0, columnspan=2,
                                              sticky="w", padx=12, pady=(12, 6))

        self._v_anggota = tk.StringVar()
        self._v_jumlah  = tk.StringVar()
        self._v_jangka  = tk.StringVar(value="12")
        self._v_bunga   = tk.StringVar(value="1.5")
        self._v_ket     = tk.StringVar()
        self._v_status  = tk.StringVar(value="aktif")

        self._cb_anggota = self.lbl_combo(form_card, "Anggota *", 1, self._v_anggota, [])
        self.lbl_entry(form_card, "Jumlah (Rp) *",  2, self._v_jumlah)

        # ── Jangka angsuran: Spinbox + Slider ────────────────────────────
        tk.Label(form_card, text="Jangka (kali) *", font=("Arial", 9),
                 bg=C_WHITE, fg=C_DARK, width=16, anchor="w").grid(
            row=3, column=0, sticky="w", padx=(12, 4), pady=(4, 0))

        jangka_frame = tk.Frame(form_card, bg=C_WHITE)
        jangka_frame.grid(row=3, column=1, sticky="w", padx=(0, 12), pady=(4, 0))

        # Spinbox: user bisa ketik manual atau klik atas/bawah
        self._spn_jangka = tk.Spinbox(
            jangka_frame, textvariable=self._v_jangka,
            from_=1, to=360, increment=1, width=6,
            font=("Arial", 10, "bold"), fg="#1E3A5F",
            relief="flat", bd=1, highlightthickness=1,
            highlightbackground="#CBD5E0",
            command=self._on_jangka_spin
        )
        self._spn_jangka.pack(side="left")
        tk.Label(jangka_frame, text="kali", font=("Arial", 9),
                 bg=C_WHITE, fg=C_GRAY).pack(side="left", padx=(4, 12))

        # Tombol cepat preset jangka
        preset_frame = tk.Frame(form_card, bg=C_WHITE)
        preset_frame.grid(row=4, column=0, columnspan=2, sticky="w",
                          padx=(12, 4), pady=(2, 4))
        tk.Label(preset_frame, text="Preset:", font=("Arial", 8),
                 bg=C_WHITE, fg=C_GRAY).pack(side="left", padx=(0, 4))
        for preset in [6, 12, 18, 24, 36, 48, 60]:
            tk.Button(
                preset_frame, text=f"{preset}x",
                font=("Arial", 8), bg="#EBF8FF", fg="#2B6CB0",
                relief="flat", bd=0, padx=6, pady=2, cursor="hand2",
                activebackground="#BEE3F8",
                command=lambda v=preset: self._set_jangka(v)
            ).pack(side="left", padx=2)

        # Slider jangka (1–120)
        slider_frame = tk.Frame(form_card, bg=C_WHITE)
        slider_frame.grid(row=5, column=0, columnspan=2, sticky="ew",
                          padx=12, pady=(0, 4))
        tk.Label(slider_frame, text="1x", font=("Arial", 7),
                 bg=C_WHITE, fg=C_GRAY).pack(side="left")
        self._slider_jangka = tk.Scale(
            slider_frame, from_=1, to=120,
            orient="horizontal", variable=self._v_jangka,
            showvalue=False, length=180,
            bg=C_WHITE, fg="#2B6CB0",
            troughcolor="#EBF8FF", highlightthickness=0,
            command=lambda _: self._update_preview()
        )
        self._slider_jangka.pack(side="left", padx=4)
        tk.Label(slider_frame, text="120x", font=("Arial", 7),
                 bg=C_WHITE, fg=C_GRAY).pack(side="left")

        self.lbl_entry(form_card, "Bunga/bln (%) *", 6, self._v_bunga, width=10)
        tk.Label(form_card, text="Tanggal *", font=("Arial", 9),
                 bg=C_WHITE, fg=C_DARK, width=16, anchor="w").grid(
            row=7, column=0, sticky="w", padx=(12, 4), pady=4)
        self._dp_tgl = DatePickerWidget(form_card, label="", bg=C_WHITE)
        self._dp_tgl.grid(row=7, column=1, sticky="w", padx=(0, 12), pady=4)
        self.lbl_entry(form_card, "Keterangan",      8, self._v_ket)
        self.lbl_combo(form_card, "Status",          9, self._v_status, ["aktif", "lunas"])

        # Preview angsuran detail
        prev = tk.Frame(form_card, bg="#EBF8FF", padx=10, pady=8)
        prev.grid(row=10, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 0))

        tk.Label(prev, text="📊 Estimasi Angsuran", font=("Arial", 8, "bold"),
                 bg="#EBF8FF", fg=C_DARK).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,4))

        def _lbl_pair(parent, row, label, var, fg=C_BLUE):
            tk.Label(parent, text=label, font=("Arial", 8), bg="#EBF8FF",
                     fg=C_GRAY, anchor="w").grid(row=row, column=0, sticky="w")
            tk.Label(parent, textvariable=var, font=("Arial", 9, "bold"),
                     bg="#EBF8FF", fg=fg, anchor="w").grid(row=row, column=1, sticky="w", padx=(6,0))

        self._v_preview      = tk.StringVar(value="-")   # angsuran/bln
        self._v_prev_pokok   = tk.StringVar(value="-")   # pokok/bln
        self._v_prev_bunga   = tk.StringVar(value="-")   # bunga/bln
        self._v_prev_total_b = tk.StringVar(value="-")   # total bunga selama jangka
        self._v_prev_total   = tk.StringVar(value="-")   # total seluruh pembayaran

        _lbl_pair(prev, 1, "Angsuran/bulan :", self._v_preview,      C_BLUE)
        _lbl_pair(prev, 2, "  – Pokok       :", self._v_prev_pokok,  "#374151")
        _lbl_pair(prev, 3, "  – Bunga/bln   :", self._v_prev_bunga,  C_GRAY)
        _lbl_pair(prev, 4, "Total Bunga     :", self._v_prev_total_b, "#C05621")
        _lbl_pair(prev, 5, "Total Bayar     :", self._v_prev_total,   "#276749")

        # ── Trace: format titik ribuan + update preview ──
        # Jumlah: format titik ribuan dulu, lalu update preview
        self._v_jumlah.trace_add("write", self._on_jumlah_change)
        # Jangka & bunga: langsung update preview
        self._v_jangka.trace_add("write", lambda *_: self._update_preview())
        self._v_bunga.trace_add("write",  lambda *_: self._update_preview())

        btn_row = tk.Frame(form_card, bg=C_WHITE)
        btn_row.grid(row=11, column=0, columnspan=2, pady=10, padx=12, sticky="ew")
        self.btn(btn_row, "💾 Simpan",    self._save,   C_BLUE).pack(side="left", padx=(0, 4))
        self.btn(btn_row, "🗑 Hapus",     self._delete, C_RED).pack(side="left", padx=4)
        self.btn(btn_row, "✖ Bersihkan", self._clear,  "#718096", fg="white").pack(side="left", padx=4)

        # ── Treeview ─────────────────────────────────────
        tv_card = tk.Frame(top, bg=C_WHITE,
                           highlightbackground="#E2E8F0", highlightthickness=1)
        tv_card.grid(row=0, column=1, sticky="nsew")

        tk.Label(tv_card, text="Daftar Pinjaman", font=("Arial", 10, "bold"),
                 bg=C_WHITE, fg=C_DARK).pack(anchor="w", padx=12, pady=(10, 4))

        # Filter status
        fb = tk.Frame(tv_card, bg=C_WHITE)
        fb.pack(fill="x", padx=12, pady=(0, 6))
        tk.Label(fb, text="Filter:", font=("Arial", 9), bg=C_WHITE).pack(side="left")
        self._v_filter = tk.StringVar(value="Semua")
        cb_f = ttk.Combobox(fb, textvariable=self._v_filter,
                             values=["Semua", "aktif", "lunas"], width=10, state="readonly")
        cb_f.pack(side="left", padx=6)
        cb_f.bind("<<ComboboxSelected>>", lambda _: self._load_tv())

        cols = ("ID", "No. Anggota", "Nama", "Jumlah", "Jangka", "Bunga", "Angsuran/bln", "Tgl", "Status")
        self._tv = self.make_tree(tv_card, cols, height=18)
        for col, w, a in zip(cols,
                              [40, 100, 160, 120, 70, 70, 110, 100, 80],
                              ["center", "center", "w", "e", "center", "center", "e", "center", "center"]):
            self._tv.heading(col, text=col)
            self._tv.column(col, width=w, anchor=a)
        self._tv.column("ID", width=0, stretch=False)
        self._tv.bind("<<TreeviewSelect>>", self._on_select)

    # ── Format titik ribuan ───────────────────────────────────────────────────

    def _on_jumlah_change(self, *_):
        """Callback: format _v_jumlah dengan titik ribuan setiap kali berubah."""
        if self._jumlah_updating:
            return
        self._jumlah_updating = True
        try:
            formatted = _fmt_angka(self._v_jumlah.get())
            self._v_jumlah.set(formatted)
        finally:
            self._jumlah_updating = False
        self._update_preview()

    def _jumlah_raw(self):
        """Kembalikan nilai jumlah bersih (tanpa titik) sebagai string angka."""
        return self._v_jumlah.get().replace(".", "").replace(",", "")

    # ── Refresh & load ────────────────────────────────────────────────────────

    def refresh(self):
        self._load_anggota()
        self._load_tv()

    def _load_anggota(self):
        conn = get_conn()
        try:
            rows = conn.execute(
                "SELECT id, no_anggota, nama FROM anggota ORDER BY nama"
            ).fetchall()
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
            q = """
                SELECT p.id, a.no_anggota, a.nama, p.jumlah, p.jangka,
                       p.bunga, p.tgl, p.status
                FROM pinjaman p JOIN anggota a ON p.anggota_id = a.id
            """
            f = self._v_filter.get()
            params = ()
            if f != "Semua":
                q += " WHERE p.status = ?"
                params = (f,)
            q += " ORDER BY p.tgl DESC"
            rows = conn.execute(q, params).fetchall()
        finally:
            conn.close()
        self.tv_clear(self._tv)
        for r in rows:
            ang = hitung_angsuran(r[3], r[4], r[5])
            self.tv_insert(self._tv, (
                r[0], r[1], r[2], fmt_rp(r[3]),
                f"{r[4]} bln", f"{r[5]}%", fmt_rp(ang),
                r[6], r[7].upper()
            ))

    # ── Preview angsuran ──────────────────────────────────────────────────────

    def _set_jangka(self, val):
        """Set jangka dari tombol preset, update slider & preview."""
        self._v_jangka.set(str(val))
        try:
            self._slider_jangka.set(min(val, 120))
        except Exception:
            pass
        self._update_preview()

    def _on_jangka_spin(self):
        """Dipanggil saat spinbox naik/turun."""
        try:
            val = int(self._v_jangka.get())
            self._slider_jangka.set(min(max(val, 1), 120))
        except Exception:
            pass
        self._update_preview()

    def _update_preview(self):
        try:
            j  = float(self._jumlah_raw())
            jk = int(self._v_jangka.get())
            b  = float(self._v_bunga.get())
            if j <= 0 or jk <= 0:
                raise ValueError
            pokok_bln  = round(j / jk)
            bunga_bln  = round(j * b / 100)
            ang        = pokok_bln + bunga_bln
            total_bunga= bunga_bln * jk
            total_bayar= j + total_bunga

            # Estimasi durasi dari tanggal pinjam (pure stdlib)
            try:
                from datetime import datetime
                tgl_str   = self._dp_tgl.get()
                tgl_mulai = datetime.strptime(tgl_str, "%Y-%m-%d").date()
                m_selesai = tgl_mulai.month - 1 + jk
                thn_sel   = tgl_mulai.year + m_selesai // 12
                bln_sel   = m_selesai % 12 + 1
                from datetime import date as _date
                tgl_selesai = _date(thn_sel, bln_sel, 1)
                durasi_txt  = (f"  ({tgl_mulai.strftime('%b %Y')} → "
                               f"{tgl_selesai.strftime('%b %Y')})")
            except Exception:
                durasi_txt = ""

            self._v_preview.set(f"{fmt_rp(ang)}{durasi_txt}")
            self._v_prev_pokok.set(fmt_rp(pokok_bln))
            self._v_prev_bunga.set(f"{fmt_rp(bunga_bln)}  ({b}%/bln)")
            self._v_prev_total_b.set(fmt_rp(total_bunga))
            self._v_prev_total.set(fmt_rp(total_bayar))
        except Exception:
            for v in [self._v_preview, self._v_prev_pokok,
                      self._v_prev_bunga, self._v_prev_total_b, self._v_prev_total]:
                v.set("-")

    # ── Pilih baris treeview ──────────────────────────────────────────────────

    def _on_select(self, _=None):
        sel = self._tv.selection()
        if not sel:
            return
        vals = self._tv.item(sel[0], "values")
        self._sel_id = vals[0]
        conn = get_conn()
        try:
            r = conn.execute(
                "SELECT * FROM pinjaman WHERE id = ?", (self._sel_id,)
            ).fetchone()
            a = conn.execute(
                "SELECT no_anggota, nama FROM anggota WHERE id = ?",
                (r["anggota_id"],)
            ).fetchone()
        finally:
            conn.close()

        key = f"{a['no_anggota']} – {a['nama']}"
        self._v_anggota.set(key)

        # Set jumlah dengan format titik ribuan (gunakan flag agar tidak double-format)
        self._jumlah_updating = True
        try:
            self._v_jumlah.set(_fmt_angka(str(int(r["jumlah"]))))
        finally:
            self._jumlah_updating = False

        self._v_jangka.set(str(r["jangka"]))
        self._v_bunga.set(str(r["bunga"]))
        self._dp_tgl.set(r["tgl"])
        self._v_ket.set(r["keterangan"] or "")
        self._v_status.set(r["status"])
        self._update_preview()

    # ── Simpan ────────────────────────────────────────────────────────────────

    def _save(self):
        # Cek periode aktif
        if not cek_periode_aktif():
            messagebox.showwarning(
                "Periode Tutup",
                "Tidak ada periode aktif!\n\n"
                "Buka atau tambah periode aktif di menu Periode\n"
                "sebelum mencatat pinjaman."
            )
            return

        anggota_key = self._v_anggota.get()
        jumlah_str  = self._jumlah_raw()          # ← pakai helper, bukan .get() langsung
        tgl         = self._dp_tgl.get()
        ket         = self._v_ket.get().strip()
        status      = self._v_status.get()

        if not anggota_key or not jumlah_str or not tgl:
            messagebox.showwarning("Perhatian", "Anggota, Jumlah, dan Tanggal wajib diisi!")
            return
        try:
            jumlah = float(jumlah_str)
            jangka = int(self._v_jangka.get())
            bunga  = float(self._v_bunga.get())
            if jumlah <= 0 or jangka <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Perhatian", "Jumlah/Jangka/Bunga tidak valid!")
            return

        anggota_id = self._anggota_map.get(anggota_key)
        if not anggota_id:
            messagebox.showerror("Error", "Anggota tidak ditemukan!")
            return

        conn = get_conn()
        try:
            if self._sel_id:
                conn.execute(
                    "UPDATE pinjaman "
                    "SET anggota_id=?, jumlah=?, jangka=?, bunga=?, tgl=?, keterangan=?, status=? "
                    "WHERE id=?",
                    (anggota_id, jumlah, jangka, bunga, tgl, ket or None, status, self._sel_id)
                )
                messagebox.showinfo("Berhasil", "Data pinjaman diperbarui.")
            else:
                from datetime import datetime as _dt
                try:
                    _tahun = _dt.strptime(tgl, "%Y-%m-%d").year
                except Exception:
                    _tahun = None
                conn.execute(
                    "INSERT INTO pinjaman "
                    "(anggota_id, jumlah, jangka, bunga, tgl, tahun, keterangan, status) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (anggota_id, jumlah, jangka, bunga, tgl, _tahun, ket or None, status)
                )
                messagebox.showinfo("Berhasil", "Pinjaman berhasil dicatat.")
            conn.commit()
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()

        self._clear()
        self.refresh()

    # ── Hapus ─────────────────────────────────────────────────────────────────

    def _delete(self):
        if not self._sel_id:
            messagebox.showwarning("Perhatian", "Pilih pinjaman yang ingin dihapus!")
            return
        if not messagebox.askyesno(
            "Konfirmasi", "Yakin hapus pinjaman ini? (Angsuran terkait juga terhapus)"
        ):
            return
        conn = get_conn()
        try:
            conn.execute("DELETE FROM angsuran WHERE pinjaman_id=?", (self._sel_id,))
            conn.execute("DELETE FROM pinjaman WHERE id=?", (self._sel_id,))
            conn.commit()
            messagebox.showinfo("Berhasil", "Pinjaman dihapus.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()
        self._clear()
        self.refresh()

    # ── Bersihkan form ────────────────────────────────────────────────────────

    def _clear(self):
        self._sel_id = None
        # Set jumlah tanpa memicu format ulang yang tidak perlu
        self._jumlah_updating = True
        try:
            self._v_jumlah.set("")
        finally:
            self._jumlah_updating = False
        self._v_jangka.set("12")
        self._v_bunga.set("1.5")
        self._dp_tgl._set_today()
        self._v_ket.set("")
        self._v_status.set("aktif")
        self._v_preview.set("-")
        self._tv.selection_remove(self._tv.selection())