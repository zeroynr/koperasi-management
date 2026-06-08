"""
koperasi_app.py  ←  Entry point utama
"""
import sys, os

def resource_path(relative_path: str) -> str:
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)

def db_user_path() -> str:
    if sys.platform == 'win32':
        base = os.environ.get('APPDATA', os.path.expanduser('~'))
    else:
        base = os.path.join(os.path.expanduser('~'), '.local', 'share')
    app_dir = os.path.join(base, 'ATRA')
    os.makedirs(app_dir, exist_ok=True)
    return os.path.join(app_dir, 'koperasi.db')

_db_dest = db_user_path()
if not os.path.exists(_db_dest):
    import shutil
    _db_src = resource_path('koperasi.db')
    if os.path.exists(_db_src):
        shutil.copy2(_db_src, _db_dest)

os.environ['ATRA_DB_PATH'] = _db_dest
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import ttk
from datetime import date
from database import init_db

C_NAV  = "#1E293B"
C_NAV2 = "#0F172A"
C_SEL  = "#334155"
C_BG   = "#F1F5F9"

class KoperasiApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ATRA – Sistem Manajemen Koperasi")
        self.geometry("1180x720")
        self.minsize(960,620)
        self.configure(bg=C_BG)

        # ── Set icon taskbar & title bar ──────────────────────
        # Gunakan iconphoto() dengan PNG untuk hasil terbaik lintas platform,
        # plus iconbitmap() untuk .exe Windows agar taskbar terbaca dengan benar.
        _ico = resource_path("atra.ico")
        _png = resource_path("atra.png")

        # Coba pakai iconphoto (lebih reliable untuk taskbar di Windows 10/11)
        try:
            from PIL import Image, ImageTk
            img_pil = Image.open(_png).convert("RGBA")
            # Sediakan berbagai ukuran agar Windows pilih yang terbaik
            _icons = []
            for sz in (16, 32, 48, 64, 128, 256):
                _icons.append(ImageTk.PhotoImage(img_pil.resize((sz, sz), Image.LANCZOS)))
            # Simpan referensi agar tidak di-GC
            self._icon_images = _icons
            self.iconphoto(True, *_icons)
        except Exception:
            # Fallback ke iconbitmap jika PIL tidak tersedia
            if os.path.exists(_ico):
                try:
                    self.iconbitmap(_ico)
                except Exception:
                    pass

        # Khusus Windows: set AppUserModelID agar taskbar gunakan icon EXE,
        # bukan icon default Python/pythonw.exe
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                    u"ATRA.KoperasiManagement.1.0"
                )
            except Exception:
                pass

        init_db()
        self._setup_style()
        self._build()

    def _setup_style(self):
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("Treeview", font=("Arial",9), rowheight=24,
                    background="white", foreground="#1E293B", fieldbackground="white")
        s.configure("Treeview.Heading", font=("Arial",9,"bold"),
                    background="#2B6CB0", foreground="white", padding=4)
        s.map("Treeview", background=[("selected","#DBEAFE")])
        s.configure("TNotebook.Tab", font=("Arial",9), padding=(10,5))
        s.map("TNotebook.Tab", background=[("selected","white")])
        s.configure("TEntry", padding=4)
        s.configure("TCombobox", padding=3)

    def _build(self):
        # ── Sidebar ──────────────────────────────────────────
        sb = tk.Frame(self, bg=C_NAV, width=210)
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)

        # Logo ATRA di sidebar
        logo = tk.Frame(sb, bg=C_NAV2); logo.pack(fill="x")
        logo_inner = tk.Frame(logo, bg=C_NAV2)
        logo_inner.pack(fill="x", padx=14, pady=(14,10))

        cv = tk.Canvas(logo_inner, width=36, height=36, bg=C_NAV2,
                       highlightthickness=0)
        cv.pack(side="left", padx=(0,10))
        cv.create_rectangle(2, 2, 34, 34, fill="#0C447C", outline="", width=0)
        s = 36/512
        def sx(x): return x*s
        def sy(y): return y*s
        cv.create_line(sx(106),sy(148), sx(256),sy(310), sx(406),sy(148),
                       fill="#1D9E75", width=5, capstyle="round", joinstyle="round")
        cv.create_line(sx(106),sy(364), sx(256),sy(202), sx(406),sy(364),
                       fill="white", width=5, capstyle="round", joinstyle="round")
        cx2, cy2, r1, r2 = sx(256), sy(256), sx(22), sx(11)
        cv.create_oval(cx2-r1,cy2-r1,cx2+r1,cy2+r1, fill="#0C447C", outline="")
        cv.create_oval(cx2-r2,cy2-r2,cx2+r2,cy2+r2, fill="#1D9E75", outline="")

        txt_f = tk.Frame(logo_inner, bg=C_NAV2)
        txt_f.pack(side="left")
        tk.Label(txt_f, text="ATRA", font=("Arial",14,"bold"),
                 bg=C_NAV2, fg="white", anchor="w").pack(anchor="w")
        tk.Label(txt_f, text="Manajemen Koperasi", font=("Arial",7),
                 bg=C_NAV2, fg="#64748B", anchor="w").pack(anchor="w")

        menus = [
            ("Dashboard",       "dashboard",  "▪"),
            ("Periode",         "periode",    "▪"),
            ("Anggota",         "anggota",    "▪"),
            (None,None,None),
            ("Simpanan",        "simpanan",   "▸"),
            (None,None,None),
            ("Pinjaman",        "pinjaman",   "▸"),
            ("Angsuran",        "angsuran",   "▸"),
            (None,None,None),
            ("Rekap & Export",  "rekap",      "▸"),
            (None,None,None),
            ("Buku Kas",        "kas",        "▸"),
        ]
        self._nav_btns = {}
        for label,key,icon in menus:
            if label is None:
                tk.Frame(sb,bg="#334155",height=1).pack(fill="x",padx=12,pady=2)
            else:
                b = tk.Button(sb, text=f"  {icon}  {label}",
                              font=("Arial",10), bg=C_NAV, fg="#94A3B8",
                              bd=0, anchor="w", padx=8, pady=8,
                              cursor="hand2",
                              activebackground=C_SEL, activeforeground="white",
                              command=lambda k=key: self._nav(k))
                b.pack(fill="x"); self._nav_btns[key]=b

        tk.Label(sb, text=date.today().strftime("%d %b %Y"),
                 font=("Arial",8), bg=C_NAV, fg="#475569").pack(
            side="bottom", pady=8)

        # ── Main ─────────────────────────────────────────────
        main = tk.Frame(self, bg=C_BG); main.pack(side="left",fill="both",expand=True)

        tb = tk.Frame(main, bg="white", height=46); tb.pack(fill="x"); tb.pack_propagate(False)
        tk.Frame(tb, bg="#E2E8F0", height=1).pack(side="bottom",fill="x")
        self._title_var = tk.StringVar(value="Dashboard")
        tk.Label(tb, textvariable=self._title_var, font=("Arial",12,"bold"),
                 bg="white", fg="#1E293B").pack(side="left",padx=20,pady=12)
        tk.Label(tb, text="ATRA", font=("Arial",10,"bold"),
                 bg="white", fg="#0C447C").pack(side="right",padx=(0,6),pady=12)
        tk.Label(tb, text="●", font=("Arial",8),
                 bg="white", fg="#1D9E75").pack(side="right",pady=12)

        self._content = tk.Frame(main, bg=C_BG)
        self._content.pack(fill="both",expand=True)

        self._pages = {}
        self._nav("dashboard")

    def _nav(self, key):
        from pages.dashboard import DashboardPage
        from pages.periode   import PeriodePage
        from pages.anggota   import AnggotaPage
        from pages.simpanan  import SimpananPage
        from pages.pinjaman  import PinjamanPage
        from pages.angsuran  import AngsuranPage
        from pages.rekap     import RekapPage
        from pages.kas       import KasPage

        builders = {
            "dashboard": DashboardPage,
            "periode":   PeriodePage,
            "anggota":   AnggotaPage,
            "simpanan":  SimpananPage,
            "pinjaman":  PinjamanPage,
            "angsuran":  AngsuranPage,
            "rekap":     RekapPage,
            "kas":       KasPage,
        }
        titles = {
            "dashboard":"Dashboard","periode":"Manajemen Periode",
            "anggota":"Data Anggota","simpanan":"Kelola Simpanan",
            "pinjaman":"Pinjaman","angsuran":"Angsuran",
            "rekap":"Rekap & Export Excel",
            "kas":"Buku Kas Masuk",
        }
        for k,b in self._nav_btns.items():
            b.config(bg=C_SEL if k==key else C_NAV,
                     fg="white" if k==key else "#94A3B8")
        self._title_var.set(titles.get(key,key))
        for p in self._pages.values(): p.pack_forget()
        if key not in self._pages:
            self._pages[key] = builders[key](self._content)
        elif hasattr(self._pages[key],"refresh"):
            self._pages[key].refresh()
        self._pages[key].pack(fill="both",expand=True)

if __name__ == "__main__":
    app = KoperasiApp()
    app.mainloop()