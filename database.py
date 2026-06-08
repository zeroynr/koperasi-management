"""
database.py  –  Koneksi & inisialisasi SQLite
Tabel: anggota, periode, simpanan, pinjaman, angsuran
"""
import sqlite3, os

# Saat dijalankan sebagai .exe, koperasi_app.py menetapkan ATRA_DB_PATH
# agar database disimpan di folder user yang bisa ditulis.
# Saat development, fallback ke lokasi file ini (perilaku lama).
DB_FILE = os.environ.get(
    'ATRA_DB_PATH',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "koperasi.db")
)

def get_conn():
    c = sqlite3.connect(DB_FILE)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys = ON")
    return c


def _migrate(cur):
    """Tambah kolom yang belum ada (migrasi skema lama → baru)."""
    migrations = [
        ("simpanan",  "periode_id", "INTEGER"),
        ("simpanan",  "bulan",      "INTEGER"),
        ("simpanan",  "tahun",      "INTEGER"),
        ("pinjaman",  "periode_id", "INTEGER"),
        ("pinjaman",  "tahun",      "INTEGER"),
        ("angsuran",  "bulan",      "INTEGER"),
        ("angsuran",  "tahun",      "INTEGER"),
        ("angsuran",  "keterangan", "TEXT DEFAULT ''"),
        ("angsuran",  "status",     "TEXT NOT NULL DEFAULT 'lunas'"),
    ]
    for tbl, col, coltype in migrations:
        existing = [r[1] for r in cur.execute(f"PRAGMA table_info({tbl})").fetchall()]
        if col not in existing:
            try:
                cur.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {coltype}")
            except Exception:
                pass

    # Isi kolom tahun/bulan dari tgl yang sudah ada
    try:
        cur.execute("""
            UPDATE simpanan SET
                tahun = CAST(strftime('%Y', tgl) AS INTEGER),
                bulan = CAST(strftime('%m', tgl) AS INTEGER)
            WHERE tahun IS NULL AND tgl IS NOT NULL
        """)
    except Exception: pass
    try:
        cur.execute("""
            UPDATE pinjaman SET
                tahun = CAST(strftime('%Y', tgl) AS INTEGER)
            WHERE tahun IS NULL AND tgl IS NOT NULL
        """)
    except Exception: pass
    try:
        cur.execute("""
            UPDATE angsuran SET
                tahun = CAST(strftime('%Y', tgl) AS INTEGER),
                bulan = CAST(strftime('%m', tgl) AS INTEGER)
            WHERE tahun IS NULL AND tgl IS NOT NULL
        """)
    except Exception: pass
    # Pastikan tabel periode ada
    cur.execute("""
        CREATE TABLE IF NOT EXISTS periode (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nama        TEXT NOT NULL,
            tahun       INTEGER NOT NULL,
            bulan_mulai INTEGER NOT NULL,
            bulan_akhir INTEGER NOT NULL,
            keterangan  TEXT DEFAULT '',
            status      TEXT NOT NULL DEFAULT 'aktif'
        )
    """)

def init_db():
    c = get_conn(); cur = c.cursor()
    _migrate(cur)  # migrasi skema lama
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS anggota (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        no_anggota TEXT NOT NULL UNIQUE,
        nama       TEXT NOT NULL,
        alamat     TEXT DEFAULT '',
        no_hp      TEXT DEFAULT '',
        tgl_masuk  TEXT NOT NULL DEFAULT (date('now'))
    );
    CREATE TABLE IF NOT EXISTS periode (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        nama        TEXT NOT NULL,
        tahun       INTEGER NOT NULL,
        bulan_mulai INTEGER NOT NULL,
        bulan_akhir INTEGER NOT NULL,
        keterangan  TEXT DEFAULT '',
        status      TEXT NOT NULL DEFAULT 'aktif' CHECK(status IN ('aktif','tutup'))
    );
    CREATE TABLE IF NOT EXISTS simpanan (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        periode_id INTEGER,
        anggota_id INTEGER NOT NULL,
        jenis      TEXT NOT NULL CHECK(jenis IN ('pokok','wajib','sukarela','khusus','hariraya')),
        jumlah     REAL NOT NULL CHECK(jumlah > 0),
        tgl        TEXT NOT NULL,
        bulan      INTEGER,
        tahun      INTEGER,
        keterangan TEXT DEFAULT '',
        FOREIGN KEY (anggota_id) REFERENCES anggota(id) ON DELETE RESTRICT,
        FOREIGN KEY (periode_id) REFERENCES periode(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS pinjaman (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        periode_id INTEGER,
        anggota_id INTEGER NOT NULL,
        jumlah     REAL NOT NULL CHECK(jumlah > 0),
        jangka     INTEGER NOT NULL CHECK(jangka > 0),
        bunga      REAL NOT NULL DEFAULT 1.5,
        tgl        TEXT NOT NULL,
        tahun      INTEGER,
        keterangan TEXT DEFAULT '',
        status     TEXT NOT NULL DEFAULT 'aktif' CHECK(status IN ('aktif','lunas')),
        FOREIGN KEY (anggota_id) REFERENCES anggota(id) ON DELETE RESTRICT,
        FOREIGN KEY (periode_id) REFERENCES periode(id) ON DELETE SET NULL
    );
    CREATE TABLE IF NOT EXISTS angsuran (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        pinjaman_id INTEGER NOT NULL,
        ke          INTEGER NOT NULL,
        jumlah      REAL NOT NULL CHECK(jumlah > 0),
        tgl         TEXT NOT NULL,
        bulan       INTEGER,
        tahun       INTEGER,
        keterangan  TEXT DEFAULT '',
        status      TEXT NOT NULL DEFAULT 'lunas',
        FOREIGN KEY (pinjaman_id) REFERENCES pinjaman(id) ON DELETE RESTRICT
    );
    CREATE TABLE IF NOT EXISTS kas (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        periode_id   INTEGER,
        bulan        INTEGER NOT NULL,
        tahun        INTEGER NOT NULL,
        tgl          TEXT NOT NULL,
        no_bukti     TEXT,
        no_urut      INTEGER,
        uraian       TEXT NOT NULL,
        keterangan   TEXT,
        kas          REAL DEFAULT 0,
        piutang      REAL DEFAULT 0,
        jasa         REAL DEFAULT 0,
        sim_wajib    REAL DEFAULT 0,
        sihara       REAL DEFAULT 0,
        sim_sukarela REAL DEFAULT 0,
        simkus       REAL DEFAULT 0,
        sim_pokok    REAL DEFAULT 0,
        lain_lain    REAL DEFAULT 0,
        jenis        TEXT DEFAULT 'masuk' CHECK(jenis IN ('masuk','keluar')),
        keterangan_ref_angsuran INTEGER,
        keterangan_ref_simpanan INTEGER,
        FOREIGN KEY (periode_id) REFERENCES periode(id)
    );
    """)
    c.commit(); c.close()