"""helpers.py – Konstanta & fungsi utilitas"""
from datetime import date

JENIS_LIST  = ['pokok','wajib','sukarela','khusus','hariraya']
JENIS_LABEL = {'pokok':'Simpanan Pokok','wajib':'Simpanan Wajib',
               'sukarela':'Simpanan Sukarela','khusus':'Simpanan Khusus',
               'hariraya':'Simpanan Hari Raya'}
BULAN_NAMA  = {1:'Januari',2:'Februari',3:'Maret',4:'April',5:'Mei',6:'Juni',
               7:'Juli',8:'Agustus',9:'September',10:'Oktober',11:'November',12:'Desember'}
BULAN_SHORT = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'Mei',6:'Jun',
               7:'Jul',8:'Agu',9:'Sep',10:'Okt',11:'Nov',12:'Des'}

def fmt_rp(n):
    try:   return f"Rp {int(round(float(n or 0))):,}".replace(",",".")
    except: return "Rp 0"

def hitung_angsuran(jumlah, jangka, bunga):
    return (jumlah / jangka) + (jumlah * bunga / 100)

def today_str():  return str(date.today())
def this_year():  return date.today().year
def this_month(): return date.today().month

def get_periode_aktif():
    """Return periode aktif tunggal dari DB. Return None jika tidak ada."""
    try:
        from database import get_conn
        conn = get_conn()
        row = conn.execute(
            "SELECT id, nama, tahun FROM periode WHERE status='aktif' ORDER BY tahun DESC LIMIT 1"
        ).fetchone()
        conn.close()
        return row  # None jika tidak ada
    except Exception:
        return None

def cek_periode_aktif():
    """Return True jika ada periode aktif."""
    return get_periode_aktif() is not None


# ── Integrasi Otomatis ke Buku Kas ────────────────────────────────────────────
def catat_kas_dari_angsuran(conn, angsuran_id, pin, ke, jumlah, tgl, bulan, tahun):
    """
    Dipanggil setelah INSERT angsuran berhasil.
    Mencatat 1 baris di tabel kas: piutang = pokok, jasa = bunga 1%.
    """
    periode_id = conn.execute(
        "SELECT id FROM periode WHERE tahun=? LIMIT 1", (tahun,)
    ).fetchone()
    periode_id = periode_id[0] if periode_id else None

    saldo   = float(pin.get("jumlah", 0))
    bunga   = float(pin.get("bunga", 1.0))
    jasa    = round(saldo * bunga / 100)
    piutang = max(0, jumlah - jasa)
    kas_tot = jumlah  # total yang masuk ke kas

    nama_anggota = conn.execute(
        "SELECT a.nama FROM anggota a JOIN pinjaman p ON p.anggota_id=a.id WHERE p.id=?",
        (pin["id"],)
    ).fetchone()
    nama = nama_anggota[0] if nama_anggota else "?"

    uraian = f"{nama.upper()}"
    ket    = f"{pin.get('jangka',0)}/'{ke}"

    # Hindari duplikat (angsuran_id unik per baris)
    existing = conn.execute(
        "SELECT id FROM kas WHERE keterangan_ref_angsuran=?", (angsuran_id,)
    ).fetchone()
    if existing:
        return

    conn.execute("""
        INSERT INTO kas (periode_id, bulan, tahun, tgl, uraian, keterangan,
                         kas, piutang, jasa, sim_wajib, sihara, sim_sukarela,
                         simkus, sim_pokok, lain_lain, jenis, keterangan_ref_angsuran)
        VALUES (?,?,?,?,?,?,?,?,?,0,0,0,0,0,0,'masuk',?)
    """, (periode_id, bulan, tahun, tgl, uraian, ket,
          kas_tot, piutang, jasa, angsuran_id))


def catat_kas_dari_simpanan(conn, simpanan_id, anggota_nama, jenis,
                             jumlah, tgl, bulan, tahun):
    """
    Dipanggil setelah INSERT simpanan berhasil.
    Jenis: wajib→sim_wajib, sukarela→sim_sukarela, pokok→sim_pokok,
           sihara→sihara, simkus→simkus
    """
    periode_id = conn.execute(
        "SELECT id FROM periode WHERE tahun=? LIMIT 1", (tahun,)
    ).fetchone()
    periode_id = periode_id[0] if periode_id else None

    col_map = {
        "wajib":    "sim_wajib",
        "sukarela": "sim_sukarela",
        "pokok":    "sim_pokok",
        "sihara":   "sihara",
        "simkus":   "simkus",
    }
    col = col_map.get(jenis, "sim_wajib")

    # Hindari duplikat
    existing = conn.execute(
        "SELECT id FROM kas WHERE keterangan_ref_simpanan=?", (simpanan_id,)
    ).fetchone()
    if existing:
        return

    uraian = anggota_nama.upper()
    ket_kode = {"wajib":"SW","sukarela":"SK","pokok":"SP","sihara":"SHR","simkus":"SIMKUS"}.get(jenis,"")

    conn.execute(f"""
        INSERT INTO kas (periode_id, bulan, tahun, tgl, uraian, keterangan,
                         kas, piutang, jasa, sim_wajib, sihara, sim_sukarela,
                         simkus, sim_pokok, lain_lain, jenis, keterangan_ref_simpanan)
        VALUES (?,?,?,?,?,?,?,0,0,
                {1 if col=='sim_wajib' else 0},{1 if col=='sihara' else 0},
                {1 if col=='sim_sukarela' else 0},{1 if col=='simkus' else 0},
                {1 if col=='sim_pokok' else 0},0,'masuk',?)
    """, (periode_id, bulan, tahun, tgl, uraian, ket_kode,
          jumlah, simpanan_id))

    # Update nilai kolom yang benar
    last_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.execute(f"UPDATE kas SET {col}=?, kas=? WHERE id=?",
                 (jumlah, jumlah, last_id))