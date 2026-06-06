"""
import_anggota.py
Script untuk memasukkan data anggota ke database koperasi.db
Jalankan: python import_anggota.py
"""
import sqlite3
import os

# Sesuaikan path ke database Anda
# Script ini harus diletakkan di folder yang sama dengan koperasi.db
DB_FILE = os.environ.get(
    'ATRA_DB_PATH',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "koperasi.db")
)

# Daftar nama anggota (127 orang)
NAMA_ANGGOTA = [
    "Dra. H. Sumarmi",
    "Senan",
    "Dra. Pudji Suminiwati",
    "Dra. Yuniani Sri Utami",
    "Hj. Sunarwati, SE",
    "Dra. Hj. Mistiani, MM",
    "Dra. Erneni Achiria Putri",
    "Dra. Hj. Endang Nurina",
    "Susiati, S.Pd, MM",
    "Lilis Suryani, S.Pd",
    "Dra. Rahmah Nuhbatul F., M.Pd",
    "Dra. Diah Primuarini, MM",
    "Mulyadi",
    "A. Hermawan",
    "Suhartono, SE",
    "Dra.Tutut Endri Purbowati, MM",
    "Istilahwati, S.Pd, MM",
    "Dra. Sri Utami Endah L.",
    "Dra. Marti Widasih",
    "Hj. Suparmiati, S.Pd, MM",
    "Dra. Iskun Sri S, M.Pd",
    "Bambang Hendro W",
    "Aminatussuhra, S.Pd.Ekop",
    "Dra. Supeni Lestari, MM",
    "Drs. Tri Setya B, S.Pd.",
    "Indung Amami, S.Kom",
    "Sri Mujianah, SH",
    "Fitriyah, S.Kom",
    "Mardiana As S, S.Pd, M.Pd",
    "Hadi Fatah",
    "Bimo P. S.Kom",
    "Hetty Trisnawati, S.Pd.",
    "Uswatun Bayyina",
    "Retno Kuswati, S.Pd, MM",
    "Sidha N. S.Pd",
    "Dwi Indriani, SE",
    "Yekti Nurnaningrum, S.Sn, MM",
    "Syamsuddin Chalim, S.Pd",
    "Suprayitno, S.Pd",
    "Suwono, S.Pd, MM",
    "Abdullah Musafak, S.Pd.I, M.Pd.I",
    "Miftahu Surur, S.Pd",
    "Aprilia Devina R, SS, M.Pd",
    "Imam Mahfud, S.Pd, M.Si",
    "Andri, S.Kom",
    "Sintia Eka, SE",
    "Ari Wijayanti, S.Pd",
    "Dra. Mardiyah Hayati",
    "Norma Imamah, A.Md",
    "Edin Andreas, ST",
    "Andik S.",
    "Sugiarto",
    "Isa Hamdan, S.Kom",
    "M. Habibi",
    "Nanang Priyo, SH",
    "Marta Rahayu W, S.Pd.K",
    "Asmaul Chusnah, S.Pd",
    "Siti Aisyah, S.Pd.I",
    "Sri Wulandari, S.Pd, M.Si",
    "Kasmah Budi Rahayu, S.Pd",
    "M. Zainul A. S.Pd",
    "Dra. Susiana",
    "Anis Nurhudayani, BT. Scol",
    "Nanang Tri W. S.Pd",
    "Sutiah, S.Pd",
    "M. Hanafi M, S.ST",
    "Sriarianie, S.Kom",
    "Sri Rahayu, S.Pd, M.Pd",
    "Gadis Ary Pratiwi, S.Pd",
    "Winny Excela L. ST",
    "Ainiyah Rochmawati S, S.Pd",
    "Dewi Hidayati, S.Pd",
    "Jaka Subari, S.Pd, M.Pd",
    "Tri Musyafak",
    "Dra. Hj. Mariya Ernawati, M.M",
    "Elfiana Sri Wulandari, S.Pd",
    "M. Salman El Faris, S.Pd",
    "Gunanto Tri Wibowo, S.Kom",
    "M. Nasir",
    "Agus Setiawan",
    "Muhammad Haris, S.Pd",
    "Mochamad Indra, S.Pd",
    "Yuniar Eka Fauzi, S.Pd",
    "Tita Ratnawati, S.Pd",
    "Rahayu Agustina",
    "Serlina Candra Wardina, S.Pd, M.Pd",
    "M. Agung W, S.Pd",
    "Azendio Agung W. K, S.Pd",
    "Damar Setya K. A, ST",
    "Erlyn Argianita, S.Pd",
    "Moh. Novianto Eka P, S.Pd",
    "Sakinah Alfi R, S.Ak",
    "Uswatun Dian I, S.Pd",
    "Emilia Dwi Rahayu N, S.Pd",
    "Imam Subagyo",
    "Dwi Sigit Cahyono, S.Pd",
    "Rahmad Hariadi",
    "Akhmad Zakaria Elmi, S.Pd",
    "Devyana Eka, S.Pd, Gr",
    "Fatmawati U, S.Pd",
    "Vondra",
    "M. Wakid",
    "Ajeng Arianatasari, S.Pd",
    "Firdaus Oscar M, S.Pd",
    "Nurul Latifa H, S.Pd",
    "Lelly Sagitarissa, S.Pd",
    "Bintang Wiyono B.Y, S.Pd",
    "Annissa Atus S, S.Pd",
    "Hanik Rofiqoh, S.Pd",
    "Alfin Shalahuddinta, S.Pd",
    "Lita Niapasa, S.Pd",
    "Dyah Agustin, S.S",
    "Chusnul Chotimah, S.Pd",
    "Asnia Kartikasari D, S.Pd",
    "Shinta Ragil I.P, S.Pd",
    "Romiatin, S.Pd",
    "Risa Rozalia, S.Pd",
    "Rizky Amalia, S.Pd",
    "Ernawati, S.Pd",
    "Nur Halimah, S.Pd",
    "Zeva Agustin, S.Pd",
    "M. Nasir, S.Pd",
    "Winantiu, S.Pd",
    "Tri Wahyuni K, S.Pd",
    "Minhatun N, S.Pd",
    "Rizqi Nailul Hidayat, S.Pd",
    "Mauliddyah R, S.Pd",
]

def import_anggota():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    # Cek no_anggota terakhir yang sudah ada
    last = cur.execute("SELECT no_anggota FROM anggota ORDER BY id DESC LIMIT 1").fetchone()
    if last:
        # Ambil angka dari format KOP-XXX
        try:
            start_num = int(last[0].split("-")[1]) + 1
        except:
            start_num = 1
    else:
        start_num = 1

    berhasil = 0
    gagal = 0
    duplikat = 0

    for i, nama in enumerate(NAMA_ANGGOTA):
        no_anggota = f"KOP-{start_num + i:03d}"
        try:
            cur.execute(
                "INSERT INTO anggota(no_anggota, nama, alamat, no_hp, tgl_masuk) VALUES (?, ?, ?, ?, ?)",
                (no_anggota, nama.strip(), "Jl. Jenggolo No. 2A, Siwalanpanji, Kecamatan Buduran, Kabupaten Sidoarjo", "857-0807-9305", "2025-01-01")
            )
            berhasil += 1
            print(f"  [OK] {no_anggota} - {nama.strip()}")
        except sqlite3.IntegrityError as e:
            if "UNIQUE" in str(e):
                duplikat += 1
                print(f"  [SKIP] {no_anggota} sudah ada (duplikat)")
            else:
                gagal += 1
                print(f"  [ERROR] {nama}: {e}")

    conn.commit()
    conn.close()

    print(f"\n{'='*50}")
    print(f"Selesai! Berhasil: {berhasil} | Duplikat: {duplikat} | Gagal: {gagal}")
    print(f"Total anggota dimasukkan: {berhasil}")

if __name__ == "__main__":
    print(f"Memulai import ke: {DB_FILE}\n")
    import_anggota()
