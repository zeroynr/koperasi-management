# 🏦 ATRA – Sistem Manajemen Koperasi

Aplikasi desktop manajemen koperasi berbasis Python + Tkinter. Dirancang untuk memudahkan pengelolaan data anggota, simpanan, pinjaman, angsuran, periode, dan laporan keuangan koperasi secara offline tanpa memerlukan koneksi internet.

---

## 📋 Daftar Isi

- [Fitur Utama](#fitur-utama)
- [Teknologi](#teknologi)
- [Persyaratan Sistem](#persyaratan-sistem)
- [Instalasi](#instalasi)
- [Cara Menjalankan](#cara-menjalankan)
- [Struktur Proyek](#struktur-proyek)
- [Lisensi](#lisensi)

---

## ✨ Fitur Utama

### 📊 Dashboard
- Ringkasan statistik total anggota, simpanan, pinjaman, dan angsuran
- Tampilan cepat transaksi simpanan dan pinjaman terbaru
- Data diperbarui otomatis setiap kali membuka halaman

### 👥 Manajemen Anggota
- Tambah, edit, dan hapus data anggota
- Data anggota meliputi: No. Anggota, Nama, Alamat, No. HP, Tanggal Masuk
- Fitur pencarian/filter anggota secara real-time
- Validasi nomor anggota agar tidak duplikat

### 🗂️ Manajemen Periode
- Buat dan kelola periode simpanan (contoh: Periode 2024, Periode 2025)
- Setiap periode memiliki: nama, tahun, bulan mulai–akhir, status (aktif/tutup)
- Tutup dan buka kembali periode dengan konfirmasi
- Hanya boleh ada **satu periode aktif** pada satu waktu
- Hapus periode beserta seluruh data simpanan terkait (dengan konfirmasi)
- Filter tampilan: Semua / Aktif / Tutup

### 💰 Simpanan
Mendukung **5 jenis simpanan** dalam satu halaman:
| Jenis | Keterangan |
|---|---|
| Simpanan Pokok | Dibayar sekali saat masuk |
| Simpanan Wajib | Dibayar rutin setiap periode |
| Simpanan Sukarela | Bebas nominal dan waktu |
| Simpanan Hari Raya | Khusus menjelang hari raya |
| Simpanan Khusus | Keperluan tertentu |

- Catat transaksi simpanan per anggota per periode aktif
- Tampilan ringkasan total simpanan per jenis per anggota
- Filter data berdasarkan jenis simpanan, periode, bulan, dan tahun
- Edit dan hapus transaksi simpanan
- Tidak dapat menambah simpanan jika tidak ada periode aktif

### 🏦 Pinjaman
- Catat pengajuan pinjaman anggota (nominal, bunga, tenor, tanggal)
- Preview otomatis: cicilan per bulan dihitung langsung saat input
- Status pinjaman: aktif / lunas
- Filter berdasarkan status dan anggota
- Hapus pinjaman beserta seluruh angsuran terkait

### 📅 Angsuran
- Catat pembayaran angsuran per pinjaman
- Informasi detail pinjaman: sisa hutang, total sudah bayar, cicilan/bulan
- 3 tab tampilan:
  - **Catat Angsuran** — input pembayaran baru
  - **Rekap per Bulan** — ringkasan angsuran bulanan per tahun
  - **Rekap per Tahun** — akumulasi angsuran tahunan semua anggota
- Validasi agar tidak melebihi sisa hutang

### 📤 Rekap & Export Excel
Filter export: periode, jenis simpanan, bulan, tahun

| Export | Sheet yang Dihasilkan |
|---|---|
| **Rekap Lengkap** | Anggota, Simpanan (semua jenis), Pinjaman, Angsuran, Neraca |
| **Simpanan** | Rekap Simpanan, Detail matriks per bulan/tanggal, Matriks Jan–Des, Rekap per Periode |
| **Pinjaman & Angsuran** | Daftar Pinjaman, Riwayat Angsuran, Matriks per Bulan, Rekap per Tahun |
| **Data Anggota** | Daftar lengkap anggota koperasi |
| **Neraca** | Ringkasan keuangan & aset |

**Format sheet Detail Simpanan** mengikuti format matriks:
- Baris header: nama bulan (Januari, Februari, dst)
- Sub-header: tanggal pembayaran aktual
- Hanya tanggal yang ada transaksinya yang dimunculkan

**Sheet Rekap Tahunan Simpanan** menampilkan:
- Kolom per jenis simpanan (Pokok, Wajib, Sukarela, Hari Raya, Khusus)
- Sub-kolom per periode dengan nama asli + rentang tanggal (01 Jan 2024 – 31 Des 2024)
- Total per anggota dan grand total keseluruhan

---

## 🛠️ Teknologi

| Komponen | Library |
|---|---|
| GUI / Tampilan | `tkinter` + `ttk` (built-in Python) |
| Database | `SQLite3` (built-in Python, tanpa server) |
| Export Excel | `openpyxl` |
| Ikon / Gambar | `Pillow` (PIL) |

---

## 💻 Persyaratan Sistem

- **OS:** Windows 10/11, macOS, atau Linux
- **Python:** 3.9 atau lebih baru
- **RAM:** minimal 256 MB
- **Storage:** minimal 50 MB

---

## ⚙️ Instalasi

**1. Clone repositori**
```bash
git clone https://github.com/zeroynr/koperasi-management.git
cd koperasi-management
```

**2. (Opsional) Buat virtual environment**
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

**3. Install dependensi**
```bash
pip install openpyxl Pillow
```

> `tkinter` dan `sqlite3` sudah termasuk dalam instalasi Python standar dan tidak perlu diinstall terpisah.

---

## ▶️ Cara Menjalankan

```bash
python koperasi_app.py
```

Database (`koperasi.db`) akan dibuat otomatis di folder yang sama saat pertama kali dijalankan.

---

## 📁 Struktur Proyek

```
koperasi-management/
├── koperasi_app.py        # Entry point — jendela utama & navigasi
├── database.py            # Inisialisasi & koneksi SQLite
├── export_excel.py        # Logika export ke file Excel (.xlsx)
├── helpers.py             # Fungsi bantu (format angka, nama bulan, dll)
├── pages/
│   ├── dashboard.py       # Halaman dashboard & statistik
│   ├── periode.py         # Manajemen periode
│   ├── anggota.py         # Manajemen anggota
│   ├── simpanan.py        # Transaksi simpanan
│   ├── pinjaman.py        # Manajemen pinjaman
│   ├── angsuran.py        # Pembayaran angsuran
│   ├── rekap.py           # Rekap & export Excel
│   ├── base_page.py       # Base class halaman
│   └── widgets.py         # Komponen UI yang dipakai ulang
├── .gitignore
└── README.md
```

---

## 📌 Catatan

- Data tersimpan lokal di file `koperasi.db` — **tidak diunggah ke server manapun**
- File `koperasi.db` dikecualikan dari repository (lihat `.gitignore`) karena berisi data pribadi anggota
- Simpanan hanya bisa dicatat jika ada **periode aktif**
- Ekspor Excel tersimpan di folder yang dipilih pengguna saat dialog simpan

---

## 📄 Lisensi

MIT License — bebas digunakan dan dimodifikasi untuk keperluan koperasi non-komersial.
