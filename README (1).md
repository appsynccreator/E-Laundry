# 🧺 E-Laundry Management System

Aplikasi manajemen laundry berbasis **Google Apps Script** + **Google Sheets** sebagai database.

## ✨ Fitur Lengkap

| Modul | Fitur |
|---|---|
| **Dashboard** | Statistik harian, grafik 7 hari, layanan populer, transaksi terbaru |
| **Transaksi** | Buat transaksi, hitung otomatis, preview & cetak nota thermal |
| **Daftar Transaksi** | Filter multi-kriteria, update status, export Excel |
| **Pelanggan** | CRUD lengkap, riwayat belanja per pelanggan |
| **Master Layanan** | Harga normal & ekspres, kategori, satuan |
| **Karyawan** | Data karyawan, shift, gaji pokok |
| **Pengeluaran** | Catat pengeluaran per kategori, laporan total |
| **Laporan** | Grafik bulanan, laba bersih, per metode bayar |
| **Manajemen User** | Role-based access + **custom menu per user** |
| **Pengaturan** | Info usaha, metode bayar, backup Excel |
| **Lisensi Hybrid** | Online/offline, cache 24 jam, fallback 7 hari |

## 📁 Struktur File

```
e-laundry/
├── Code.gs        ← Backend Google Apps Script
├── index.html     ← Frontend SPA (Single Page App)
└── README.md      ← Dokumentasi ini
```

## 🚀 Cara Deploy

### 1. Buat Google Spreadsheet baru
Buka [sheets.new](https://sheets.new) → beri nama "E-Laundry DB"

### 2. Buka Apps Script
**Extensions → Apps Script**

### 3. Upload file
- Rename file default jadi `Code.gs` → paste isi `Code.gs`
- Klik **+** → **HTML** → beri nama `index` → paste isi `index.html`

### 4. Deploy sebagai Web App
```
Deploy → New Deployment
Type: Web App
Execute as: Me
Access: Anyone (atau Anyone with Google Account)
```

### 5. Buka URL
Klik URL yang digenerate → aplikasi langsung berjalan (DEMO_MODE=true)

## ⚙️ Konfigurasi

### Demo Mode (default: ON)
Di `Code.gs` baris `var DEMO_MODE = true;`:
- `true` → langsung masuk app tanpa lisensi
- `false` → enforce kode lisensi dari Master Sheet

### Inisialisasi Sheet
Sheet dibuat otomatis saat pertama deploy:
`CONFIG`, `LAYANAN`, `PELANGGAN`, `TRANSAKSI`, `KARYAWAN`, `PENGELUARAN`, `LAPORAN`, `USERS`

## 👥 Role & Akses Menu

| Role | Default Akses |
|---|---|
| **Owner** | Semua menu termasuk Manajemen User |
| **Admin** | Semua kecuali Manajemen User |
| **Kasir** | Dashboard, Transaksi, Pelanggan saja |

> Akses menu dapat dikustomisasi per user di **Pengaturan → Manajemen User**

## 📱 Mobile Support

- ✅ Responsive Android & iOS
- ✅ Bottom navigation bar
- ✅ Swipe gesture sidebar
- ✅ Safe area iPhone X+ (notch & home indicator)
- ✅ Prevent zoom saat input (iOS fix)
- ✅ Add to Home Screen (PWA-like)
- ✅ Theme color Chrome Android

## 🔑 Sistem Lisensi Hybrid

```
Online  → Validasi ke Master Sheet vendor
Cache   → Disimpan 24 jam di ScriptProperties
Offline → Fallback hingga 7 hari tanpa koneksi
Expired → App terkunci, tampil halaman locked
```

## 📊 Sheet Database

| Sheet | Kolom Utama |
|---|---|
| CONFIG | Key, Value |
| LAYANAN | No, Nama, Kategori, Satuan, Harga Normal, Harga Ekspres |
| PELANGGAN | No, Nama, HP, Alamat, Kota, Email, Tipe, Tgl Daftar |
| TRANSAKSI | No. Nota, Tgl Masuk, Pelanggan, Layanan, Qty, Total, Status |
| KARYAWAN | No, Nama, HP, Posisi, Shift, Tgl Masuk, Gaji Pokok |
| PENGELUARAN | No, Tanggal, Keterangan, Kategori, Jumlah |
| USERS | Nama, Email, Role, Status, Menu Akses (JSON) |

## 🛠️ Tech Stack

- **Backend**: Google Apps Script (GAS)
- **Database**: Google Sheets
- **Frontend**: Vanilla HTML/CSS/JS (No framework)
- **Charts**: Chart.js 4.4
- **Font**: Poppins + JetBrains Mono

## 📞 Support

Hubungi vendor untuk aktivasi lisensi dan support teknis.

---
*E-Laundry Management System v1.0 — Built with Google Apps Script*
