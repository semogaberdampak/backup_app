# Takom Kasir v1.9.12

Aplikasi Point of Sale (POS) / Kasir berbasis Python dengan fitur:
- **UI Modern** - Tkinter dengan tema gelap/terang
- **Barcode Scanner** - Input cepat via scanner atau manual
- **Manajemen Keranjang** - Tambah, edit qty, diskon, hapus item
- **Multi-Payment** - Tunai & Non-Tunai (QRIS, Transfer, EDC)
- **Member & Diskon** - Catatan member per transaksi
- **Nota Thermal 80mm** - **Template Editor fully customizable** (drag & drop, live preview)
- **Laporan Penjualan** - Filter tanggal, ringkasan tunai/non-tunai, export Excel
- **NOMAN Input** - Anotasi numerik per transaksi di laporan
- **Cloud Sync** - Sinkronisasi master produk dari Supabase (PostgreSQL)
- **Offline-First** - Cache lokal Parquet untuk operasi offline
- **Auto-Update** - Launcher cek update aplikasi **& dependensi (requirements.txt)** dari GitHub otomatis, lengkap dengan verifikasi checksum
- **Ingat Username** - Checkbox di dialog login, username terakhir tersimpan otomatis
- **Cross-Platform** - Satu codebase untuk Windows 10+ maupun Linux

---

## 🚀 Quick Start

### 1. Clone & Install
```bash
git clone <repo-url>
cd backup_app

# Buat virtual environment (recommended)
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # Linux/Mac

# Install dependencies
pip install -r requirements.txt
# atau untuk development:
pip install -e ".[dev]"
```

### 2. Konfigurasi Environment
```bash
cp .env.example .env
# Edit .env dengan kredensial Supabase Anda
```

**Contoh `.env`:**
```env
# Supabase (untuk fitur realtime/client)
SUPABASE_URL=https://your-project.supabase.co
SUPABASE_KEY=your-anon-key
SUPABASE_TABLE=master_produk

# Database PostgreSQL (untuk sync_master.py)
# Opsi 1: URL lengkap (recommended untuk Supabase Pooler)
DATABASE_URL=postgresql://user:pass@host:port/dbname?sslmode=require

# Opsi 2: Komponen terpisah
DB_HOST=db.your-project.supabase.co
DB_PORT=5432
DB_NAME=postgres
DB_USER=postgres
DB_PASS=your-password
```

### 3. Jalankan Aplikasi
```bash
# Via launcher (auto-check update dari GitHub)
python launcher.py

# Atau langsung
python app_scan.py
```

---

## 📦 Struktur Project

```
backup_app/
├── app_scan.py          # Aplikasi utama (KasirApp)
├── sync_master.py       # Modul sinkronisasi master data
├── launcher.py          # Auto-updater dari GitHub
├── version.json         # Manifest versi
├── requirements.txt     # Dependencies production
├── pyproject.toml       # Modern packaging config
├── .env.example         # Template konfigurasi
├── .gitignore
├── tests/               # Unit tests
│   ├── conftest.py
│   ├── test_sync_master.py
│   └── test_app_config.py
└── Takomers.zip         # Backup/asset (review needed)
```

---

## 🔧 Development

### Install Dev Dependencies
```bash
pip install -e ".[dev]"
```

### Run Tests
```bash
# Semua test
pytest

# Dengan coverage
pytest --cov=. --cov-report=html

# Test spesifik
pytest tests/test_sync_master.py -v
```

### Code Formatting
```bash
# Format dengan Black
black .

# Lint dengan Flake8
flake8 .
```

---

## 🗄️ Database Schema (Supabase)

Tabel `master_produk`:
```sql
CREATE TABLE master_produk (
    barcode TEXT PRIMARY KEY,
    judul TEXT NOT NULL,
    harga NUMERIC(12,2) NOT NULL DEFAULT 0
);

-- Index untuk pencarian cepat
CREATE INDEX idx_master_produk_judul ON master_produk(judul);
```

---

## 🖨️ Printing

- **Windows**: Menggunakan `pywin32` (RAW print ke printer default)
- **Linux/Mac**: Menggunakan `lpr` (raw mode)
- **Template**: Disimpan di `config_nota.json`, editable via UI

---

## 🔄 Auto-Update System

`launcher.py` memeriksa `version.json` di GitHub:
```json
{
  "version": "v1.8.0",
  "download_url": "https://raw.githubusercontent.com/.../app_scan.py",
  "changelog": "Deskripsi perubahan"
}
```

Jika versi berbeda → download `app_scan.py` baru → simpan `version.local` → jalankan.

---

## 📋 Changelog

### v1.9.11 (2026-10-21)
- 🎨 UI nuansa Windows XP: background abu-abu `#d4d0c8` (lebih gelap, tidak menyilaukan), tema ttk `xpnative` (fallback `clam`)
- 🎨 Tombol ttk: teks hitam jelas, auto-center (`anchor="center"`, `justify="center"`)
- 🎨 Palet login: biru XP (`#003399` / `#316ac5`)

### v1.9.10 (2026-10-21)
- 🐛 Sync master produk via REST kini pagination (ambil semua data, bukan hanya 1000 row)
- 🔢 Version bump v1.9.10

### v1.9.9 (2026-10-21)
- 🔐 Default Supabase URL/anon key di kode (`.env` opsional), `.env` tidak wajib
- 🛡️ Hardening: RLS + RPC `verify_pin` (file `supabase_rls.sql`)

### v1.9.8 (2026-10-21)
- 🌐 Sync master produk via REST Supabase bila pooler PostgreSQL tidak terjangkau (offline-first tetap jalan)

### v1.9.7 (2026-10-20)
- ✨ Launcher men-sinkronkan `requirements.txt` dari repo: checksum diverifikasi (`requirements_sha256` di `version.json`), menawarkan `pip install -r requirements.txt` otomatis saat dependensi berubah — install tanpa perlu zip ulang
- 🧪 Unit test launcher (sinkron requirements) ditambahkan

### v1.9.6 (2026-10-20)
- ✨ Checkbox **"Ingat username"** di samping field username pada dialog login (preferensi tersimpan di `~/.takom_kasir/login_prefs.json`)
- 🐛 Fix maximize window (`state("zoomed")`) hanya jalan di Windows — kini fallback ke `attributes("-zoomed")` (Linux) dan geometry default
- 🐛 Tambah driver `psycopg` v3 — sinkronisasi cloud PostgreSQL kini jalan di Linux (SQLAlchemy 2.1 tidak lagi memakai psycopg2 saja)
- 📦 Dependensi (`requirements.txt`, `pyproject.toml`) & versi aplikasi diselaraskan ke 1.9.6

### v1.9.5
- Kelola role kustom + checklist hak akses (owner), role dinamis dari tabel `app_roles`
- Perbaikan RLS `app_users`/`audit_log`/`arus_kas`/`kas_setoran` yang memblokir simpan user & audit
- Fallback owner/123456 hanya saat DB belum punya user owner

### v1.9.4
- Dialog login modern di tengah layar, palet aplikasi, toggle PIN, pesan error inline, status koneksi

### v1.9.3
- Tab Admin Keuangan, kelola user, arus kas, audit, proteksi tab per peran

### v1.9.2
- RBAC login PIN + audit, proteksi harga/master/setoran, integrasi `kas_setoran`, skema transaksi baru

---

## 📝 License

MIT License - Bebas digunakan & dimodifikasi.

---

## 🤝 Contributing

1. Fork repository
2. Buat branch feature (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push ke branch (`git push origin feature/amazing-feature`)
5. Buat Pull Request

---

## ⚠️ Known Issues / TODO

- [ ] Split `app_scan.py` (3000+ lines) ke modul terpisah (UI, Printing, Data, Sync)
- [ ] Tambah type hints di seluruh codebase
- [ ] Review & hapus `Takomers.zip` (11MB) jika tidak diperlukan
- [ ] Tambah integration test untuk flow transaksi penuh
- [ ] Dockerfile untuk deployment containerized