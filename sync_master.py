import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import Engine
import os
from dotenv import load_dotenv
from typing import Optional, Tuple

# Load environment variables
load_dotenv()

# =========================================================================
# KONFIGURASI DATABASE DARI ENVIRONMENT VARIABLES
# =========================================================================
# Prioritas 1: DATABASE_URL lengkap (recommended untuk Supabase pooler)
# Prioritas 2: Komponen terpisah (DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASS)

DATABASE_URL: Optional[str] = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    # Build dari komponen terpisah
    DB_HOST: str = os.getenv("DB_HOST", "db.ibhkaoacvxonfhfxjyzr.supabase.co")
    DB_PORT: str = os.getenv("DB_PORT", "5432")
    DB_NAME: str = os.getenv("DB_NAME", "postgres")
    DB_USER: str = os.getenv("DB_USER", "postgres")
    DB_PASS: str = os.getenv("DB_PASS", "")

    if DB_PASS:
        DATABASE_URL = (
            f"postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}?sslmode=require"
        )
    else:
        DATABASE_URL = None
        print("⚠️ PERINGATAN: DB_PASS tidak diset. Koneksi database akan gagal.")

CACHE_FILE: str = "master_cache.parquet"  # Format penyimpanan lokal yang sangat cepat & ringan

if DATABASE_URL:
    print(f"✅ Database URL dikonfigurasi: {DB_HOST}:{DB_PORT}/{DB_NAME}")
else:
    print("❌ ERROR: DATABASE_URL atau kredensial DB tidak lengkap di .env")


class MasterDataSync:
    def __init__(self) -> None:
        self.is_online: bool = False
        self.df_master: Optional[pd.DataFrame] = None

    def load_data(self) -> Tuple[pd.DataFrame, bool]:
        """
        Mencoba mengambil data dari Supabase Cloud.
        Jika gagal (offline), otomatis memakai data dari local cache.
        """
        # Cek apakah DATABASE_URL tersedia
        if not DATABASE_URL:
            print("❌ DATABASE_URL tidak dikonfigurasi. Langsung ke mode OFFLINE.")
            self.is_online = False
            if os.path.exists(CACHE_FILE):
                self.df_master = pd.read_parquet(CACHE_FILE)
                print("🟡 Data Master berhasil dimuat dari Cache Lokal!")
            else:
                self.df_master = pd.DataFrame(columns=["barcode", "judul", "harga"])
                print("⚠️ Tidak ada Cache Lokal ditemukan. Data master kosong.")
            return self.df_master, self.is_online

        try:
            # 1. Coba Koneksi ke Database Online
            engine: Engine = create_engine(DATABASE_URL, connect_args={"connect_timeout": 5})
            query: str = "SELECT barcode, judul, harga FROM master_produk"

            # Fetch data dari cloud ke DataFrame Pandas
            self.df_master = pd.read_sql(query, engine)
            self.is_online = True

            # Simpan ke Cache Lokal untuk Cadangan Offline
            self.df_master.to_parquet(CACHE_FILE, index=False)
            print(
                "🟢 [ONLINE] Data Master berhasil ditarik dari Cloud Supabase & Cache diperbarui!"
            )

        except Exception as e:
            # 2. Jika Offline / Gagal Koneksi -> Gunakan Local Cache
            self.is_online = False
            print(
                f"🔴 [OFFLINE MODE] Gagal terhubung ke Cloud ({str(e)}). Membaca dari Cache Lokal..."
            )

            if os.path.exists(CACHE_FILE):
                self.df_master = pd.read_parquet(CACHE_FILE)
                print("🟡 Data Master berhasil dimuat dari Cache Lokal!")
            else:
                self.df_master = pd.DataFrame(columns=["barcode", "judul", "harga"])
                print("⚠️ Tidak ada Cache Lokal ditemukan. Data master kosong.")

        return self.df_master, self.is_online


if __name__ == "__main__":
    # Uji Coba Modul Sync
    sync = MasterDataSync()
    df, status = sync.load_data()
    print("\n--- Hasil Uji Coba Sync ---")
    print("Status Online:", status)
    print("Jumlah Produk Terload:", len(df))
    print(df.head())
