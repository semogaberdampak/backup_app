import pandas as pd
from sqlalchemy import create_engine
import os

# =========================================================================
# KREDENSIAL DATABASE CLOUD SUPABASE KAMU
# (Samakan dengan yang ada di setup_db_online.py)
# =========================================================================
DB_HOST = "db.ibhkaoacvxonfhfxjyzr.supabase.co"  # Ganti dengan Host Supabase kamu
DB_PORT = "5432"
DB_NAME = "postgres"
DB_USER = "postgres"
DB_PASS = "snderrorcah"                 # Ganti dengan Password kamu

DATABASE_URL="postgresql://postgres.ibhkaoacvxonfhfxjyzr:snderrorcah@aws-0-ap-northeast-2.pooler.supabase.com:6543/postgres?"
CACHE_FILE = "master_cache.parquet" # Format penyimpanan lokal yang sangat cepat & ringan

class MasterDataSync:
    def __init__(self):
        self.is_online = False
        self.df_master = None

    def load_data(self):
        """
        Mencoba mengambil data dari Supabase Cloud.
        Jika gagal (offline), otomatis memakai data dari local cache.
        """
        try:
            # 1. Coba Koneksi ke Database Online
            engine = create_engine(DATABASE_URL, connect_args={'connect_timeout': 5})
            query = "SELECT barcode, judul, harga FROM master_produk"
            
            # Fetch data dari cloud ke DataFrame Pandas
            self.df_master = pd.read_sql(query, engine)
            self.is_online = True
            
            # Simpan ke Cache Lokal untuk Cadangan Offline
            self.df_master.to_parquet(CACHE_FILE, index=False)
            print("🟢 [ONLINE] Data Master berhasil ditarik dari Cloud Supabase & Cache diperbarui!")

        except Exception as e:
            # 2. Jika Offline / Gagal Koneksi -> Gunakan Local Cache
            self.is_online = False
            print(f"🔴 [OFFLINE MODE] Gagal terhubung ke Cloud ({str(e)}). Membaca dari Cache Lokal...")
            
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