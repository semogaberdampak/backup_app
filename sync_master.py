import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import Engine
import os
import socket
from dotenv import load_dotenv
from typing import Optional, Tuple

# Load environment variables
load_dotenv()

# =========================================================================
# KONFIGURASI DATABASE DARI ENVIRONMENT VARIABLES
# =========================================================================
# Prioritas 1: DATABASE_URL lengkap (recommended untuk Supabase pooler)
# Prioritas 2: Komponen terpisah (DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASS)
# ==========================================================================

# --- Helper resolusi host: pilih alamat yang SILAKAN DITERUSKAN (IPv6/Ipv4) ---
def _resolve_db_host(host: str, port: int = 5432) -> str:
    """Pilih alamat yang dapat dihubungi. Mencegah kegagalan Network is unreachable
    akibat kombinasi stack IPv4/IPv6 yang tidak cocok di environment host.

    Prioritas: alamat yang berhasil ditunjukkan resolve -> hostname asli.
    """
    addrs = []
    try:
        addrs = socket.getaddrinfo(host, port, socket.AF_UNSPEC, socket.SOCK_STREAM)
    except socket.gaierror:
        addrs = []

    reachable = None
    for fam, _typ, _proto, _canon, sa in addrs:
        ip = sa[0]
        s = socket.socket(fam, socket.SOCK_STREAM)
        s.settimeout(2.0)
        try:
            s.connect((ip, port))
            reachable = ip
            s.close()
            break
        except OSError:
            try:
                s.close()
            except OSError:
                pass

    if reachable:
        return reachable

    # Fallback: coba hostname asli (resolusi dapat berbeda antar runtime)
    return host


# ==========================================================================
# KONFIGURASI DATABASE DARI ENVIRONMENT VARIABLES
# ==========================================================================
# Prioritas 1: DATABASE_URL lengkap (recommended untuk Supabase pooler)
# Prioritas 2: Komponen terpisah (DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASS)
# ==========================================================================
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
        1) PostgreSQL (jika URL tersedia) -> online
        2) Supabase REST (apikey anon) -> online
        3) Cache lokal offline
        """
        # --- 1. coba koneksi database langsung (jika adapter tersedia) ---
        if DATABASE_URL:
            try:
                engine: Engine = create_engine(DATABASE_URL, connect_args={"connect_timeout": 5})
                query: str = "SELECT barcode, judul, harga FROM master_produk"
                self.df_master = pd.read_sql(query, engine)
                self.is_online = True

                self.df_master.to_parquet(CACHE_FILE, index=False)
                print(
                    "🟢 [ONLINE] Data Master berhasil ditarik dari Cloud Supabase & Cache diperbarui!"
                )
                return self.df_master, self.is_online
            except Exception as e:
                print(
                    f"🔴 [OFFLINE MODE] Gagal terhubung ke Cloud ({str(e)}). Silakan coba REST/offline..."
                )

        # --- 2. fallback: ambil via REST (Supabase HTTP) ---
        df = self._fetch_master_rest()
        if len(df):
            df.to_parquet(CACHE_FILE, index=False)
            self.is_online = True
            self.df_master = df
            print(
                "🟢 [ONLINE] Data Master berhasil ditarik via REST Supabase & Cache diperbarui!"
            )
            return self.df_master, self.is_online

        # --- 3. mode OFFLINE: pakai cache lokal ---
        if os.path.exists(CACHE_FILE):
            self.df_master = pd.read_parquet(CACHE_FILE)
            print("🟡 Data Master berhasil dimuat dari Cache Lokal!")
            return self.df_master, self.is_online

        self.df_master = pd.DataFrame(columns=["barcode", "judul", "harga"])
        print("⚠️ Tidak ada Cache Lokal ditemukan. Data master kosong.")
        return self.df_master, self.is_online

    def _fetch_master_rest(self) -> pd.DataFrame:
        """Ambil master produk lewat REST Supabase (HTTP). Aman kalau Postgres pooler tidak ada jarinya."""
        try:
            from supabase import create_client  # type: ignore

            url = os.getenv("SUPABASE_URL")
            key = os.getenv("SUPABASE_KEY")
            if not url or not key:
                print("⚠️ SUPABASE_URL/SUPABASE_KEY tidak ada di .env. Mode OFFLINE.")
                return pd.DataFrame(columns=["barcode", "judul", "harga"])

            client = create_client(url, key)
            res = client.table("master_produk").select("barcode, judul, harga").execute()
            rows = res.data or []
            df = pd.DataFrame(rows)
            if len(df):
                df["harga"] = pd.to_numeric(df["harga"], errors="coerce").fillna(0.0)
            return df
        except Exception as e:
            logging.warning(f"Gagal ambil master via REST Supabase: {e}")
            return pd.DataFrame(columns=["barcode", "judul", "harga"])


if __name__ == "__main__":
    # Uji Coba Modul Sync
    sync = MasterDataSync()
    df, status = sync.load_data()
    print()
    print("--- Hasil Uji Coba Sync ---")
    print("Status Online:", status)
    print("Jumlah Produk Terload:", len(df))
    print(df.head())




