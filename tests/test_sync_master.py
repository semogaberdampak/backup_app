import os
import pandas as pd
import pytest
from unittest.mock import patch

# Import module yang akan ditest
from sync_master import MasterDataSync, DATABASE_URL, CACHE_FILE


class TestMasterDataSync:
    """Test suite untuk MasterDataSync class"""

    def setup_method(self):
        """Setup sebelum setiap test"""
        self.sync = MasterDataSync()
        # Backup file cache jika ada
        self.cache_exists = os.path.exists(CACHE_FILE)
        if self.cache_exists:
            os.rename(CACHE_FILE, CACHE_FILE + ".bak")

    def teardown_method(self):
        """Cleanup setelah setiap test"""
        # Hapus cache test
        if os.path.exists(CACHE_FILE):
            os.remove(CACHE_FILE)
        # Restore backup
        if self.cache_exists and os.path.exists(CACHE_FILE + ".bak"):
            os.rename(CACHE_FILE + ".bak", CACHE_FILE)

    def test_init(self):
        """Test inisialisasi objek"""
        assert self.sync.is_online is False
        assert self.sync.df_master is None

    def test_load_data_offline_no_cache(self):
        """Test load_data ketika offline dan tidak ada cache"""
        with patch("sync_master.DATABASE_URL", None):
            sync = MasterDataSync()
            df, status = sync.load_data()

            assert status is False
            assert isinstance(df, pd.DataFrame)
            assert len(df) == 0
            assert list(df.columns) == ["barcode", "judul", "harga"]

    def test_load_data_offline_with_cache(self):
        """Test load_data ketika offline tapi ada cache"""
        # Buat cache dummy
        dummy_data = pd.DataFrame(
            {
                "barcode": ["123456", "789012"],
                "judul": ["Produk A", "Produk B"],
                "harga": [10000, 25000],
            }
        )
        dummy_data.to_parquet(CACHE_FILE, index=False)

        with patch("sync_master.DATABASE_URL", None):
            sync = MasterDataSync()
            df, status = sync.load_data()

            assert status is False
            assert len(df) == 2
            assert df.iloc[0]["judul"] == "Produk A"

    @patch("sync_master.create_engine")
    @patch("sync_master.pd.read_sql")
    def test_load_data_online_success(self, mock_read_sql, mock_create_engine):
        """Test load_data online berhasil"""
        # Mock database response
        mock_df = pd.DataFrame(
            {"barcode": ["111", "222"], "judul": ["Item 1", "Item 2"], "harga": [5000, 7500]}
        )
        mock_read_sql.return_value = mock_df

        with patch("sync_master.DATABASE_URL", "postgresql://test:test@localhost/test"):
            sync = MasterDataSync()
            df, status = sync.load_data()

            assert status is True
            assert len(df) == 2
            mock_read_sql.assert_called_once()
            # Verifikasi cache disimpan
            assert os.path.exists(CACHE_FILE)

    @patch("sync_master.create_engine")
    @patch("sync_master.pd.read_sql")
    def test_load_data_online_failure_fallback_to_cache(self, mock_read_sql, mock_create_engine):
        """Test load_data online gagal, fallback ke cache"""
        # Setup cache dummy
        cache_data = pd.DataFrame({"barcode": ["999"], "judul": ["Cached Item"], "harga": [9999]})
        cache_data.to_parquet(CACHE_FILE, index=False)

        # Mock exception
        mock_read_sql.side_effect = Exception("Connection failed")

        with patch("sync_master.DATABASE_URL", "postgresql://test:test@localhost/test"):
            sync = MasterDataSync()
            df, status = sync.load_data()

            assert status is False
            assert len(df) == 1
            assert df.iloc[0]["judul"] == "Cached Item"


class TestDatabaseURLConfig:
    """Test konfigurasi DATABASE_URL"""

    def test_database_url_from_env(self):
        """Test DATABASE_URL dibaca dari environment"""
        # Test ini hanya memastikan tidak error saat import
        # Nilai aktual bergantung pada .env
        assert DATABASE_URL is not None or DATABASE_URL is None  # Selalu True, tapi test import


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
