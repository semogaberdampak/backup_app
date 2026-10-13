import pytest


# Test konfigurasi dan import dasar
class TestAppConfig:
    """Test konfigurasi aplikasi"""

    def test_dotenv_loaded(self):
        """Test bahwa dotenv dapat diimport"""
        from dotenv import load_dotenv

        assert callable(load_dotenv)

    def test_supabase_config_variables(self):
        """Test variabel konfigurasi Supabase"""
        # Reload module untuk test
        import importlib
        import app_scan

        importlib.reload(app_scan)

        # Cek konstanta ada
        assert hasattr(app_scan, "SUPABASE_URL")
        assert hasattr(app_scan, "SUPABASE_KEY")
        assert hasattr(app_scan, "SUPABASE_TABLE")
        assert hasattr(app_scan, "SUPABASE_AVAILABLE")

    def test_master_data_sync_import(self):
        """Test import MasterDataSync"""
        import app_scan

        # Harus bisa import tanpa error
        assert hasattr(app_scan, "MasterDataSync")

    def test_kasir_app_class_exists(self):
        """Test kelas KasirApp tersedia"""
        import app_scan

        assert hasattr(app_scan, "KasirApp")
        assert callable(app_scan.KasirApp)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
