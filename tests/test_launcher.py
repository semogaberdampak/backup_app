"""Unit test launcher: sinkronisasi requirements.txt (mock — tanpa jaringan & GUI)."""
import hashlib
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import launcher


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


class TestSyncRequirements:
    """launcher.sync_requirements(): cek, verifikasi checksum, tawar install pip."""

    def _prepare(self, tmp_path, monkeypatch, isi_lokal: bytes):
        monkeypatch.chdir(tmp_path)
        (tmp_path / launcher.REQS_FILE).write_bytes(isi_lokal)
        return tmp_path

    def test_isi_sama_tidak_perlu_install(self, tmp_path, monkeypatch):
        """requirements identik dengan repo → tidak ada prompt, tidak ada pip."""
        isi = b"pandas>=2.0.0\n"
        ws = self._prepare(tmp_path, monkeypatch, isi)

        def tulis_tmp(url, dst, *a, **k):
            Path(dst).write_bytes(isi)

        with (
            patch.object(launcher.urllib.request, "urlretrieve", side_effect=tulis_tmp),
            patch.object(launcher.messagebox, "askyesno") as mock_ask,
            patch.object(launcher.subprocess, "run") as mock_run,
        ):
            launcher.sync_requirements(_sha(isi))

        mock_ask.assert_not_called()
        mock_run.assert_not_called()
        assert (ws / launcher.REQS_FILE).read_bytes() == isi
        assert not (ws / (launcher.REQS_FILE + ".new")).exists()

    def test_checksum_salah_file_rusak_ditolak(self, tmp_path, monkeypatch):
        """Checksum tidak cocok → file lokal tidak ditimpa, tidak ada prompt/pip."""
        isi_lokal = b"pandas>=2.0.0\n"
        isi_remote = b"pandas>=2.0.0\nEVIL=1\n"
        ws = self._prepare(tmp_path, monkeypatch, isi_lokal)

        def tulis_tmp(url, dst, *a, **k):
            Path(dst).write_bytes(isi_remote)

        with (
            patch.object(launcher.urllib.request, "urlretrieve", side_effect=tulis_tmp),
            patch.object(launcher.messagebox, "askyesno") as mock_ask,
            patch.object(launcher.subprocess, "run") as mock_run,
        ):
            launcher.sync_requirements(_sha(isi_lokal))  # hash lokal ≠ remote → ditolak

        mock_ask.assert_not_called()
        mock_run.assert_not_called()
        assert (ws / launcher.REQS_FILE).read_bytes() == isi_lokal
        assert not (ws / (launcher.REQS_FILE + ".new")).exists()

    def test_berubah_user_setuju_install(self, tmp_path, monkeypatch):
        """requirements berubah + user setuju → file ditimpa & pip dijalankan."""
        isi_lokal = b"pandas>=2.0.0\n"
        isi_baru = b"pandas>=2.0.0\npsycopg[binary]>=3.1.0\n"
        ws = self._prepare(tmp_path, monkeypatch, isi_lokal)

        def tulis_tmp(url, dst, *a, **k):
            Path(dst).write_bytes(isi_baru)

        with (
            patch.object(launcher.urllib.request, "urlretrieve", side_effect=tulis_tmp),
            patch.object(launcher.messagebox, "askyesno", return_value=True),
            patch.object(launcher.subprocess, "run", return_value=SimpleNamespace(returncode=0)) as mock_run,
        ):
            launcher.sync_requirements(_sha(isi_baru))

        mock_run.assert_called_once()
        # pip dijalankan dengan interpreter aktif + -r requirements.txt
        perintah = mock_run.call_args[0][0]
        assert perintah[1:3] == ["-m", "pip"]
        assert launcher.REQS_FILE in perintah
        assert (ws / launcher.REQS_FILE).read_bytes() == isi_baru

    def test_berubah_user_tolak_install(self, tmp_path, monkeypatch):
        """User menolak prompt → file lokal tetap lama, pip tidak dijalankan."""
        isi_lokal = b"pandas>=2.0.0\n"
        isi_baru = b"pandas>=2.0.0\npsycopg[binary]>=3.1.0\n"
        ws = self._prepare(tmp_path, monkeypatch, isi_lokal)

        def tulis_tmp(url, dst, *a, **k):
            Path(dst).write_bytes(isi_baru)

        with (
            patch.object(launcher.urllib.request, "urlretrieve", side_effect=tulis_tmp),
            patch.object(launcher.messagebox, "askyesno", return_value=False),
            patch.object(launcher.subprocess, "run") as mock_run,
        ):
            launcher.sync_requirements(_sha(isi_baru))

        mock_run.assert_not_called()
        assert (ws / launcher.REQS_FILE).read_bytes() == isi_lokal
        assert not (ws / (launcher.REQS_FILE + ".new")).exists()

    def test_gagal_download_mode_offline(self, tmp_path, monkeypatch):
        """Jaringan gagal → dilewati dengan tenang, file lokal tidak berubah."""
        isi_lokal = b"pandas>=2.0.0\n"
        ws = self._prepare(tmp_path, monkeypatch, isi_lokal)

        with (
            patch.object(
                launcher.urllib.request,
                "urlretrieve",
                side_effect=OSError("gagal jaringan"),
            ),
            patch.object(launcher.subprocess, "run") as mock_run,
        ):
            launcher.sync_requirements(_sha(b"isi-lain\n"))  # tidak boleh melempar error

        mock_run.assert_not_called()
        assert (ws / launcher.REQS_FILE).read_bytes() == isi_lokal