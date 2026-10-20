import hashlib
import os
import sys
import json
import urllib.request
import subprocess
import tkinter as tk
from tkinter import messagebox
from typing import Optional

# Konfigurasi URL mentah (raw URL) dari file version.json di GitHub Anda
VERSION_URL: str = "https://raw.githubusercontent.com/semogaberdampak/backup_app/main/version.json"
LOCAL_VERSION_FILE: str = "version.local"
MAIN_APP_FILE: str = "app_scan.py"
REQS_FILE: str = "requirements.txt"
REQS_URL: str = "https://raw.githubusercontent.com/semogaberdampak/backup_app/main/requirements.txt"


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def get_local_version() -> str:
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r") as f:
            return f.read().strip()
    return "v0.0.0"


def save_local_version(ver: str) -> None:
    with open(LOCAL_VERSION_FILE, "w") as f:
        f.write(ver)


def sync_requirements(expected_sha256: str) -> None:
    """Sinkronkan requirements.txt dari repo & tawarkan install bila isinya berbeda.

    Aman offline: gagal download dilewati tanpa menghentikan aplikasi.
    Checksum selalu diverifikasi sebelum file lokal ditimpa.
    """
    tmp_path = REQS_FILE + ".new"
    try:
        print(f"Mengecek dependensi dari {REQS_URL} ...")
        urllib.request.urlretrieve(REQS_URL, tmp_path)

        got = sha256_file(tmp_path)
        if got.lower() != expected_sha256.lower():
            os.remove(tmp_path)
            print(f"Checksum requirements tidak cocok (file rusak/disusupi): got {got}")
            return

        lokal = b""
        if os.path.exists(REQS_FILE):
            with open(REQS_FILE, "rb") as f:
                lokal = f.read()
        with open(tmp_path, "rb") as f:
            baru = f.read()

        if lokal == baru:
            os.remove(tmp_path)
            print("requirements.txt sudah terbaru, tidak ada instalasi.")
            return

        tanya: bool = messagebox.askyesno(
            "Dependensi Berubah",
            "requirements.txt di repo berubah.\n"
            "Install dependensi baru via pip sekarang?\n\n"
            "Pilih 'Tidak' jika ingin install manual kapan saja:\n"
            "pip install -r requirements.txt",
        )
        if not tanya:
            os.remove(tmp_path)
            print("Update dependensi dilewati. Install manual: pip install -r requirements.txt")
            return

        os.replace(tmp_path, REQS_FILE)
        print("Menjalankan: pip install -r requirements.txt")
        hasil = subprocess.run([sys.executable, "-m", "pip", "install", "-r", REQS_FILE])
        if hasil.returncode == 0:
            print("✅ Dependensi berhasil diupdate.")
        else:
            print("⚠️ pip mengembalikan error. Jalankan manual: pip install -r requirements.txt")
    except Exception as e:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        print(f"Gagal sinkron requirements (lanjut tanpa update dependensi): {e}")


def check_and_update() -> None:
    root = tk.Tk()
    root.withdraw()  # Sembunyikan window utama launcher sementara mengecek

    local_ver: str = get_local_version()
    print(f"Versi lokal saat ini: {local_ver}")

    try:
        # Ambil data versi terbaru dari GitHub
        req = urllib.request.urlopen(VERSION_URL, timeout=5)
        data: dict = json.loads(req.read().decode("utf-8"))

        remote_ver: Optional[str] = data.get("version")
        download_url: Optional[str] = data.get("download_url")
        changelog: str = data.get("changelog", "Pembaruan sistem.")
        expected_sha256: Optional[str] = data.get("sha256")
        # Checksum requirements.txt di repo — untuk sinkronisasi dependensi
        req_sha256: Optional[str] = data.get("requirements_sha256")
        boleh_sync: bool = False

        # Validate required fields
        if remote_ver is None or download_url is None:
            print("ERROR: Data versi dari server tidak lengkap (version/download_url kosong)")
            print("Aplikasi sudah menggunakan versi paling mutakhir.")
        elif remote_ver != local_ver:
            # Jika versi berbeda/baru, beritahu user dan download patch otomatis
            tanya: bool = messagebox.askyesno(
                "Pembaruan Tersedia (Modular Patch)",
                f"Ditemukan versi baru: {remote_ver}\nVersi Anda saat ini: {local_ver}\n\nChangelog:\n{changelog}\n\nDownload patch modular sekarang?",  # noqa: E501
            )

            if tanya:
                print(f"Mendownload patch dari {download_url}...")
                tmp_path = MAIN_APP_FILE + ".new"
                urllib.request.urlretrieve(download_url, tmp_path)
                if expected_sha256:
                    got = sha256_file(tmp_path)
                    if got.lower() != expected_sha256.lower():
                        os.remove(tmp_path)
                        messagebox.showerror(
                            "Update Gagal",
                            "Checksum patch tidak cocok (kemungkinan file rusak / disusupi). Update dibatalkan.",
                        )
                        print(f"Checksum mismatch: got {got}, expected {expected_sha256}")
                        root.destroy()
                        return
                os.replace(tmp_path, MAIN_APP_FILE)
                save_local_version(remote_ver)
                boleh_sync = True
                messagebox.showinfo(
                    "Sukses", "Patch berhasil didownload & diverifikasi! Aplikasi akan dijalankan."

                )
        else:
            print("Aplikasi sudah menggunakan versi paling mutakhir.")
            boleh_sync = True

        # Sinkronisasi requirements.txt bila checksum-nya berbeda dari repo
        # (hanya saat patch diterima / versi sudah sama — tidak saat user menolak update)
        if boleh_sync and req_sha256:
            lokal_req_sha: str = sha256_file(REQS_FILE) if os.path.exists(REQS_FILE) else ""
            if lokal_req_sha.lower() != req_sha256.lower():
                sync_requirements(req_sha256)
            else:
                print("requirements.txt sudah sesuai versi repo.")

    except Exception as e:
        print(f"Gagal memeriksa update dari GitHub (Mode Offline): {e}")
        # Jika gagal koneksi/offline, tetap jalankan aplikasi eksisting jika ada
        if not os.path.exists(MAIN_APP_FILE):
            messagebox.showerror(
                "Error",
                f"File utama {MAIN_APP_FILE} tidak ditemukan dan gagal terhubung ke server update!",
            )
            sys.exit()

    root.destroy()

    # Jalankan aplikasi utama (app_scan.py)
    if os.path.exists(MAIN_APP_FILE):
        subprocess.run([sys.executable, MAIN_APP_FILE])
    else:
        messagebox.showerror("Error", f"File {MAIN_APP_FILE} tidak ada.")


if __name__ == "__main__":
    check_and_update()
