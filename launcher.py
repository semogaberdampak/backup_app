import os
import sys
import json
import urllib.request
import subprocess
import tkinter as tk
from tkinter import messagebox

# Konfigurasi URL mentah (raw URL) dari file version.json di GitHub Anda
VERSION_URL = "https://raw.githubusercontent.com/USERNAME_ANDA/REPO_ANDA/main/version.json"
LOCAL_VERSION_FILE = "version.local"
MAIN_APP_FILE = "app_scan.py"

def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r") as f:
            return f.read().strip()
    return "v0.0.0"

def save_local_version(ver):
    with open(LOCAL_VERSION_FILE, "w") as f:
        f.write(ver)

def check_and_update():
    root = tk.Tk()
    root.withdraw() # Sembunyikan window utama launcher sementara mengecek
    
    local_ver = get_local_version()
    print(Versi lokal saat ini: {local_ver})

    try:
        # Ambil data versi terbaru dari GitHub
        req = urllib.request.urlopen(VERSION_URL, timeout=5)
        data = json.loads(req.read().decode('utf-8'))
        
        remote_ver = data.get("version")
        download_url = data.get("download_url")
        changelog = data.get("changelog", "Pembaruan sistem.")

        if remote_ver != local_ver:
            # Jika versi berbeda/baru, beritahu user dan download patch otomatis
            tanya = messagebox.askyesno(
                "Pembaruan Tersedia (Modular Patch)", 
                f"Ditemukan versi baru: {remote_ver}\nVersi Anda saat ini: {local_ver}\n\nChangelog:\n{changelog}\n\nDownload patch modular sekarang?"
            )
            
            if tanya:
                print(Mendownload patch dari {download_url}...)
                urllib.request.urlretrieve(download_url, MAIN_APP_FILE)
                save_local_version(remote_ver)
                messagebox.showinfo("Sukses", "Patch berhasil didownload! Aplikasi akan segera dijalankan.")
        else:
            print("Aplikasi sudah menggunakan versi paling mutakhir.")

    except Exception as e:
        print("Gagal memeriksa update dari GitHub (Mode Offline):", e)
        # Jika gagal koneksi/offline, tetap jalankan aplikasi eksisting jika ada
        if not os.path.exists(MAIN_APP_FILE):
            messagebox.showerror("Error", f"File utama {MAIN_APP_FILE} tidak ditemukan dan gagal terhubung ke server update!")
            sys.exit()

    root.destroy()
    
    # Jalankan aplikasi utama (app_scan.py)
    if os.path.exists(MAIN_APP_FILE):
        subprocess.run([sys.executable, MAIN_APP_FILE])
    else:
        messagebox.showerror("Error", f"File {MAIN_APP_FILE} tidak ada.")

if __name__ == "__main__":
    check_and_update()
