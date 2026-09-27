import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from datetime import datetime
import os
from dotenv import load_dotenv
import threading
import tempfile
import json
import logging
import subprocess
import pathlib
import copy

load_dotenv()

# ==========================================================
# 1. KONFIGURASI SUPABASE
# ==========================================================
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
SUPABASE_TABLE = os.getenv("SUPABASE_TABLE", "master_produk")

if not SUPABASE_URL or not SUPABASE_KEY:
    print("❌ ERROR: SUPABASE_URL dan SUPABASE_KEY harus diatur di file .env")

try:
    from supabase import create_client, Client
    SUPABASE_AVAILABLE = True
    print("✅ Library Supabase berhasil diimport")
except ImportError:
    SUPABASE_AVAILABLE = False
    print("❌ PERINGATAN: Library 'supabase' tidak ditemukan! Jalankan: pip install supabase")

# ============================================================================
# 2. SAFE IMPORT
# ============================================================================
try:
    from sync_master import MasterDataSync
except ImportError:
    print("PERINGATAN: 'sync_master.py' tidak ditemukan. Menggunakan mode Offline/Dummy.")
    class MasterDataSync:
        def load_data(self):
            return pd.DataFrame(columns=['barcode', 'judul', 'harga']), False

# ============================================================================
# 3. Setup Logging
# ============================================================================
def setup_app_logging():
    try:
        log_dir = pathlib.Path.home() / ".takom_kasir"
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / "app_error.log"
        logging.basicConfig(
            filename=str(log_file),
            level=logging.ERROR,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )
    except Exception:
        logging.basicConfig(
            filename='app_error.log',
            level=logging.ERROR,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )


class KasirApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Takom Kasir v1.8.0 - Tampilan UI baru, Nota Editor Adjustable")
        self.root.state('zoomed')
        try:
            self.root.iconbitmap("logo.ico")
        except Exception:
            pass

        self.sync_tool = MasterDataSync()
        self.df_master = pd.DataFrame()
        self.is_online = False
        self.is_loading = True

        # --- Supabase Client ---
        self.is_supabase_ready = False
        self.supabase = None
        if SUPABASE_AVAILABLE and "your-project-id" not in SUPABASE_URL:
            try:
                self.supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
                self.is_supabase_ready = True
                print("✅ Supabase Client berhasil diinisialisasi")
            except Exception as e:
                print(f"❌ Gagal koneksi ke Supabase: {e}")
                self.is_supabase_ready = False

        # File & Sheet Target
        self.target_file_path = ""
        self.wb_target = None

        # Cache laporan
        self.cached_report_df = None
        self.cached_sheet_name = None

        # NOMAN
        self.noman_data = {}
        self.noman_item_map = {}
        self.noman_entry_widget = None

        # Pagination Data Master
        self.dm_items_per_page = 20
        self.dm_current_page = 1
        self.dm_total_pages = 1
        self.dm_total_records = 0
        self.dm_footer_buttons = []

        # ============================================================
        # TEMPLATE NOTA KUSTOM
        # ============================================================
        self.template_nota = {
            "rows": [
                {"type": "empty", "content": "", "align": "center"},
                {"type": "text", "content": "TAKOM KANISIUS", "align": "center", "bold": True},
                {"type": "text", "content": "SANTO ANTONIUS PURBAYAN", "align": "center", "bold": False},
                {"type": "text", "content": "Jl. Arifin No 1 Surakarta", "align": "center", "bold": False},
                {"type": "text", "content": "Telp: 0823-3764-1713", "align": "center", "bold": False},
                {"type": "line", "content": "-", "align": "full"},
                {"type": "header_trx", "content": "", "align": "left"},
                {"type": "line", "content": "-", "align": "full"},
                {"type": "products", "content": "", "align": "left"},
                {"type": "line", "content": "-", "align": "full"},
                {"type": "total", "content": "", "align": "right"},
                {"type": "payment", "content": "", "align": "right"},
                {"type": "line", "content": "-", "align": "full"},
                {"type": "text", "content": "Terima Kasih Atas Kunjungan Anda", "align": "center", "bold": False},
                {"type": "text", "content": "Barang yang sudah dibeli tidak dapat ditukar/dikembalikan.", "align": "center", "bold": False},
            ]
        }
        self.load_template_nota()

        # Konfigurasi Dasar Nota
        self.config_nota = {
            "nama_toko": "TAKOM KANISIUS",
            "sub_nama": "SANTO ANTONIUS PURBAYAN",
            "alamat_toko": "Jl. Arifin No 1 Surakarta",
            "telepon_toko": "Telp: 0823-3764-1713",
            "nama_kasir": "Admin",
            "info_tambahan": "Terima Kasih Atas Kunjungan Anda",
            "pesan_penutup": "Barang yang sudah dibeli tidak dapat ditukar/dikembalikan.",
            "lebar_kertas": 48,
            "ukuran_font": 10,
            "tampilkan_logo_teks": True,
            "tampilkan_footer": True,
            "tampilkan_telp": True
        }
        self.load_config_nota_json()

        # Keranjang
        self.cart = []
        self.no_trx_counter = 1

        self.setup_custom_styles()
        self.setup_main_layout()
        threading.Thread(target=self.load_master_data_async, daemon=True).start()

    # =========================================================================
    # Helper path
    # =========================================================================
    def _get_app_dir(self):
        app_dir = pathlib.Path.home() / ".takom_kasir"
        app_dir.mkdir(exist_ok=True)
        return app_dir

    def _get_config_path(self):
        return self._get_app_dir() / "config_nota.json"

    def _get_legacy_config_path(self):
        return pathlib.Path("config_nota.json")

    def _get_template_path(self):
        return self._get_app_dir() / "template_nota.json"

    def _show_safe_error(self, parent, title, user_message, exception=None, context=""):
        if exception is not None:
            log_msg = f"{context} | {type(exception).__name__}: {exception}"
            logging.error(log_msg, exc_info=True)
        messagebox.showerror(title, user_message, parent=parent)

    # =========================================================================
    # Load/Save Template Nota
    # =========================================================================
    def load_template_nota(self):
        path = self._get_template_path()
        try:
            if path.exists():
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if "rows" in data and isinstance(data["rows"], list):
                        self.template_nota = data
                        return
        except Exception as e:
            logging.error(f"Gagal memuat template nota: {e}", exc_info=True)

    def save_template_nota(self):
        try:
            path = self._get_template_path()
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.template_nota, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            logging.error(f"Gagal menyimpan template nota: {e}", exc_info=True)
            return False

    # =========================================================================
    # Load/Save config nota
    # =========================================================================
    def load_config_nota_json(self):
        new_path = self._get_config_path()
        legacy_path = self._get_legacy_config_path()
        try:
            if new_path.exists():
                with open(new_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for key, value in data.items():
                        if key in self.config_nota:
                            self.config_nota[key] = value
                return
            if legacy_path.exists():
                with open(legacy_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for key, value in data.items():
                        if key in self.config_nota:
                            self.config_nota[key] = value
                try:
                    self.save_config_nota_json()
                except Exception:
                    pass
        except Exception as e:
            logging.error(f"Gagal memuat config nota: {e}", exc_info=True)

    def save_config_nota_json(self):
        try:
            config_path = self._get_config_path()
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.config_nota, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Gagal menyimpan config nota: {e}", exc_info=True)

    # =========================================================================
    # STYLES UNIFIED - RESPONSIVE
    # =========================================================================
    def setup_custom_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        # === BACKGROUND GLOBAL ===
        style.configure(".", background="#f5f6fa", fieldbackground="#ffffff")
        style.configure("TFrame", background="#f5f6fa")
        style.configure("TLabelFrame", background="#f5f6fa", foreground="#2c3e50")
        style.configure("TLabelFrame.Label", font=("Segoe UI", 10, "bold"), foreground="#2c3e50")
        style.configure("TLabel", background="#f5f6fa", foreground="#2c3e50")

        # === TOMBOL TTK - TEKS HITAM JELAS (FIX BUG TIDAK TERBACA) ===
        style.configure("TButton",
                        font=("Segoe UI", 10, "bold"),
                        foreground="#1a1a2e",
                        background="#ffffff",
                        padding=8)
        style.map("TButton",
                  background=[('active', '#e0e0e0'), ('pressed', '#d0d0d0')],
                  foreground=[('active', '#1a1a2e'), ('pressed', '#1a1a2e')])

        # === TABEL UNIFIED ===
        for tbl_name in ["Unified.Treeview", "DataMaster.Treeview", "DarkReport.Treeview", "Template.Treeview"]:
            style.configure(tbl_name,
                            background="#ffffff",
                            foreground="#2c3e50",
                            fieldbackground="#ffffff",
                            rowheight=28,
                            font=("Segoe UI", 10))
            style.configure(f"{tbl_name}.Heading",
                            background="#2c3e50",
                            foreground="#ffffff",
                            font=("Segoe UI", 10, "bold"),
                            relief="flat")
            style.map(tbl_name,
                      background=[('selected', '#3498db')],
                      foreground=[('selected', '#ffffff')])

        # === COMBOBOX ===
        style.configure("TCombobox", font=("Segoe UI", 10))
        style.map("TCombobox",
                  fieldbackground=[('readonly', '#ffffff')],
                  selectbackground=[('readonly', '#3498db')])

        # === ENTRY ===
        style.configure("TEntry", font=("Segoe UI", 10))

        # === CHECKBUTTON & RADIOBUTTON ===
        style.configure("TCheckbutton", background="#f5f6fa", font=("Segoe UI", 9))
        style.configure("TRadiobutton", background="#f5f6fa", font=("Segoe UI", 9))

    def center_popup(self, popup, width, height):
        popup.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        # Pastikan popup tidak melebihi layar
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        width = min(width, screen_w - 40)
        height = min(height, screen_h - 40)
        x = root_x + (root_w // 2) - (width // 2)
        y = root_y + (root_h // 2) - (height // 2)
        x = max(10, min(x, screen_w - width - 10))
        y = max(10, min(y, screen_h - height - 10))
        popup.geometry(f'{width}x{height}+{x}+{y}')

    def create_rounded_button(self, parent, text, command, bg_color, fg_color):
        canvas = tk.Canvas(parent, width=175, height=38, bg="#1e2a38", highlightthickness=0, cursor="hand2")

        def draw_button(color, text_color, border_color=None):
            canvas.delete("all")
            x1, y1, x2, y2, r = 2, 2, 173, 36, 10
            if border_color:
                canvas.create_arc(x1-1, y1-1, x1 + 2*r+1, y1 + 2*r+1, start=90, extent=90, fill=border_color, outline=border_color)
                canvas.create_arc(x2 - 2*r-1, y1-1, x2+1, y1 + 2*r+1, start=0, extent=90, fill=border_color, outline=border_color)
                canvas.create_arc(x1-1, y2 - 2*r-1, x1 + 2*r+1, y2+1, start=180, extent=90, fill=border_color, outline=border_color)
                canvas.create_arc(x2 - 2*r-1, y2 - 2*r-1, x2+1, y2+1, start=270, extent=90, fill=border_color, outline=border_color)
                canvas.create_rectangle(x1 + r, y1-1, x2 - r, y2+1, fill=border_color, outline=border_color)
                canvas.create_rectangle(x1-1, y1 + r, x2+1, y2 - r, fill=border_color, outline=border_color)
            canvas.create_arc(x1, y1, x1 + 2*r, y1 + 2*r, start=90, extent=90, fill=color, outline=color)
            canvas.create_arc(x2 - 2*r, y1, x2, y1 + 2*r, start=0, extent=90, fill=color, outline=color)
            canvas.create_arc(x1, y2 - 2*r, x1 + 2*r, y2, start=180, extent=90, fill=color, outline=color)
            canvas.create_arc(x2 - 2*r, y2 - 2*r, x2, y2, start=270, extent=90, fill=color, outline=color)
            canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=color, outline=color)
            canvas.create_rectangle(x1, y1 + r, x2, y2 - r, fill=color, outline=color)
            canvas.create_text(87, 19, text=text, fill=text_color, font=("Segoe UI", 9, "bold"))

        canvas.draw = draw_button
        canvas.draw(bg_color, fg_color)
        canvas.command = command
        canvas.configure(takefocus=1)
        canvas.bind("<Button-1>", lambda e: command())
        canvas.bind("<Return>", lambda e: command())
        canvas.bind("<space>", lambda e: command())
        canvas.bind("<Enter>", lambda e: canvas.draw("#16a085" if bg_color == "#1abc9c" else "#3d566e", fg_color))
        canvas.bind("<Leave>", lambda e: canvas.draw(bg_color, fg_color))
        canvas.bind("<FocusIn>", lambda e: canvas.draw("#16a085" if bg_color == "#1abc9c" else "#3d566e", fg_color))
        canvas.bind("<FocusOut>", lambda e: canvas.draw(bg_color, fg_color))
        return canvas

    def setup_main_layout(self):
        self.toolbar_frame = tk.Frame(self.root, bg="#1e2a38", height=56)
        self.toolbar_frame.pack(side=tk.TOP, fill=tk.X)
        self.toolbar_frame.pack_propagate(False)

        # === TOMBOL TAB UTAMA ===
        self.btn_transaksi = self.create_rounded_button(
            self.toolbar_frame, "🛒 Transaksi Kasir",
            lambda: self.switch_tab("transaksi"), "#1abc9c", "white"
        )
        self.btn_transaksi.pack(side=tk.LEFT, padx=8, pady=10)

        self.btn_data_master = self.create_rounded_button(
            self.toolbar_frame, "📋 Data Master",
            lambda: self.switch_tab("data_master"), "#2c3e50", "white"
        )
        self.btn_data_master.pack(side=tk.LEFT, padx=4, pady=10)

        self.btn_laporan = self.create_rounded_button(
            self.toolbar_frame, "📊 Laporan & Sheet",
            lambda: self.switch_tab("laporan"), "#2c3e50", "white"
        )
        self.btn_laporan.pack(side=tk.LEFT, padx=4, pady=10)

        # === SEPARATOR ===
        sep = tk.Frame(self.toolbar_frame, bg="#34495e", width=2)
        sep.pack(side=tk.LEFT, padx=15, pady=10, fill="y")

        # === TOMBOL AKSI ===
        self.btn_tambah_produk = self.create_rounded_button(
            self.toolbar_frame, "➕ Tambah Produk (F2)",
            self.buka_popup_tambah_produk, "#34495e", "white"
        )
        self.btn_tambah_produk.pack(side=tk.RIGHT, padx=4, pady=10)

        self.btn_ubah_harga = self.create_rounded_button(
            self.toolbar_frame, "✏️ Ubah Harga",
            self.buka_popup_ubah_harga, "#34495e", "white"
        )
        self.btn_ubah_harga.pack(side=tk.RIGHT, padx=4, pady=10)

        self.btn_editor_template = self.create_rounded_button(
            self.toolbar_frame, "🎨 Editor Nota",
            self.buka_popup_editor_template_nota, "#34495e", "white"
        )
        self.btn_editor_template.pack(side=tk.RIGHT, padx=4, pady=10)

        # === CONTAINER HALAMAN ===
        self.container = ttk.Frame(self.root)
        self.container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.frame_transaksi = ttk.Frame(self.container)
        self.frame_data_master = ttk.Frame(self.container)
        self.frame_laporan = ttk.Frame(self.container)

        self.setup_halaman_transaksi(self.frame_transaksi)
        self.setup_halaman_data_master(self.frame_data_master)
        self.setup_halaman_laporan(self.frame_laporan)

        self.switch_tab("transaksi")
        self.setup_global_keyboard()

    def switch_tab(self, tab_name):
        self.frame_transaksi.pack_forget()
        self.frame_data_master.pack_forget()
        self.frame_laporan.pack_forget()

        self.btn_transaksi.draw("#2c3e50", "white")
        self.btn_data_master.draw("#2c3e50", "white")
        self.btn_laporan.draw("#2c3e50", "white")

        if tab_name == "transaksi":
            self.frame_transaksi.pack(fill=tk.BOTH, expand=True)
            self.btn_transaksi.draw("#1abc9c", "white")
            self.ent_search.focus_force()
        elif tab_name == "data_master":
            self.frame_data_master.pack(fill=tk.BOTH, expand=True)
            self.btn_data_master.draw("#1abc9c", "white")
            self.refresh_data_master_view()
        elif tab_name == "laporan":
            self.frame_laporan.pack(fill=tk.BOTH, expand=True)
            self.btn_laporan.draw("#1abc9c", "white")
            self.refresh_data_laporan()

    # =========================================================================
    # HALAMAN DATA MASTER
    # =========================================================================
    def setup_halaman_data_master(self, parent):
        top_frame = tk.Frame(parent, bg="#f5f6fa")
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        ttk.Label(top_frame, text="📋 Database Master Produk",
                  font=("Segoe UI", 13, "bold"), foreground="#2c3e50").pack(side="left", padx=5)

        self.dm_lbl_info = ttk.Label(top_frame, text="Memuat data...",
                                     font=("Segoe UI", 9, "italic"), foreground="gray")
        self.dm_lbl_info.pack(side="right", padx=10)

        filter_frame = tk.Frame(parent, bg="#f5f6fa")
        filter_frame.pack(fill="x", padx=10, pady=(0, 5))

        ttk.Label(filter_frame, text=" Filter:", font=("Segoe UI", 9, "bold")).pack(side="left", padx=5)
        self.dm_ent_filter = ttk.Entry(filter_frame, font=("Segoe UI", 10), width=35)
        self.dm_ent_filter.pack(side="left", padx=5)
        self.dm_ent_filter.bind("<KeyRelease>", lambda e: self.dm_apply_filter())

        btn_refresh_dm = ttk.Button(filter_frame, text="🔄 Refresh dari Cloud",
                                    command=self.dm_force_refresh_cloud)
        btn_refresh_dm.pack(side="right", padx=5)

        table_frame = ttk.Frame(parent)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        dm_columns = ("no", "barcode", "judul", "harga")
        self.dm_tree = ttk.Treeview(table_frame, columns=dm_columns, show="headings",
                                    selectmode="browse", style="Unified.Treeview")

        self.dm_tree.heading("no", text="No.")
        self.dm_tree.heading("barcode", text="Barcode")
        self.dm_tree.heading("judul", text="Judul Produk")
        self.dm_tree.heading("harga", text="Harga (Rp)")

        self.dm_tree.column("no", width=60, anchor="center", stretch=False)
        self.dm_tree.column("barcode", width=160, anchor="w", stretch=False)
        self.dm_tree.column("judul", width=500, anchor="w", stretch=True)
        self.dm_tree.column("harga", width=130, anchor="e", stretch=False)

        sb_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.dm_tree.yview)
        sb_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.dm_tree.xview)
        self.dm_tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.dm_tree.pack(side="left", fill="both", expand=True)

        self.dm_footer_frame = tk.Frame(parent, bg="#2c3e50", height=42)
        self.dm_footer_frame.pack(side="bottom", fill="x", padx=0, pady=0)
        self.dm_footer_frame.pack_propagate(False)

        self.dm_lbl_page_info = tk.Label(self.dm_footer_frame,
                                         text="Halaman 1 / 1  |  Total: 0 produk",
                                         font=("Segoe UI", 9, "bold"),
                                         bg="#2c3e50", fg="#ecf0f1")
        self.dm_lbl_page_info.pack(side="left", padx=15, pady=8)

        self.dm_btn_container = tk.Frame(self.dm_footer_frame, bg="#2c3e50")
        self.dm_btn_container.pack(side="right", padx=10, pady=4)

    def refresh_data_master_view(self):
        if self.df_master is None or self.df_master.empty:
            self.dm_total_records = 0
            self.dm_total_pages = 1
            self.dm_current_page = 1
            self.dm_lbl_info.config(text="⚠️ Belum ada data. Tunggu sync selesai...",
                                    foreground="orange")
            for item in self.dm_tree.get_children():
                self.dm_tree.delete(item)
            self.dm_render_footer()
            return

        filter_text = self.dm_ent_filter.get().strip().lower() if hasattr(self, 'dm_ent_filter') else ""
        if filter_text:
            df_filtered = self.df_master[
                self.df_master['barcode'].astype(str).str.lower().str.contains(filter_text, na=False) |
                self.df_master['judul'].astype(str).str.lower().str.contains(filter_text, na=False)
            ]
        else:
            df_filtered = self.df_master

        self.dm_total_records = len(df_filtered)
        self.dm_total_pages = max(1, (self.dm_total_records + self.dm_items_per_page - 1) // self.dm_items_per_page)

        if self.dm_current_page > self.dm_total_pages:
            self.dm_current_page = self.dm_total_pages

        self.dm_render_page(df_filtered)
        self.dm_render_footer()

        status = "🟢 ONLINE" if self.is_online else "🟡 OFFLINE"
        self.dm_lbl_info.config(
            text=f"{status} | Menampilkan {self.dm_total_records:,} dari {len(self.df_master):,} produk",
            foreground="green" if self.is_online else "orange"
        )

    def dm_render_page(self, df_filtered=None):
        for item in self.dm_tree.get_children():
            self.dm_tree.delete(item)

        if df_filtered is None:
            if self.df_master is None or self.df_master.empty:
                return
            filter_text = self.dm_ent_filter.get().strip().lower() if hasattr(self, 'dm_ent_filter') else ""
            if filter_text:
                df_filtered = self.df_master[
                    self.df_master['barcode'].astype(str).str.lower().str.contains(filter_text, na=False) |
                    self.df_master['judul'].astype(str).str.lower().str.contains(filter_text, na=False)
                ]
            else:
                df_filtered = self.df_master

        if df_filtered.empty:
            return

        start_idx = (self.dm_current_page - 1) * self.dm_items_per_page
        end_idx = min(start_idx + self.dm_items_per_page, len(df_filtered))
        page_df = df_filtered.iloc[start_idx:end_idx]

        for i, (_, row) in enumerate(page_df.iterrows()):
            global_no = start_idx + i + 1
            barcode_val = str(row.get('barcode', '')).strip()
            if barcode_val.endswith('.0') and barcode_val[:-2].isdigit():
                barcode_val = barcode_val[:-2]
            judul_val = str(row.get('judul', '')).strip()
            try:
                harga_val = float(row.get('harga', 0))
                harga_str = f"Rp {harga_val:,.0f}"
            except (ValueError, TypeError):
                harga_str = "Rp 0"

            self.dm_tree.insert("", "end", values=(global_no, barcode_val, judul_val, harga_str))

    def dm_render_footer(self):
        for btn in self.dm_footer_buttons:
            try:
                btn.destroy()
            except Exception:
                pass
        self.dm_footer_buttons.clear()

        self.dm_lbl_page_info.config(
            text=f"Halaman {self.dm_current_page} / {self.dm_total_pages}  |  "
                 f"Total: {self.dm_total_records:,} produk  |  "
                 f"{self.dm_items_per_page} baris/halaman"
        )

        if self.dm_total_pages <= 1:
            return

        def make_page_btn(parent, text, page_num, is_current=False):
            bg = "#1abc9c" if is_current else "#34495e"
            fg = "#ffffff"
            btn = tk.Button(parent, text=text, font=("Segoe UI", 9, "bold"),
                            bg=bg, fg=fg, activebackground="#16a085", activeforeground="white",
                            relief="flat", padx=8, pady=2, cursor="hand2",
                            command=lambda p=page_num: self.dm_go_to_page(p))
            btn.pack(side="left", padx=2)
            self.dm_footer_buttons.append(btn)
            return btn

        def make_label(parent, text):
            lbl = tk.Label(parent, text=text, font=("Segoe UI", 9, "bold"),
                           bg="#2c3e50", fg="#95a5a6")
            lbl.pack(side="left", padx=2)
            self.dm_footer_buttons.append(lbl)

        if self.dm_current_page > 1:
            make_page_btn(self.dm_btn_container, "◀ Prev", self.dm_current_page - 1)

        pages_to_show = []
        total = self.dm_total_pages
        cur = self.dm_current_page

        pages_to_show.append(1)
        start = max(2, cur - 2)
        end = min(total - 1, cur + 2)

        if start > 2:
            pages_to_show.append("...")
        for p in range(start, end + 1):
            pages_to_show.append(p)
        if end < total - 1:
            pages_to_show.append("...")
        if total > 1:
            pages_to_show.append(total)

        seen = set()
        unique_pages = []
        for p in pages_to_show:
            if p not in seen:
                seen.add(p)
                unique_pages.append(p)

        for p in unique_pages:
            if p == "...":
                make_label(self.dm_btn_container, " ... ")
            else:
                make_page_btn(self.dm_btn_container, str(p), p, is_current=(p == cur))

        if self.dm_current_page < self.dm_total_pages:
            make_page_btn(self.dm_btn_container, "Next ▶", self.dm_current_page + 1)

    def dm_go_to_page(self, page_num):
        if page_num < 1 or page_num > self.dm_total_pages:
            return
        self.dm_current_page = page_num
        self.refresh_data_master_view()

    def dm_apply_filter(self):
        self.dm_current_page = 1
        self.refresh_data_master_view()

    def dm_force_refresh_cloud(self):
        if self.is_loading:
            messagebox.showinfo("Info", "Proses loading masih berjalan. Mohon tunggu.")
            return
        self.dm_lbl_info.config(text=" Sedang menyinkronkan ulang dari cloud...",
                                foreground="blue")
        self.root.update_idletasks()
        threading.Thread(target=self.load_master_data_async, daemon=True).start()

    # =========================================================================
    # EDITOR TEMPLATE NOTA - UKURAN PROPORSIONAL
    # =========================================================================
    def buka_popup_editor_template_nota(self):
        popup = tk.Toplevel(self.root)
        popup.title(" Editor Template Nota Kustom")
        # UKURAN PROPORSIONAL - tidak full screen, tombol bawah tetap terlihat
        self.center_popup(popup, 1100, 680)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: popup.destroy())

        TIPE_BARIS = {
            "empty": "Baris Kosong",
            "text": "Teks Bebas",
            "line": "Garis Pemisah",
            "header_trx": "Header Transaksi (No Trx/Kasir/Waktu)",
            "products": "Daftar Produk (Otomatis)",
            "total": "Total Belanja (Otomatis)",
            "payment": "Pembayaran & Kembalian (Otomatis)",
            "footer_text": "Teks Footer (dari config)",
        }

        ALIGN_OPTIONS = {
            "left": "Left (Rata Kiri)",
            "center": "Center (Rata Tengah)",
            "right": "Right (Rata Kanan)",
            "full": "Full (Sepanjang Lebar)",
        }

        working_template = copy.deepcopy(self.template_nota)

        # === Panel Utama: Editor (kiri) + Preview (kanan) ===
        main_split = ttk.Frame(popup)
        main_split.pack(fill="both", expand=True, padx=10, pady=(10, 5))

        # === PANEL EDITOR (KIRI) ===
        frame_editor = ttk.LabelFrame(main_split, text=" 📝 Daftar Baris Template ")
        frame_editor.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        toolbar_ed = ttk.Frame(frame_editor)
        toolbar_ed.pack(fill="x", padx=5, pady=5)

        btn_tambah = ttk.Button(toolbar_ed, text="➕ Tambah Baris")
        btn_tambah.pack(side="left", padx=2)
        btn_hapus = ttk.Button(toolbar_ed, text="🗑️ Hapus")
        btn_hapus.pack(side="left", padx=2)
        btn_naik = ttk.Button(toolbar_ed, text="⬆️ Naik")
        btn_naik.pack(side="left", padx=2)
        btn_turun = ttk.Button(toolbar_ed, text="⬇️ Turun")
        btn_turun.pack(side="left", padx=2)
        btn_duplikat = ttk.Button(toolbar_ed, text="📋 Duplikat")
        btn_duplikat.pack(side="left", padx=2)
        btn_reset = ttk.Button(toolbar_ed, text="🔄 Reset Default")
        btn_reset.pack(side="right", padx=2)
        btn_setting = ttk.Button(toolbar_ed, text="⚙️ Pengaturan Dasar")
        btn_setting.pack(side="right", padx=2)

        tree_frame = ttk.Frame(frame_editor)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        cols = ("no", "tipe", "content", "align")
        tree_tpl = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                selectmode="browse", style="Template.Treeview")
        tree_tpl.heading("no", text="No")
        tree_tpl.heading("tipe", text="Tipe Baris")
        tree_tpl.heading("content", text="Isi / Teks")
        tree_tpl.heading("align", text="Alignment")
        tree_tpl.column("no", width=50, anchor="center", stretch=False)
        tree_tpl.column("tipe", width=200, anchor="w", stretch=False)
        tree_tpl.column("content", width=400, anchor="w", stretch=True)
        tree_tpl.column("align", width=100, anchor="center", stretch=False)

        sb_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree_tpl.yview)
        tree_tpl.configure(yscrollcommand=sb_y.set)
        sb_y.pack(side="right", fill="y")
        tree_tpl.pack(side="left", fill="both", expand=True)

        panel_edit = ttk.LabelFrame(frame_editor, text=" ✏️ Edit Baris Terpilih ")
        panel_edit.pack(fill="x", padx=5, pady=5)

        row_type = ttk.Frame(panel_edit)
        row_type.pack(fill="x", padx=5, pady=3)
        ttk.Label(row_type, text="Tipe:", width=8).pack(side="left")
        combo_type = ttk.Combobox(row_type, values=list(TIPE_BARIS.keys()), state="readonly", width=25)
        combo_type.pack(side="left", padx=5)
        ttk.Label(row_type, text="Align:", width=6).pack(side="left", padx=(15, 0))
        combo_align = ttk.Combobox(row_type, values=list(ALIGN_OPTIONS.keys()), state="readonly", width=20)
        combo_align.pack(side="left", padx=5)
        bold_var = tk.BooleanVar(value=False)
        chk_bold = ttk.Checkbutton(row_type, text="Bold", variable=bold_var)
        chk_bold.pack(side="left", padx=10)

        row_content = ttk.Frame(panel_edit)
        row_content.pack(fill="x", padx=5, pady=3)
        ttk.Label(row_content, text="Isi:", width=8).pack(side="left")
        ent_content = ttk.Entry(row_content, font=("Segoe UI", 10))
        ent_content.pack(side="left", fill="x", expand=True, padx=5)

        # === PANEL PREVIEW (KANAN) ===
        frame_preview = ttk.LabelFrame(main_split, text=" 👁️ Live Preview Nota ")
        frame_preview.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        txt_preview = tk.Text(frame_preview, width=48, height=35,
                              font=("Courier New", 9), bg="#1e1e2e", fg="#ecf0f1",
                              insertbackground="white")
        txt_preview.pack(side="top", fill="both", expand=True, padx=5, pady=5)

        # =================================================================
        # FUNGSI-2 EDITOR (nested closure)
        # =================================================================
        def _tpl_refresh_tree():
            for item in tree_tpl.get_children():
                tree_tpl.delete(item)
            for idx, row in enumerate(working_template["rows"], 1):
                tipe_display = TIPE_BARIS.get(row.get("type", "text"), row.get("type", "?"))
                content_display = row.get("content", "")
                if not content_display and row.get("type") in ("header_trx", "products", "total", "payment", "footer_text"):
                    content_display = f"[{tipe_display}]"
                align_display = row.get("align", "center").upper()
                tree_tpl.insert("", "end", values=(idx, tipe_display, content_display, align_display))

        def _tpl_get_selected_index():
            sel = tree_tpl.selection()
            if not sel:
                return None
            try:
                return tree_tpl.index(sel[0])
            except Exception:
                return None

        def _tpl_on_select(event=None):
            idx = _tpl_get_selected_index()
            if idx is None or idx < 0 or idx >= len(working_template["rows"]):
                return
            row = working_template["rows"][idx]
            combo_type.set(row.get("type", "text"))
            combo_align.set(row.get("align", "center"))
            ent_content.delete(0, tk.END)
            ent_content.insert(0, row.get("content", ""))
            bold_var.set(bool(row.get("bold", False)))
            if row.get("type") in ("line", "empty", "header_trx", "products", "total", "payment"):
                ent_content.config(state="disabled")
            else:
                ent_content.config(state="normal")

        def _tpl_apply_edit(event=None):
            idx = _tpl_get_selected_index()
            if idx is None or idx < 0 or idx >= len(working_template["rows"]):
                return
            new_type = combo_type.get()
            new_align = combo_align.get()
            new_content = ent_content.get()
            is_bold = bool(bold_var.get())
            working_template["rows"][idx]["type"] = new_type
            working_template["rows"][idx]["align"] = new_align
            working_template["rows"][idx]["content"] = new_content
            working_template["rows"][idx]["bold"] = is_bold
            _tpl_refresh_tree()
            _tpl_update_preview()
            children = tree_tpl.get_children()
            if idx < len(children):
                tree_tpl.selection_set(children[idx])
                tree_tpl.focus(children[idx])

        def _tpl_update_preview():
            lebar = int(self.config_nota.get("lebar_kertas", 48))
            dummy_no_trx = "1025"
            dummy_kasir = self.config_nota.get("nama_kasir", "Admin")
            dummy_waktu = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            dummy_items = [
                {"barcode": "1123LB0001", "judul": "Rosario Kayu Wangi Eksklusif",
                 "jumlah": 2, "harga_akhir": 25000, "total": 50000, "diskon": 0},
                {"barcode": "1124LB0002", "judul": "Buku Doa Katolik Saku",
                 "jumlah": 1, "harga_akhir": 15000, "total": 15000, "diskon": 0},
            ]
            dummy_grand = 65000
            dummy_bayar = 70000
            dummy_kembali = 5000
            dummy_metode = "TUNAI"

            def _fmt(text, align, width):
                clean = str(text).replace("\n", "").replace("\r", "")
                if align == "center":
                    return clean.center(width)
                elif align == "right":
                    return clean.rjust(width)
                elif align == "full":
                    if len(clean) == 1:
                        return clean * width
                    return clean.ljust(width)[:width]
                else:
                    return clean.ljust(width)[:width]

            preview_lines = []
            for row in working_template.get("rows", []):
                rtype = row.get("type", "text")
                align = row.get("align", "center")
                content = row.get("content", "")

                if rtype == "empty":
                    preview_lines.append("")
                elif rtype == "text":
                    preview_lines.append(_fmt(content, align, lebar))
                elif rtype == "line":
                    char = content if content else "-"
                    preview_lines.append(_fmt(char, "full", lebar))
                elif rtype == "header_trx":
                    preview_lines.append(f"No. Trx  : #{dummy_no_trx}".ljust(lebar // 2) +
                                         f"Kasir : {dummy_kasir}".rjust(lebar - (lebar // 2)))
                    preview_lines.append(f"Waktu   : {dummy_waktu}".ljust(lebar))
                elif rtype == "products":
                    for item in dummy_items:
                        bc = str(item.get('barcode', '')).strip()
                        if bc.endswith('.0') and bc[:-2].isdigit():
                            bc = bc[:-2]
                        judul = str(item.get('judul', ''))
                        baris = f"{bc} {judul}".strip() if bc else judul
                        preview_lines.append(baris[:lebar])
                        qty_price = f"  {item['jumlah']} x @{item['harga_akhir']:,.0f}"
                        subtotal = f"Rp. {item['total']:,.0f}"
                        spasi = max(2, lebar - len(qty_price) - len(subtotal))
                        preview_lines.append(qty_price + (" " * spasi) + subtotal)
                        if item.get('diskon', 0) > 0:
                            preview_lines.append(f"  (Diskon {item['diskon']}%)".ljust(lebar))
                elif rtype == "total":
                    preview_lines.append(f"TOTAL BELANJA".ljust(lebar // 2) +
                                         f"Rp. {dummy_grand:,.0f}".rjust(lebar - (lebar // 2)))
                elif rtype == "payment":
                    preview_lines.append(f"PEMBAYARAN ({dummy_metode})".ljust(lebar // 2) +
                                         f"Rp. {dummy_bayar:,.0f}".rjust(lebar - (lebar // 2)))
                    preview_lines.append(f"KEMBALIAN".ljust(lebar // 2) +
                                         f"Rp. {dummy_kembali:,.0f}".rjust(lebar - (lebar // 2)))
                elif rtype == "footer_text":
                    footer1 = self.config_nota.get("info_tambahan", "")
                    footer2 = self.config_nota.get("pesan_penutup", "")
                    if footer1:
                        preview_lines.append(_fmt(footer1, align, lebar))
                    if footer2:
                        preview_lines.append(_fmt(footer2, align, lebar))

            txt_preview.config(state="normal")
            txt_preview.delete("1.0", tk.END)
            txt_preview.insert("1.0", "\n".join(preview_lines))
            txt_preview.config(state="disabled")

        # =================================================================
        # HANDLER TOMBOL
        # =================================================================
        def on_tambah():
            idx = _tpl_get_selected_index()
            new_row = {"type": "text", "content": "Teks Baru", "align": "center", "bold": False}
            if idx is not None and 0 <= idx < len(working_template["rows"]):
                working_template["rows"].insert(idx + 1, new_row)
            else:
                working_template["rows"].append(new_row)
            _tpl_refresh_tree()
            _tpl_update_preview()
            children = tree_tpl.get_children()
            new_idx = (idx + 1) if idx is not None else len(children) - 1
            if 0 <= new_idx < len(children):
                tree_tpl.selection_set(children[new_idx])
                tree_tpl.focus(children[new_idx])
                _tpl_on_select()

        def on_hapus():
            idx = _tpl_get_selected_index()
            if idx is None:
                messagebox.showinfo("Info", "Pilih baris yang ingin dihapus terlebih dahulu.", parent=popup)
                return
            if messagebox.askyesno("Konfirmasi", f"Hapus baris ke-{idx + 1}?", parent=popup):
                working_template["rows"].pop(idx)
                _tpl_refresh_tree()
                _tpl_update_preview()
                children = tree_tpl.get_children()
                if children:
                    sel_idx = min(idx, len(children) - 1)
                    tree_tpl.selection_set(children[sel_idx])
                    tree_tpl.focus(children[sel_idx])
                    _tpl_on_select()

        def on_naik():
            idx = _tpl_get_selected_index()
            if idx is None or idx <= 0:
                return
            working_template["rows"][idx], working_template["rows"][idx - 1] = \
                working_template["rows"][idx - 1], working_template["rows"][idx]
            _tpl_refresh_tree()
            _tpl_update_preview()
            children = tree_tpl.get_children()
            if idx - 1 < len(children):
                tree_tpl.selection_set(children[idx - 1])
                tree_tpl.focus(children[idx - 1])
                _tpl_on_select()

        def on_turun():
            idx = _tpl_get_selected_index()
            if idx is None or idx >= len(working_template["rows"]) - 1:
                return
            working_template["rows"][idx], working_template["rows"][idx + 1] = \
                working_template["rows"][idx + 1], working_template["rows"][idx]
            _tpl_refresh_tree()
            _tpl_update_preview()
            children = tree_tpl.get_children()
            if idx + 1 < len(children):
                tree_tpl.selection_set(children[idx + 1])
                tree_tpl.focus(children[idx + 1])
                _tpl_on_select()

        def on_duplikat():
            idx = _tpl_get_selected_index()
            if idx is None:
                messagebox.showinfo("Info", "Pilih baris yang ingin diduplikat.", parent=popup)
                return
            dup = copy.deepcopy(working_template["rows"][idx])
            working_template["rows"].insert(idx + 1, dup)
            _tpl_refresh_tree()
            _tpl_update_preview()
            children = tree_tpl.get_children()
            if idx + 1 < len(children):
                tree_tpl.selection_set(children[idx + 1])
                tree_tpl.focus(children[idx + 1])
                _tpl_on_select()

        def on_reset():
            if messagebox.askyesno("Konfirmasi",
                "Reset template ke default?\nSemua perubahan akan hilang.",
                parent=popup):
                working_template["rows"] = [
                    {"type": "empty", "content": "", "align": "center"},
                    {"type": "text", "content": "TOKO KASIR TAKOM", "align": "center", "bold": True},
                    {"type": "text", "content": "Pusat Grosir & Eceran Terlengkap", "align": "center", "bold": False},
                    {"type": "text", "content": "Jl. Raya Toko No. 88, Kota Anda", "align": "center", "bold": False},
                    {"type": "text", "content": "Telp: 0812-3456-7890", "align": "center", "bold": False},
                    {"type": "line", "content": "-", "align": "full"},
                    {"type": "header_trx", "content": "", "align": "left"},
                    {"type": "line", "content": "-", "align": "full"},
                    {"type": "products", "content": "", "align": "left"},
                    {"type": "line", "content": "-", "align": "full"},
                    {"type": "total", "content": "", "align": "right"},
                    {"type": "payment", "content": "", "align": "right"},
                    {"type": "line", "content": "-", "align": "full"},
                    {"type": "text", "content": "Terima Kasih Atas Kunjungan Anda", "align": "center", "bold": False},
                    {"type": "text", "content": "Barang yang sudah dibeli tidak dapat ditukar/dikembalikan.", "align": "center", "bold": False},
                ]
                _tpl_refresh_tree()
                _tpl_update_preview()

        def on_setting():
            mini = tk.Toplevel(popup)
            mini.title("️ Pengaturan Dasar Nota")
            self.center_popup(mini, 450, 380)
            mini.transient(popup)
            mini.grab_set()

            frm = ttk.Frame(mini, padding=10)
            frm.pack(fill="both", expand=True)

            ttk.Label(frm, text="Lebar Kertas (kolom):", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
            spin_lebar = ttk.Spinbox(frm, from_=32, to=80, width=10)
            spin_lebar.set(self.config_nota.get("lebar_kertas", 48))
            spin_lebar.pack(fill="x", pady=(0, 8))

            ttk.Label(frm, text="Nama Kasir Default:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
            ent_kasir = ttk.Entry(frm, width=40)
            ent_kasir.insert(0, self.config_nota.get("nama_kasir", "Admin"))
            ent_kasir.pack(fill="x", pady=(0, 8))

            ttk.Label(frm, text="Footer Line 1 (Terima Kasih):", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
            ent_f1 = ttk.Entry(frm, width=40)
            ent_f1.insert(0, self.config_nota.get("info_tambahan", ""))
            ent_f1.pack(fill="x", pady=(0, 8))

            ttk.Label(frm, text="Footer Line 2 (Syarat & Ketentuan):", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5, 2))
            ent_f2 = ttk.Entry(frm, width=40)
            ent_f2.insert(0, self.config_nota.get("pesan_penutup", ""))
            ent_f2.pack(fill="x", pady=(0, 8))

            def simpan_setting():
                try:
                    self.config_nota["lebar_kertas"] = int(spin_lebar.get())
                except ValueError:
                    self.config_nota["lebar_kertas"] = 48
                self.config_nota["nama_kasir"] = ent_kasir.get().strip() or "Admin"
                self.config_nota["info_tambahan"] = ent_f1.get().strip()
                self.config_nota["pesan_penutup"] = ent_f2.get().strip()
                self.save_config_nota_json()
                _tpl_update_preview()
                messagebox.showinfo("Sukses", "Pengaturan dasar berhasil disimpan!", parent=mini)
                mini.destroy()

            btn_frame = ttk.Frame(frm)
            btn_frame.pack(fill="x", pady=10)
            ttk.Button(btn_frame, text="Batal", command=mini.destroy).pack(side="left", padx=5)
            ttk.Button(btn_frame, text=" Simpan", command=simpan_setting).pack(side="right", padx=5)

        btn_tambah.configure(command=on_tambah)
        btn_hapus.configure(command=on_hapus)
        btn_naik.configure(command=on_naik)
        btn_turun.configure(command=on_turun)
        btn_duplikat.configure(command=on_duplikat)
        btn_reset.configure(command=on_reset)
        btn_setting.configure(command=on_setting)

        tree_tpl.bind("<<TreeviewSelect>>", _tpl_on_select)
        combo_type.bind("<<ComboboxSelected>>", lambda e: _tpl_apply_edit())
        combo_align.bind("<<ComboboxSelected>>", lambda e: _tpl_apply_edit())
        ent_content.bind("<Return>", lambda e: _tpl_apply_edit())
        ent_content.bind("<FocusOut>", lambda e: _tpl_apply_edit())
        chk_bold.configure(command=lambda: _tpl_apply_edit())

        _tpl_refresh_tree()
        _tpl_update_preview()

        children = tree_tpl.get_children()
        if children:
            tree_tpl.selection_set(children[0])
            tree_tpl.focus(children[0])
            _tpl_on_select()

        # =================================================================
        # TOMBOL SIMPAN & BATAL DI BAGIAN BAWAH (PASTI TERLIHAT)
        # =================================================================
        frame_aksi = tk.Frame(popup, bg="#f5f6fa", height=50)
        frame_aksi.pack(side="bottom", fill="x", padx=10, pady=10)
        frame_aksi.pack_propagate(False)

        def simpan_template():
            self.template_nota = copy.deepcopy(working_template)
            if self.save_template_nota():
                messagebox.showinfo("Sukses",
                    "Template nota berhasil disimpan!\n"
                    "Perubahan akan diterapkan pada nota berikutnya.",
                    parent=popup)
                popup.destroy()
            else:
                messagebox.showerror("Gagal", "Gagal menyimpan template nota.", parent=popup)

        def batal_template():
            if messagebox.askyesno("Konfirmasi", "Batalkan perubahan template?", parent=popup):
                popup.destroy()

        # Gunakan tk.Button dengan warna eksplisit agar teks selalu terbaca
        btn_batal = tk.Button(frame_aksi, text="❌ Batal", command=batal_template,
                              font=("Segoe UI", 10, "bold"),
                              bg="#e74c3c", fg="white",
                              activebackground="#c0392b", activeforeground="white",
                              relief="flat", padx=20, pady=6, cursor="hand2")
        btn_batal.pack(side="left", padx=5, pady=8)

        btn_simpan = tk.Button(frame_aksi, text=" Simpan Template", command=simpan_template,
                               font=("Segoe UI", 10, "bold"),
                               bg="#27ae60", fg="white",
                               activebackground="#1e8449", activeforeground="white",
                               relief="flat", padx=20, pady=6, cursor="hand2")
        btn_simpan.pack(side="right", padx=5, pady=8)

    # =========================================================================
    # UBAH HARGA PRODUK (Partial Search)
    # =========================================================================
    def buka_popup_ubah_harga(self):
        if not getattr(self, 'is_supabase_ready', False):
            messagebox.showerror("Error Koneksi",
                "Klien Supabase belum terkonfigurasi atau library belum diinstal.\n"
                "1. Pastikan 'pip install supabase' sudah dijalankan.\n"
                "2. Isi SUPABASE_URL dan SUPABASE_KEY di bagian atas kode.")
            return

        popup = tk.Toplevel(self.root)
        popup.title("Ubah Harga Produk (Live Sync)")
        self.center_popup(popup, 480, 360)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: popup.destroy())

        frame_cari = ttk.LabelFrame(popup, text="1. Cari Produk (Partial Match)")
        frame_cari.pack(fill="x", padx=15, pady=10)
        row_cari = ttk.Frame(frame_cari)
        row_cari.pack(fill="x", padx=10, pady=8)
        ttk.Label(row_cari, text="Ketik Barcode/Judul:", font=("Segoe UI", 9)).pack(side="left", padx=5)
        ent_barcode = ttk.Entry(row_cari, width=22, font=("Segoe UI", 10, "bold"))
        ent_barcode.pack(side="left", padx=5)
        ent_barcode.focus_force()

        frame_info = ttk.LabelFrame(popup, text="2. Informasi Produk")
        frame_info.pack(fill="x", padx=15, pady=5)
        lbl_nama = ttk.Label(frame_info, text="Nama Produk: -", font=("Segoe UI", 9), foreground="gray")
        lbl_nama.pack(anchor="w", padx=10, pady=(6, 2))
        lbl_harga_lama = ttk.Label(frame_info, text="Harga Saat Ini: Rp. 0", font=("Segoe UI", 10, "bold"), foreground="blue")
        lbl_harga_lama.pack(anchor="w", padx=10, pady=(2, 6))

        frame_baru = ttk.LabelFrame(popup, text="3. Harga Baru")
        frame_baru.pack(fill="x", padx=15, pady=5)
        row_baru = ttk.Frame(frame_baru)
        row_baru.pack(fill="x", padx=10, pady=8)
        ttk.Label(row_baru, text="Rp.", font=("Segoe UI", 10, "bold")).pack(side="left", padx=5)
        ent_harga_baru = ttk.Entry(row_baru, width=18, font=("Segoe UI", 11, "bold"))
        ent_harga_baru.pack(side="left", padx=5)

        produk_ditemukan = {}

        def cari_produk(event=None):
            keyword = ent_barcode.get().strip()
            if not keyword:
                return
            keyword_lower = keyword.lower()
            match = self.df_master[
                self.df_master['barcode'].astype(str).str.lower().str.contains(keyword_lower, na=False) |
                self.df_master['judul'].astype(str).str.lower().str.contains(keyword_lower, na=False)
            ]
            if match.empty:
                produk_ditemukan.clear()
                lbl_nama.config(text="Nama Produk: ❌ Tidak Ditemukan", foreground="red")
                lbl_harga_lama.config(text="Harga Saat Ini: Rp. 0", foreground="gray")
                ent_harga_baru.delete(0, tk.END)
                messagebox.showwarning("Tidak Ditemukan",
                    f"Tidak ada produk dengan keyword '{keyword}' di database master.",
                    parent=popup)
            elif len(match) == 1:
                item = match.iloc[0]
                produk_ditemukan['barcode'] = str(item['barcode'])
                produk_ditemukan['judul'] = str(item['judul'])
                produk_ditemukan['harga'] = float(item['harga'])
                lbl_nama.config(text=f"Nama Produk: {produk_ditemukan['judul']}", foreground="black")
                lbl_harga_lama.config(text=f"Harga Saat Ini: Rp. {produk_ditemukan['harga']:,.0f}")
                ent_harga_baru.delete(0, tk.END)
                ent_harga_baru.insert(0, str(int(produk_ditemukan['harga'])))
                ent_harga_baru.focus_force()
                ent_harga_baru.selection_range(0, tk.END)
            else:
                self.buka_popup_pilihan_produk_ubah_harga(match, popup, produk_ditemukan,
                                                          lbl_nama, lbl_harga_lama, ent_harga_baru)

        btn_cari = ttk.Button(frame_cari, text=" Cari", command=cari_produk)
        btn_cari.pack(side="left", padx=10, pady=5)
        ent_barcode.bind("<Return>", cari_produk)

        def simpan_ke_supabase():
            if not produk_ditemukan:
                messagebox.showwarning("Peringatan", "Cari produk yang valid terlebih dahulu!", parent=popup)
                return
            try:
                harga_baru = float(ent_harga_baru.get().replace(",", "").strip())
                if harga_baru < 0:
                    raise ValueError("Harga tidak boleh negatif")
            except ValueError:
                messagebox.showerror("Input Error", "Masukkan nominal harga yang valid (angka)!", parent=popup)
                return
            popup.config(cursor="watch")
            popup.update()
            try:
                debug_select = self.supabase.table(SUPABASE_TABLE).select("barcode, harga").eq("barcode", produk_ditemukan['barcode']).execute()
                if not debug_select.data:
                    raise Exception(f"Tidak ditemukan data dengan barcode='{produk_ditemukan['barcode']}'.")
                response = self.supabase.table(SUPABASE_TABLE).update({
                    "harga": harga_baru
                }).eq("barcode", produk_ditemukan['barcode']).execute()
                idx = self.df_master.index[self.df_master['barcode'].astype(str) == produk_ditemukan['barcode']].tolist()
                if idx:
                    self.df_master.at[idx[0], 'harga'] = harga_baru
                messagebox.showinfo("Sukses",
                    f"Harga '{produk_ditemukan['judul']}' berhasil diubah menjadi Rp. {harga_baru:,.0f}\n"
                    "Data telah tersinkronisasi ke Database & Cache Lokal.", parent=popup)
                popup.config(cursor="")
                popup.destroy()
                self.ent_search.focus_force()
                self.refresh_data_master_view()
            except Exception as e:
                popup.config(cursor="")
                messagebox.showerror("Gagal Sync", f"Gagal mengupdate ke Supabase:\n{str(e)}", parent=popup)

        frame_aksi = ttk.Frame(popup)
        frame_aksi.pack(fill="x", padx=15, pady=15)
        ttk.Button(frame_aksi, text="Batal", command=popup.destroy).pack(side="left", padx=5)
        ttk.Button(frame_aksi, text=" Simpan & Sync ke Supabase", command=simpan_ke_supabase).pack(side="right", padx=5)

    def buka_popup_pilihan_produk_ubah_harga(self, df_pilihan, parent_popup, produk_ditemukan,
                                              lbl_nama, lbl_harga_lama, ent_harga_baru):
        popup_pilih = tk.Toplevel(self.root)
        popup_pilih.title(f"Pilih Produk ({len(df_pilihan)} hasil)")
        self.center_popup(popup_pilih, 580, 380)
        popup_pilih.transient(self.root)
        popup_pilih.grab_set()
        popup_pilih.bind("<Escape>", lambda e: popup_pilih.destroy())

        ttk.Label(popup_pilih, text=f"Ditemukan {len(df_pilihan)} produk. Pilih salah satu:",
                  font=("Segoe UI", 9, "bold")).pack(pady=8)

        frame_list = ttk.Frame(popup_pilih)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("barcode", "judul", "harga")
        tree_select = ttk.Treeview(frame_list, columns=cols, show="headings", selectmode="browse",
                                   style="Unified.Treeview")
        tree_select.heading("barcode", text="Barcode")
        tree_select.heading("judul", text="Judul")
        tree_select.heading("harga", text="Harga")
        tree_select.column("barcode", width=120)
        tree_select.column("judul", width=310)
        tree_select.column("harga", width=100, anchor="e")
        sb = ttk.Scrollbar(frame_list, orient="vertical", command=tree_select.yview)
        tree_select.configure(yscroll=sb.set)
        sb.pack(side="right", fill="y")
        tree_select.pack(side="left", fill="both", expand=True)

        for _, row in df_pilihan.iterrows():
            harga_val = float(row['harga'])
            tree_select.insert("", "end", values=(row['barcode'], row['judul'], f"Rp. {harga_val:,.0f}"))

        children = tree_select.get_children()
        if children:
            tree_select.selection_set(children[0])
            tree_select.focus_set()
            tree_select.focus(children[0])

        def pilih_item_terpilih(event=None):
            selected = tree_select.selection()
            if not selected:
                return
            idx = tree_select.index(selected[0])
            item_selected = df_pilihan.iloc[idx]
            produk_ditemukan.clear()
            produk_ditemukan['barcode'] = str(item_selected['barcode'])
            produk_ditemukan['judul'] = str(item_selected['judul'])
            produk_ditemukan['harga'] = float(item_selected['harga'])
            lbl_nama.config(text=f"Nama Produk: {produk_ditemukan['judul']}", foreground="black")
            lbl_harga_lama.config(text=f"Harga Saat Ini: Rp. {produk_ditemukan['harga']:,.0f}")
            ent_harga_baru.delete(0, tk.END)
            ent_harga_baru.insert(0, str(int(produk_ditemukan['harga'])))
            ent_harga_baru.focus_force()
            ent_harga_baru.selection_range(0, tk.END)
            popup_pilih.destroy()

        tree_select.bind("<Return>", pilih_item_terpilih)
        popup_pilih.bind("<Return>", pilih_item_terpilih)
        ttk.Button(popup_pilih, text="Pilih Produk Ini (ENTER)", command=pilih_item_terpilih).pack(pady=8)

    # =========================================================================
    # TAMBAH PRODUK
    # =========================================================================
    def buka_popup_tambah_produk(self):
        if not getattr(self, 'is_supabase_ready', False):
            messagebox.showerror("Error Koneksi",
                "Klien Supabase belum terkonfigurasi atau library belum diinstal.\n"
                "1. Pastikan 'pip install supabase' sudah dijalankan.\n"
                "2. Isi SUPABASE_URL dan SUPABASE_KEY di file .env.")
            return

        popup = tk.Toplevel(self.root)
        popup.title("Tambah Produk Baru (Live Sync)")
        self.center_popup(popup, 500, 480)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: (popup.destroy(), self.ent_search.focus_force()))

        frame_form = ttk.LabelFrame(popup, text="1. Data Produk (Urut Kolom: barcode | judul | harga)")
        frame_form.pack(fill="x", padx=15, pady=10)

        row_barcode = ttk.Frame(frame_form)
        row_barcode.pack(fill="x", padx=10, pady=(8, 2))
        ttk.Label(row_barcode, text="Barcode:", width=12, font=("Segoe UI", 10)).pack(side="left")
        ent_barcode = ttk.Entry(row_barcode, font=("Segoe UI", 11, "bold"))
        ent_barcode.pack(side="left", fill="x", expand=True)

        row_judul = ttk.Frame(frame_form)
        row_judul.pack(fill="x", padx=10, pady=2)
        ttk.Label(row_judul, text="Judul:", width=12, font=("Segoe UI", 10)).pack(side="left")
        ent_judul = ttk.Entry(row_judul, font=("Segoe UI", 11))
        ent_judul.pack(side="left", fill="x", expand=True)

        row_harga = ttk.Frame(frame_form)
        row_harga.pack(fill="x", padx=10, pady=(2, 8))
        ttk.Label(row_harga, text="Harga (Rp.):", width=12, font=("Segoe UI", 10)).pack(side="left")
        ent_harga = ttk.Entry(row_harga, font=("Segoe UI", 11, "bold"))
        ent_harga.pack(side="left", fill="x", expand=True)

        frame_preview = ttk.LabelFrame(popup, text="2. Preview Baris Tabel")
        frame_preview.pack(fill="both", expand=True, padx=15, pady=5)
        tree_preview = ttk.Treeview(frame_preview, columns=("barcode", "judul", "harga"), show="headings", height=4,
                                    style="Unified.Treeview")
        tree_preview.heading("barcode", text="barcode")
        tree_preview.heading("judul", text="judul")
        tree_preview.heading("harga", text="harga")
        tree_preview.column("barcode", width=130, anchor="w")
        tree_preview.column("judul", width=230, anchor="w")
        tree_preview.column("harga", width=90, anchor="e")
        tree_preview.pack(fill="both", expand=True, padx=5, pady=5)

        def update_preview(event=None):
            for r in tree_preview.get_children():
                tree_preview.delete(r)
            tree_preview.insert("", "end", values=(
                ent_barcode.get().strip(),
                ent_judul.get().strip(),
                ent_harga.get().strip()
            ))

        for ent in (ent_barcode, ent_judul, ent_harga):
            ent.bind("<KeyRelease>", update_preview)
        update_preview()
        ent_barcode.focus_force()

        def fokus_judul(event=None):
            ent_judul.focus_force()
            return "break"

        def fokus_harga(event=None):
            ent_harga.focus_force()
            return "break"

        def simpan_produk(event=None):
            barcode_baru = ent_barcode.get().strip()
            judul_baru = ent_judul.get().strip()
            if not barcode_baru or not judul_baru:
                messagebox.showwarning("Peringatan", "Barcode dan Judul produk wajib diisi!", parent=popup)
                return
            try:
                harga_baru = float(ent_harga.get().replace(",", "").strip())
                if harga_baru < 0:
                    raise ValueError("Harga tidak boleh negatif")
            except ValueError:
                messagebox.showerror("Input Error", "Masukkan nominal harga yang valid (angka)!", parent=popup)
                return
            if self.is_loading:
                messagebox.showwarning("Proses Loading", "Database master sedang dimuat. Mohon tunggu sebentar!", parent=popup)
                return
            try:
                if self.df_master is not None and not self.df_master.empty:
                    duplikat = self.df_master[self.df_master['barcode'].astype(str) == barcode_baru]
                    if not duplikat.empty:
                        messagebox.showwarning("Duplikat", f"Barcode '{barcode_baru}' sudah terdaftar di tabel master!", parent=popup)
                        return
                cek_cloud = self.supabase.table(SUPABASE_TABLE).select("barcode").eq("barcode", barcode_baru).execute()
                if cek_cloud.data:
                    messagebox.showwarning("Duplikat", f"Barcode '{barcode_baru}' sudah terdaftar di tabel master!", parent=popup)
                    return
            except Exception as e:
                logging.error(f"Gagal cek duplikat barcode: {e}", exc_info=True)
            popup.config(cursor="watch")
            popup.update()
            try:
                response = self.supabase.table(SUPABASE_TABLE).insert({
                    "barcode": barcode_baru,
                    "judul": judul_baru,
                    "harga": harga_baru
                }).execute()
                baris_baru = pd.DataFrame([{"barcode": barcode_baru, "judul": judul_baru, "harga": harga_baru}])
                if self.df_master is None or self.df_master.empty:
                    self.df_master = baris_baru
                else:
                    self.df_master = pd.concat([self.df_master, baris_baru], ignore_index=True)
                try:
                    self.sync_tool.df_master = self.df_master
                except Exception:
                    pass
                try:
                    self.df_master.to_parquet("master_cache.parquet", index=False)
                except Exception as e:
                    print(f"⚠️ Gagal memperbarui cache parquet: {e}")
                self.lbl_status.config(
                    text=f"[ 🟢 ONLINE Cloud: {len(self.df_master):,} Produk ]",
                    foreground="green"
                )
                popup.config(cursor="")
                popup.destroy()
                messagebox.showinfo("Sukses",
                    f"Produk '{judul_baru}' berhasil ditambahkan ke tabel '{SUPABASE_TABLE}'.\n"
                    "Data telah tersinkronisasi ke Database & Cache Lokal.")
                self.ent_search.focus_force()
                self.refresh_data_master_view()
            except Exception as e:
                popup.config(cursor="")
                messagebox.showerror("Gagal Sync", f"Gagal menambahkan produk ke Supabase:\n{str(e)}", parent=popup)

        ent_barcode.bind("<Return>", fokus_judul)
        ent_judul.bind("<Return>", fokus_harga)
        ent_harga.bind("<Return>", simpan_produk)

        frame_aksi = ttk.Frame(popup)
        frame_aksi.pack(fill="x", padx=15, pady=15)
        ttk.Button(frame_aksi, text="Batal (ESC)", command=popup.destroy).pack(side="left", padx=5)
        ttk.Button(frame_aksi, text="💾 Simpan & Sync (ENTER)", command=simpan_produk).pack(side="right", padx=5)

    # =========================================================================
    # HALAMAN TRANSAKSI - RESPONSIVE + TOMBOL TERBACA
    # =========================================================================
    def setup_halaman_transaksi(self, parent):
        top_container = ttk.Frame(parent)
        top_container.pack(fill="x", padx=10, pady=5)

        frame_config = ttk.LabelFrame(top_container, text=" Pengaturan File & Sheet Target ")
        frame_config.pack(side="left", fill="both", expand=True)

        row_master = ttk.Frame(frame_config)
        row_master.pack(fill="x", padx=5, pady=6)
        self.lbl_status = ttk.Label(row_master, text="[ 🔄 Connecting Database... Mohon Tunggu ]", font=("Segoe UI", 9, "bold"), foreground="blue")
        self.lbl_status.pack(side="left", padx=5)

        row_target = ttk.Frame(frame_config)
        row_target.pack(fill="x", padx=5, pady=6)
        btn_target = ttk.Button(row_target, text="Pilih File Target Output", command=self.pilih_target_file)
        btn_target.pack(side="left", padx=5)
        self.lbl_target_path = ttk.Label(row_target, text="Belum dipilih", foreground="red", font=("Segoe UI", 9, "bold"))
        self.lbl_target_path.pack(side="left", padx=5)
        ttk.Label(row_target, text="Pilih Sheet Target:").pack(side="left", padx=(20, 2))
        self.combo_sheet = ttk.Combobox(row_target, width=15, state="readonly")
        self.combo_sheet.pack(side="left", padx=2)
        self.combo_sheet.bind("<<ComboboxSelected>>", self.on_sheet_changed)
        btn_add_sheet = ttk.Button(row_target, text="+ Sheet", width=8, command=self.tambah_sheet_baru)
        btn_add_sheet.pack(side="left", padx=5)

        info_struktur = ttk.Label(
            frame_config,
            text=" Struktur Kolom Excel Baku (A s.d. L): [A] No | [B] Kode | [C] Judul | [D] Qty | [E] Harga | [F] Total | [G-H] Metode | [I] Diskon | [J] Member | [K] Waktu | [L] NOMAN",
            font=("Segoe UI", 8, "italic"),
            foreground="darkslategray"
        )
        info_struktur.pack(anchor="w", padx=10, pady=(4, 8))

        scan_frame = ttk.LabelFrame(parent, text=" Area Scan / Cari Produk ")
        scan_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(scan_frame, text="Cari Barcode / Judul:", font=("Segoe UI", 10, "bold")).pack(side="left", padx=5)
        self.ent_search = ttk.Entry(scan_frame, font=("Segoe UI", 11))
        self.ent_search.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        self.ent_search.bind("<Return>", self.proses_scan_trigger)
        self.ent_search.bind("<Down>", self.fokus_ke_tabel)
        self.ent_search.focus()

        table_frame = ttk.Frame(parent)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("no", "barcode", "judul", "jumlah", "harga_normal", "diskon", "harga_akhir", "total")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                 selectmode="browse", style="Unified.Treeview")
        self.tree.heading("no", text="No")
        self.tree.heading("barcode", text="Barcode")
        self.tree.heading("judul", text="Judul")
        self.tree.heading("jumlah", text="Jumlah")
        self.tree.heading("harga_normal", text="Harga Normal")
        self.tree.heading("diskon", text="Diskon")
        self.tree.heading("harga_akhir", text="Harga Akhir")
        self.tree.heading("total", text="Total")
        self.tree.column("no", width=50, anchor="center")
        self.tree.column("barcode", width=130)
        self.tree.column("judul", width=400)
        self.tree.column("jumlah", width=70, anchor="center")
        self.tree.column("harga_normal", width=110, anchor="e")
        self.tree.column("diskon", width=80, anchor="center")
        self.tree.column("harga_akhir", width=110, anchor="e")
        self.tree.column("total", width=130, anchor="e")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>", self.buka_popup_edit_qty_langsung)
        self.tree.bind("<Delete>", self.hapus_item_terpilih)
        self.tree.bind("<BackSpace>", self.hapus_item_terpilih)
        self.tree.bind("<Escape>", lambda event: self.ent_search.focus_force())
        self.root.bind("<F5>", lambda event: self.buka_popup_pembayaran())

                # ============================================================
        # BOTTOM BAR - TOMBOL SERAGAM (ukuran = tombol Editor Nota)
        # ============================================================
        bottom_frame = tk.Frame(parent, bg="#f5f6fa")
        bottom_frame.pack(side="bottom", fill="x", padx=10, pady=8)

        # --- Summary Box ---
        summary_box = tk.Frame(bottom_frame, bg="#ecf0f1", bd=1, relief="solid")
        summary_box.pack(side="top", fill="x", pady=(0, 8))

        self.lbl_total_qty = tk.Label(summary_box,
                                      text="TOTAL ITEM / QTY: 0 Pcs",
                                      font=("Segoe UI", 12, "bold"),
                                      bg="#ecf0f1", fg="#27ae60")
        self.lbl_total_qty.pack(side="left", padx=15, pady=10)

        self.lbl_total = tk.Label(summary_box,
                                  text="TOTAL BELANJA: Rp. 0",
                                  font=("Segoe UI", 16, "bold"),
                                  bg="#ecf0f1", fg="#2980b9")
        self.lbl_total.pack(side="right", padx=15, pady=10)

        # --- Button Row (UKURAN SERAGAM DENGAN EDITOR NOTA) ---
        btn_row = tk.Frame(bottom_frame, bg="#f5f6fa")
        btn_row.pack(side="bottom", fill="x", pady=2)

        # Tombol Hapus - MERAH
        btn_delete_item = tk.Button(btn_row,
                                    text="🗑️ Hapus Item (Del)",
                                    command=self.hapus_item_terpilih,
                                    font=("Segoe UI", 10, "bold"),
                                    bg="#e74c3c", fg="#ffffff",
                                    activebackground="#c0392b",
                                    activeforeground="#ffffff",
                                    relief="flat",
                                    padx=20, pady=6,
                                    cursor="hand2",
                                    borderwidth=0)
        btn_delete_item.pack(side="left", padx=5)

        # Tombol Reset - ABU-ABU GELAP
        btn_reset = tk.Button(btn_row,
                              text="🔄 Reset Transaksi",
                              command=self.reset_transaksi,
                              font=("Segoe UI", 10, "bold"),
                              bg="#7f8c8d", fg="#ffffff",
                              activebackground="#616a6b",
                              activeforeground="#ffffff",
                              relief="flat",
                              padx=20, pady=6,
                              cursor="hand2",
                              borderwidth=0)
        btn_reset.pack(side="right", padx=5)

        # Tombol Selesai - HIJAU (lebih besar & menonjol)
        btn_finish = tk.Button(btn_row,
                               text="💰 SELESAI & SIMPAN (F5)",
                               command=self.buka_popup_pembayaran,
                               font=("Segoe UI", 11, "bold"),
                               bg="#27ae60", fg="#ffffff",
                               activebackground="#1e8449",
                               activeforeground="#ffffff",
                               relief="flat",
                               padx=25, pady=6,
                               cursor="hand2",
                               borderwidth=0)
        btn_finish.pack(side="right", padx=5)

    # =========================================================================
    # HALAMAN LAPORAN
    # =========================================================================
    def setup_halaman_laporan(self, parent):
        frame_top_rep = ttk.Frame(parent)
        frame_top_rep.pack(fill="x", padx=10, pady=10)
        lbl_rep_title = ttk.Label(frame_top_rep, text="Pilih Sheet Laporan:", font=("Segoe UI", 10, "bold"))
        lbl_rep_title.pack(side="left", padx=5)
        self.lbl_file_output = ttk.Label(frame_top_rep, text="[File: -]", font=("Segoe UI", 9, "italic"), foreground="gray")
        self.lbl_file_output.pack(side="left", padx=10)
        self.combo_sheet_report = ttk.Combobox(frame_top_rep, width=20, state="readonly")
        self.combo_sheet_report.pack(side="left", padx=5)
        self.combo_sheet_report.bind("<<ComboboxSelected>>", self.muat_tabel_laporan_excel)
        btn_refresh = ttk.Button(frame_top_rep, text="🔄 Refresh Data", command=self.force_refresh_laporan)
        btn_refresh.pack(side="left", padx=5)
        btn_setoran = ttk.Button(frame_top_rep, text="📊 Setoran Harian", command=self.proses_setoran_harian)
        btn_setoran.pack(side="left", padx=15)
        btn_reprint = ttk.Button(frame_top_rep, text="🖨️ Re-Print Nota Terpilih", command=self.reprint_nota_dari_sheet)
        btn_reprint.pack(side="left", padx=5)
        btn_save_noman = ttk.Button(frame_top_rep, text="💾 Simpan Semua NOMAN", command=self.simpan_semua_noman_ke_excel)
        btn_save_noman.pack(side="left", padx=5)

        self.frame_summary_box = tk.Frame(parent, bg="#ecf0f1", bd=1, relief="solid")
        self.frame_summary_box.pack(fill="x", padx=10, pady=5)

        self.lbl_setoran_tunai = tk.Label(self.frame_summary_box, text="Total Tunai: Rp. 0",
                                          font=("Segoe UI", 11, "bold"),
                                          bg="#ecf0f1", fg="#27ae60")
        self.lbl_setoran_tunai.pack(side="left", padx=15, pady=10)

        self.lbl_setoran_nontunai = tk.Label(self.frame_summary_box, text="Total Non-Tunai: Rp. 0",
                                             font=("Segoe UI", 11, "bold"),
                                             bg="#ecf0f1", fg="#e67e22")
        self.lbl_setoran_nontunai.pack(side="left", padx=15, pady=10)

        self.lbl_setoran_grand = tk.Label(self.frame_summary_box, text="GRAND TOTAL SETORAN: Rp. 0",
                                          font=("Segoe UI", 12, "bold"),
                                          bg="#ecf0f1", fg="#2980b9")
        self.lbl_setoran_grand.pack(side="right", padx=15, pady=10)

        table_frame_rep = ttk.Frame(parent)
        table_frame_rep.pack(fill="both", expand=True, padx=10, pady=5)
        self.report_tree = ttk.Treeview(table_frame_rep, show="headings",
                                        style="Unified.Treeview", selectmode="browse")
        sb_y = ttk.Scrollbar(table_frame_rep, orient="vertical", command=self.report_tree.yview)
        sb_x = ttk.Scrollbar(table_frame_rep, orient="horizontal", command=self.report_tree.xview)
        self.report_tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.report_tree.pack(side="left", fill="both", expand=True)
        self.report_tree.bind("<Button-1>", self.on_report_tree_click)
        self.report_tree.bind("<Motion>", self.on_report_tree_motion)

    # =========================================================================
    # FUNGSI LAPORAN (NOMAN, SETORAN, REPRINT)
    # =========================================================================
    def on_report_tree_click(self, event):
        region = self.report_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.report_tree.identify_column(event.x)
        if col != '#12':
            return
        item_id = self.report_tree.identify_row(event.y)
        if not item_id:
            return
        values = self.report_tree.item(item_id, "values")
        if len(values) < 1:
            return
        no_transaksi = str(values[0]).strip()
        if not self.is_first_row_of_transaction(item_id, no_transaksi):
            return
        self.open_noman_input_dialog(item_id, no_transaksi)

    def is_first_row_of_transaction(self, item_id, no_transaksi):
        children = list(self.report_tree.get_children())
        try:
            idx = children.index(item_id)
            if idx > 0:
                prev_id = children[idx - 1]
                prev_values = self.report_tree.item(prev_id, "values")
                if len(prev_values) > 0 and str(prev_values[0]).strip() == no_transaksi:
                    return False
            return True
        except:
            return True

    def on_report_tree_motion(self, event):
        try:
            region = self.report_tree.identify("region", event.x, event.y)
            if region == "cell":
                col = self.report_tree.identify_column(event.x)
                if col == '#12':
                    item_id = self.report_tree.identify_row(event.y)
                    if item_id:
                        values = self.report_tree.item(item_id, "values")
                        if len(values) > 0:
                            no_transaksi = str(values[0]).strip()
                            if self.is_first_row_of_transaction(item_id, no_transaksi):
                                self.report_tree.configure(cursor="hand2")
                                return
            self.report_tree.configure(cursor="")
        except Exception:
            pass

    def open_noman_input_dialog(self, item_id, no_transaksi):
        sheet_name = self.combo_sheet_report.get()
        noman_key = (sheet_name, no_transaksi)
        current_state = self.noman_data.get(noman_key, {"checked": False, "value": None})

        popup = tk.Toplevel(self.root)
        popup.title(f"Input NOMAN - Transaksi #{no_transaksi}")
        self.center_popup(popup, 380, 220)
        popup.transient(self.root)
        popup.grab_set()
        popup.configure(bg="#f0f0f0")

        ttk.Label(popup, text=f"📝 NOMAN untuk Transaksi #{no_transaksi}", font=("Segoe UI", 11, "bold")).pack(pady=(15, 5))
        ttk.Label(popup, text="Masukkan nilai numerik (boleh diawali 0):", font=("Segoe UI", 9)).pack(pady=(0, 5))
        ent_noman = ttk.Entry(popup, width=25, font=("Segoe UI", 14, "bold"), justify="center")
        ent_noman.pack(pady=5)
        if current_state["value"] is not None:
            ent_noman.insert(0, str(current_state["value"]))
        self.noman_entry_widget = ent_noman
        popup.after(100, lambda: self._force_focus_entry(ent_noman))

        def validate_input(P):
            if P == "":
                return True
            if all(c.isdigit() or c == '.' for c in P):
                if P.count('.') <= 1:
                    return True
            return False

        vcmd = (popup.register(validate_input), '%P')
        ent_noman.configure(validate="key", validatecommand=vcmd)

        btn_frame = ttk.Frame(popup)
        btn_frame.pack(pady=10)

        def simpan_nilai(event=None):
            raw = ent_noman.get().strip()
            if not raw or raw == '.':
                messagebox.showwarning("Peringatan", "Nilai NOMAN tidak boleh kosong!", parent=popup)
                popup.after(100, lambda: self._force_focus_entry(ent_noman))
                return
            nilai_str = raw
            self.noman_data[noman_key] = {"checked": True, "value": nilai_str}
            self.update_noman_display_for_transaction(no_transaksi, nilai_str)
            if self.target_file_path:
                try:
                    self.wb_target = openpyxl.load_workbook(self.target_file_path)
                    ws = self.wb_target[sheet_name]
                    first_row_updated = False
                    for r_idx in range(2, ws.max_row + 1):
                        cell_a = ws[f"A{r_idx}"].value
                        if cell_a is not None and str(cell_a).strip() == no_transaksi:
                            if not first_row_updated:
                                cell = ws[f"L{r_idx}"]
                                cell.value = nilai_str
                                cell.number_format = '@'
                                first_row_updated = True
                            else:
                                ws[f"L{r_idx}"] = ""
                    self.wb_target.save(self.target_file_path)
                    self.cached_report_df = None
                    self.cached_sheet_name = None
                except Exception as e:
                    logging.error(f"Gagal simpan NOMAN ke Excel: {e}", exc_info=True)
                    messagebox.showwarning("Peringatan", f"Nilai tersimpan di tampilan tapi gagal ditulis ke Excel:\n{str(e)}", parent=popup)
            popup.destroy()

        def batal_input():
            if noman_key in self.noman_data:
                del self.noman_data[noman_key]
            self.update_noman_display_for_transaction(no_transaksi, "")
            if self.target_file_path:
                try:
                    self.wb_target = openpyxl.load_workbook(self.target_file_path)
                    ws = self.wb_target[sheet_name]
                    for r_idx in range(2, ws.max_row + 1):
                        cell_a = ws[f"A{r_idx}"].value
                        if cell_a is not None and str(cell_a).strip() == no_transaksi:
                            ws[f"L{r_idx}"] = ""
                    self.wb_target.save(self.target_file_path)
                    self.cached_report_df = None
                    self.cached_sheet_name = None
                except Exception as e:
                    logging.error(f"Gagal reset NOMAN di Excel: {e}", exc_info=True)
            popup.destroy()

        ttk.Button(btn_frame, text=" Simpan", command=simpan_nilai).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="❌ Batal", command=batal_input).pack(side="left", padx=5)
        popup.bind("<Return>", simpan_nilai)
        popup.bind("<Escape>", lambda e: popup.destroy())

    def _force_focus_entry(self, entry_widget):
        try:
            entry_widget.focus_set()
            entry_widget.icursor(tk.END)
            entry_widget.selection_range(0, tk.END)
        except Exception as e:
            logging.error(f"Gagal focus entry: {e}", exc_info=True)

    def update_noman_display_for_transaction(self, no_transaksi, nilai_str):
        children = list(self.report_tree.get_children())
        first_row_found = False
        for child_id in children:
            values = list(self.report_tree.item(child_id, "values"))
            if len(values) > 0 and str(values[0]).strip() == no_transaksi:
                while len(values) < 12:
                    values.append("")
                if not first_row_found:
                    values[11] = nilai_str
                    first_row_found = True
                else:
                    values[11] = ""
                self.report_tree.item(child_id, values=values)

    def simpan_semua_noman_ke_excel(self):
        if not self.target_file_path:
            messagebox.showwarning("Peringatan", "File target belum dipilih!")
            return
        sheet_name = self.combo_sheet_report.get()
        if not sheet_name:
            messagebox.showwarning("Peringatan", "Pilih sheet laporan yang valid!")
            return
        sheet_noman = {k: v for k, v in self.noman_data.items() if k[0] == sheet_name and v.get("checked") and v.get("value") is not None}
        if not sheet_noman:
            messagebox.showinfo("Info", "Tidak ada data NOMAN yang perlu disimpan untuk sheet ini.")
            return
        try:
            self.wb_target = openpyxl.load_workbook(self.target_file_path)
            if sheet_name not in self.wb_target.sheetnames:
                messagebox.showwarning("Peringatan", "Sheet tidak ditemukan di file Excel!")
                return
            ws = self.wb_target[sheet_name]
            saved_count = 0
            for (s_name, no_transaksi), state in sheet_noman.items():
                try:
                    first_row_updated = False
                    for r_idx in range(2, ws.max_row + 1):
                        cell_a = ws[f"A{r_idx}"].value
                        if cell_a is not None and str(cell_a).strip() == no_transaksi:
                            if not first_row_updated:
                                cell = ws[f"L{r_idx}"]
                                cell.value = str(state["value"])
                                cell.number_format = '@'
                                first_row_updated = True
                                saved_count += 1
                            else:
                                ws[f"L{r_idx}"] = ""
                except Exception as e:
                    logging.error(f"Gagal simpan NOMAN transaksi {no_transaksi}: {e}", exc_info=True)
            self.wb_target.save(self.target_file_path)
            self.cached_report_df = None
            self.cached_sheet_name = None
            messagebox.showinfo("Sukses", f"Berhasil menyimpan {saved_count} nilai NOMAN ke Excel!")
            self.force_refresh_laporan()
        except Exception as e:
            self._show_safe_error(self.root, "Error", "Gagal menyimpan data NOMAN ke Excel.", exception=e, context="simpan_semua_noman_ke_excel")

    def force_refresh_laporan(self):
        self.cached_report_df = None
        self.cached_sheet_name = None
        self.refresh_data_laporan()
        messagebox.showinfo("Info", "Data laporan berhasil diperbarui dari file Excel.")

    def refresh_data_laporan(self):
        if not self.target_file_path or not self.wb_target:
            self.combo_sheet_report['values'] = []
            self.lbl_file_output.config(text="[File: -]")
            return
        file_name = os.path.basename(self.target_file_path)
        self.lbl_file_output.config(text=f"[File: {file_name}]")
        sheets = self.wb_target.sheetnames
        self.combo_sheet_report['values'] = sheets
        current_active = self.combo_sheet.get()
        if current_active in sheets:
            self.combo_sheet_report.set(current_active)
        elif sheets:
            self.combo_sheet_report.current(0)
        self.muat_tabel_laporan_excel()

    def muat_tabel_laporan_excel(self, event=None):
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)
        self.noman_item_map = {}
        if not self.target_file_path or not self.wb_target:
            return
        sheet_name = self.combo_sheet_report.get()
        if not sheet_name or sheet_name not in self.wb_target.sheetnames:
            return

        MAX_FILE_SIZE_MB = 50
        try:
            file_size_mb = os.path.getsize(self.target_file_path) / (1024 * 1024)
            if file_size_mb > MAX_FILE_SIZE_MB:
                if not messagebox.askyesno("Peringatan File Besar", f"Ukuran file Excel sangat besar ({file_size_mb:.1f} MB).\nMemuatnya dapat membuat aplikasi lambat atau crash.\n\nApakah Anda yakin ingin melanjutkan?"):
                    return
        except OSError:
            pass

        if self.cached_report_df is not None and self.cached_sheet_name == sheet_name:
            df_raw = self.cached_report_df
        else:
            try:
                df_raw = pd.read_excel(self.target_file_path, sheet_name=sheet_name, header=None, dtype=str)
                self.cached_report_df = df_raw
                self.cached_sheet_name = sheet_name
            except Exception as e:
                self._show_safe_error(self.root, "Error", "Gagal membaca file Excel. Pastikan file tidak sedang dibuka di aplikasi lain.", exception=e, context="muat_tabel_laporan_excel")
                return

        if df_raw.empty or len(df_raw) < 1:
            return

        row1 = df_raw.iloc[0].values if len(df_raw) > 0 else []
        row2 = df_raw.iloc[1].values if len(df_raw) > 1 else []
        headers = []
        for i in range(df_raw.shape[1]):
            val1 = str(row1[i]).strip() if i < len(row1) and pd.notna(row1[i]) and str(row1[i]).strip().lower() != 'nan' else ""
            val2 = str(row2[i]).strip() if i < len(row2) and pd.notna(row2[i]) and str(row2[i]).strip().lower() != 'nan' else ""
            if val1 and val2 and val1 != val2:
                header_text = f"{val1} - {val2}"
            elif val1:
                header_text = val1
            elif val2:
                header_text = val2
            else:
                header_text = f"Kolom {i+1}"
            if header_text.lower() in ["no transaksi", "no. transaksi"]:
                header_text = "No"
            headers.append(header_text)

        num_cols = max(df_raw.shape[1], 12)
        while len(headers) < 12:
            headers.append(f"Kolom {len(headers)+1}")
        headers[11] = "NOMAN"

        cols = [f"col_{i}" for i in range(num_cols)]
        self.report_tree["columns"] = cols
        for idx, col in enumerate(cols):
            h_text = headers[idx] if idx < len(headers) else col
            self.report_tree.heading(col, text=h_text)
            max_len = len(h_text)
            for r_idx in range(1, min(50, len(df_raw))):
                if idx < df_raw.shape[1]:
                    cell_val = str(df_raw.iloc[r_idx, idx]) if pd.notna(df_raw.iloc[r_idx, idx]) and str(df_raw.iloc[r_idx, idx]).strip().lower() != 'nan' else ""
                else:
                    cell_val = ""
                if len(cell_val) > max_len:
                    max_len = len(cell_val)
            if idx == 11:
                col_width = 110
            else:
                col_width = 45 if idx == 0 else max(max_len * 9, 90)
            self.report_tree.column(col, width=col_width, anchor="w", stretch=False)

        processed_transactions = set()
        for r_idx in range(2, len(df_raw)):
            excel_row = r_idx + 1
            row_vals = []
            for c_idx in range(num_cols):
                if c_idx < df_raw.shape[1]:
                    raw_val = df_raw.iloc[r_idx, c_idx]
                    if pd.notna(raw_val) and str(raw_val).strip().lower() != 'nan':
                        row_vals.append(str(raw_val).strip())
                    else:
                        row_vals.append("")
                else:
                    row_vals.append("")
            while len(row_vals) < 12:
                row_vals.append("")

            no_transaksi = row_vals[0].strip() if len(row_vals) > 0 else ""
            if no_transaksi and no_transaksi not in processed_transactions:
                noman_key = (sheet_name, no_transaksi)
                existing_l_value = row_vals[11] if len(row_vals) > 11 else ""
                if noman_key in self.noman_data and self.noman_data[noman_key].get("checked"):
                    display_val = str(self.noman_data[noman_key]["value"])
                elif existing_l_value and existing_l_value.strip():
                    display_val = existing_l_value.strip()
                    self.noman_data[noman_key] = {"checked": True, "value": display_val}
                else:
                    display_val = "☐"
                row_vals[11] = display_val
                processed_transactions.add(no_transaksi)
            else:
                row_vals[11] = ""

            item_id = self.report_tree.insert("", "end", values=row_vals[:num_cols])
            self.noman_item_map[item_id] = (sheet_name, no_transaksi, excel_row)

        self.hitung_dan_tampilkan_setoran(df_raw)

    def hitung_dan_tampilkan_setoran(self, df_raw):
        total_tunai = 0.0
        total_nontunai = 0.0
        for r_idx in range(2, len(df_raw)):
            try:
                val_a = str(df_raw.iloc[r_idx, 0]).strip().upper()
                if "SETORAN" in val_a or "TOTAL" in val_a or "TUTUP" in val_a:
                    continue
                val_g = df_raw.iloc[r_idx, 6] if df_raw.shape[1] > 6 else 0
                if pd.notna(val_g) and str(val_g).strip().lower() != 'nan':
                    total_tunai += float(str(val_g).replace(",", ""))
            except Exception:
                pass
            try:
                val_a = str(df_raw.iloc[r_idx, 0]).strip().upper()
                if "SETORAN" in val_a or "TOTAL" in val_a or "TUTUP" in val_a:
                    continue
                val_h = df_raw.iloc[r_idx, 7] if df_raw.shape[1] > 7 else 0
                if pd.notna(val_h) and str(val_h).strip().lower() != 'nan':
                    total_nontunai += float(str(val_h).replace(",", ""))
            except Exception:
                pass
        grand_setoran = total_tunai + total_nontunai
        self.lbl_setoran_tunai.config(text=f"Total Tunai: Rp. {total_tunai:,.0f}")
        self.lbl_setoran_nontunai.config(text=f"Total Non-Tunai: Rp. {total_nontunai:,.0f}")
        self.lbl_setoran_grand.config(text=f"GRAND TOTAL SETORAN: Rp. {grand_setoran:,.0f}")

    def reprint_nota_dari_sheet(self):
        selected_items = self.report_tree.selection()
        if not selected_items:
            messagebox.showwarning("Peringatan", "Pilih baris transaksi pada tabel laporan terlebih dahulu untuk di-reprint!")
            return
        sheet_name = self.combo_sheet_report.get()
        if not sheet_name or not self.target_file_path:
            return
        try:
            if self.cached_report_df is not None and self.cached_sheet_name == sheet_name:
                df_raw = self.cached_report_df
            else:
                df_raw = pd.read_excel(self.target_file_path, sheet_name=sheet_name, header=None)

            selected_row_values = self.report_tree.item(selected_items[0], "values")
            if not selected_row_values:
                return

            target_no_trx = None
            clicked_excel_row_idx = -1
            for r_idx in range(2, len(df_raw)):
                try:
                    r_vals = [str(df_raw.iloc[r_idx, c]) if pd.notna(df_raw.iloc[r_idx, c]) else "" for c in range(min(3, df_raw.shape[1]))]
                    if r_vals and str(r_vals[0]).strip() == str(selected_row_values[0]).strip() and str(r_vals[1]).strip() == str(selected_row_values[1]).strip():
                        clicked_excel_row_idx = r_idx
                        break
                except Exception:
                    continue

            if clicked_excel_row_idx == -1:
                for i, child_id in enumerate(self.report_tree.get_children()):
                    if child_id == selected_items[0]:
                        clicked_excel_row_idx = i + 2
                        break

            if clicked_excel_row_idx == -1:
                messagebox.showerror("Error", "Gagal mendeteksi baris transaksi pada Excel.")
                return

            current_r = clicked_excel_row_idx
            while current_r >= 2:
                try:
                    val_a = str(df_raw.iloc[current_r, 0]).strip()
                    if val_a and val_a.replace(".", "").isdigit():
                        target_no_trx = val_a
                        break
                except Exception:
                    pass
                current_r -= 1

            if not target_no_trx:
                target_no_trx = "1"

            items_reprint = []
            metode_reprint = "TUNAI"
            waktu_reprint = datetime.now().strftime("%H:%M:%S")
            member_reprint = ""
            start_scan = current_r
            while start_scan < len(df_raw):
                try:
                    val_a_check = str(df_raw.iloc[start_scan, 0]).strip()
                    if val_a_check and val_a_check.replace(".", "").isdigit() and val_a_check != target_no_trx and start_scan != current_r:
                        break
                    barcode = str(df_raw.iloc[start_scan, 1]) if pd.notna(df_raw.iloc[start_scan, 1]) else ""
                    judul = str(df_raw.iloc[start_scan, 2]) if pd.notna(df_raw.iloc[start_scan, 2]) else ""
                    qty = float(df_raw.iloc[start_scan, 3]) if pd.notna(df_raw.iloc[start_scan, 3]) else 1
                    harga = float(df_raw.iloc[start_scan, 4]) if pd.notna(df_raw.iloc[start_scan, 4]) else 0
                    subtotal = float(df_raw.iloc[start_scan, 5]) if pd.notna(df_raw.iloc[start_scan, 5]) else (qty * harga)
                    val_g = df_raw.iloc[start_scan, 6] if df_raw.shape[1] > 6 else None
                    val_h = df_raw.iloc[start_scan, 7] if df_raw.shape[1] > 7 else None
                    if pd.notna(val_g) and str(val_g).strip() != "":
                        metode_reprint = "TUNAI"
                    elif pd.notna(val_h) and str(val_h).strip() != "":
                        metode_reprint = "NON TUNAI"
                    diskon = float(df_raw.iloc[start_scan, 8]) if df_raw.shape[1] > 8 and pd.notna(df_raw.iloc[start_scan, 8]) else 0.0
                    member_val = str(df_raw.iloc[start_scan, 9]) if df_raw.shape[1] > 9 and pd.notna(df_raw.iloc[start_scan, 9]) else ""
                    if member_val:
                        member_reprint = member_val
                    time_val = str(df_raw.iloc[start_scan, 10]) if df_raw.shape[1] > 10 and pd.notna(df_raw.iloc[start_scan, 10]) else ""
                    if time_val:
                        waktu_reprint = time_val
                    if barcode or judul:
                        items_reprint.append({"barcode": barcode, "judul": judul, "jumlah": int(qty), "harga_akhir": harga, "total": subtotal, "diskon": diskon})
                except Exception:
                    pass
                start_scan += 1

            if not items_reprint:
                messagebox.showwarning("Peringatan", "Tidak ada item valid yang ditemukan untuk transaksi ini.")
                return

            grand_total_rep = sum(i['total'] for i in items_reprint)
            self.cetak_nota_thermal_80mm_custom(no_trx=target_no_trx, metode=metode_reprint, bayar=grand_total_rep, kembali=grand_total_rep, catatan_member=member_reprint, waktu_str=waktu_reprint, items_source=items_reprint, is_reprint=True)
            messagebox.showinfo("Sukses Re-Print", f"Nota Transaksi No. #{target_no_trx} berhasil dicetak ulang (Re-Print)!")
        except Exception as e:
            self._show_safe_error(self.root, "Error Re-Print", "Gagal memproses re-print nota. Silakan coba lagi atau hubungi admin.", exception=e, context="reprint_nota_dari_sheet")

    def proses_setoran_harian(self):
        if not self.target_file_path or not self.wb_target:
            messagebox.showwarning("Peringatan", "Pilih File Target Output terlebih dahulu!")
            return
        sheet_name = self.combo_sheet_report.get()
        if not sheet_name or sheet_name not in self.wb_target.sheetnames:
            messagebox.showwarning("Peringatan", "Pilih sheet laporan yang valid!")
            return
        try:
            ws = self.wb_target[sheet_name]
            if self.cached_report_df is not None and self.cached_sheet_name == sheet_name:
                df_raw = self.cached_report_df
            else:
                df_raw = pd.read_excel(self.target_file_path, sheet_name=sheet_name, header=None)

            total_tunai = 0.0
            total_nontunai = 0.0
            for r_idx in range(2, len(df_raw)):
                try:
                    val_a = str(df_raw.iloc[r_idx, 0]).strip().upper()
                    if "SETORAN" in val_a or "TOTAL" in val_a or "TUTUP" in val_a:
                        continue
                    vg = df_raw.iloc[r_idx, 6]
                    if pd.notna(vg): total_tunai += float(str(vg).replace(",", ""))
                except Exception:
                    pass
                try:
                    val_a = str(df_raw.iloc[r_idx, 0]).strip().upper()
                    if "SETORAN" in val_a or "TOTAL" in val_a or "TUTUP" in val_a:
                        continue
                    vh = df_raw.iloc[r_idx, 7]
                    if pd.notna(vh): total_nontunai += float(str(vh).replace(",", ""))
                except Exception:
                    pass

            grand_setoran = total_tunai + total_nontunai
            max_r = ws.max_row + 2
            ws.merge_cells(start_row=max_r, start_column=1, end_row=max_r, end_column=4)
            ws[f"A{max_r}"] = "SETORAN HARIAN (TUTUP BUKU)"
            ws[f"G{max_r}"] = total_tunai
            ws[f"H{max_r}"] = total_nontunai
            ws[f"F{max_r}"] = grand_setoran

            bold_font = Font(name="Calibri", size=11, bold=True)
            currency_format = '"Rp" #,##0'
            thin_border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
            for col_l in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]:
                cell = ws[f"{col_l}{max_r}"]
                cell.font = bold_font
                cell.border = thin_border
                if col_l in ["F", "G", "H"]:
                    cell.number_format = currency_format

            self.wb_target.save(self.target_file_path)
            self.cached_report_df = None
            self.cached_sheet_name = None
            self.muat_tabel_laporan_excel()
            messagebox.showinfo("Sukses", f"Setoran Harian berhasil ditutup & disimpan ke Excel!")
        except Exception as e:
            self._show_safe_error(self.root, "Error", "Gagal memproses setoran harian. Pastikan file Excel tidak sedang dibuka di aplikasi lain.", exception=e, context="proses_setoran_harian")

    def buka_popup_edit_qty_langsung(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        item_id = selected_items[0]
        item_index = self.tree.index(item_id)
        if not (0 <= item_index < len(self.cart)):
            return
        current_cart_item = self.cart[item_index]

        popup = tk.Toplevel(self.root)
        popup.title("Ubah Jumlah Qty")
        self.center_popup(popup, 320, 160)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: popup.destroy())

        ttk.Label(popup, text=f"Produk: {current_cart_item['judul']}", font=("Segoe UI", 9, "bold"), wraplength=300).pack(pady=(10, 5))
        row_q = ttk.Frame(popup)
        row_q.pack(pady=5)
        ttk.Label(row_q, text="Jumlah Baru (Qty):").pack(side="left", padx=5)
        ent_new_qty = ttk.Entry(row_q, width=10, font=("Segoe UI", 11, "bold"))
        ent_new_qty.insert(0, str(current_cart_item['jumlah']))
        ent_new_qty.pack(side="left", padx=5)
        ent_new_qty.focus_force()
        ent_new_qty.selection_range(0, tk.END)

        def save_new_qty(event=None):
            try:
                new_q = int(ent_new_qty.get().strip())
                if new_q <= 0: raise ValueError()
                current_cart_item['jumlah'] = new_q
                current_cart_item['total'] = new_q * current_cart_item['harga_akhir']
                self.update_tabel_keranjang()
                popup.destroy()
                self.ent_search.focus_force()
            except Exception:
                messagebox.showwarning("Peringatan", "Masukkan angka bulat positif untuk jumlah!", parent=popup)

        popup.bind("<Return>", save_new_qty)
        ttk.Button(popup, text="Simpan (ENTER)", command=save_new_qty).pack(pady=10)

    def cek_dan_buat_header_excel(self, ws):
        val_a1 = ws["A1"].value
        if str(val_a1).strip().lower() not in ["no transaksi", "no"]:
            thin_border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
            header_font = Font(name="Calibri", size=11, bold=True)
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            fill_standard = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            fill_nontunai = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
            fill_noman = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
            single_cols = [("A", "No"), ("B", "Kode Produk"), ("C", "Judul"), ("D", "Jumlah"), ("E", "Harga satuan"), ("F", "Total Per Produk"), ("I", "Diskon"), ("J", "Member"), ("K", "Waktu"), ("L", "NOMAN")]
            for col, text in single_cols:
                ws.merge_cells(f"{col}1:{col}2")
                cell1 = ws[f"{col}1"]
                cell2 = ws[f"{col}2"]
                cell1.value = text
                cell1.font = header_font
                cell1.alignment = align_center
                cell2.font = header_font
                cell2.alignment = align_center
                if col == "L":
                    cell1.fill = fill_noman
                    cell2.fill = fill_noman
                else:
                    cell1.fill = fill_standard
                    cell2.fill = fill_standard
                cell1.border = thin_border
                cell2.border = thin_border
            ws.merge_cells("G1:H1")
            cell_g1 = ws["G1"]
            cell_g1.value = "Metode"
            cell_g1.font = header_font
            cell_g1.alignment = align_center
            cell_g1.fill = fill_standard
            ws["G1"].border = thin_border
            ws["H1"].border = thin_border
            ws["G2"].value = "TUNAI"
            ws["G2"].font = header_font
            ws["G2"].alignment = align_center
            ws["G2"].fill = fill_standard
            ws["G2"].border = thin_border
            ws["H2"].value = "NON TUNAI"
            ws["H2"].font = header_font
            ws["H2"].alignment = align_center
            ws["H2"].fill = fill_nontunai
            ws["H2"].border = thin_border
            self.wb_target.save(self.target_file_path)

    def hitung_nomor_transaksi_berikutnya(self):
        if not self.target_file_path or not self.wb_target:
            self.no_trx_counter = 1
            return
        try:
            sheet_name = self.combo_sheet.get()
            if not sheet_name or sheet_name not in self.wb_target.sheetnames:
                self.no_trx_counter = 1
                return
            ws = self.wb_target[sheet_name]
            self.cek_dan_buat_header_excel(ws)
            max_r = ws.max_row
            last_trx = 0
            for r in range(max_r, 2, -1):
                val = ws[f"A{r}"].value
                if val is not None:
                    try:
                        last_trx = int(float(str(val).strip()))
                        break
                    except ValueError:
                        continue
            self.no_trx_counter = last_trx + 1 if last_trx > 0 else 1
        except Exception as e:
            logging.error(f"Gagal membaca nomor transaksi terakhir: {e}", exc_info=True)
            self.no_trx_counter = 1

    def on_sheet_changed(self, event=None):
        self.hitung_nomor_transaksi_berikutnya()

    def fokus_ke_tabel(self, event=None):
        children = self.tree.get_children()
        if children:
            self.tree.focus_set()
            if not self.tree.selection():
                self.tree.selection_set(children[0])
                self.tree.focus(children[0])

    def hapus_item_terpilih(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Peringatan", "Pilih item di tabel yang ingin dihapus terlebih dahulu!")
            return
        selected_item = selected_items[0]
        item_index = self.tree.index(selected_item)
        if 0 <= item_index < len(self.cart):
            del self.cart[item_index]
            self.update_tabel_keranjang()
            children = self.tree.get_children()
            if children:
                new_index = min(item_index, len(children) - 1)
                self.tree.selection_set(children[new_index])
                self.tree.focus(children[new_index])
                self.tree.focus_set()
            else:
                self.ent_search.focus_force()

    def load_master_data_async(self):
        df, is_online = self.sync_tool.load_data()
        self.df_master = df
        self.is_online = is_online
        self.is_loading = False

        def update_label():
            if self.is_online:
                self.lbl_status.config(text=f"[  ONLINE Cloud: {len(self.df_master):,} Produk ]", foreground="green")
            else:
                self.lbl_status.config(text=f"[ 🟡 OFFLINE Cache: {len(self.df_master):,} Produk ]", foreground="orange")
            self.dm_current_page = 1
            self.refresh_data_master_view()

        self.root.after(0, update_label)

    def pilih_target_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            MAX_FILE_SIZE_MB = 50
            try:
                file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
                if file_size_mb > MAX_FILE_SIZE_MB:
                    confirm = messagebox.askyesno("Peringatan File Besar", f"Ukuran file Excel sangat besar ({file_size_mb:.1f} MB).\nOperasi bisa menjadi lambat atau menyebabkan crash.\n\nApakah Anda yakin ingin menggunakan file ini?")
                    if not confirm:
                        return
            except OSError:
                pass

            self.target_file_path = file_path
            self.lbl_target_path.config(text=os.path.basename(file_path), foreground="green")
            try:
                self.wb_target = openpyxl.load_workbook(file_path)
                sheets = self.wb_target.sheetnames
                self.combo_sheet['values'] = sheets
                if sheets:
                    self.combo_sheet.current(0)
                self.cached_report_df = None
                self.cached_sheet_name = None
                self.noman_data = {}
                self.noman_item_map = {}
                self.hitung_nomor_transaksi_berikutnya()
                self.buka_popup_pilih_atau_tambah_sheet_pertama()
            except Exception as e:
                self._show_safe_error(self.root, "Error Excel", "Gagal membaca file target. Pastikan file bukan file Excel yang rusak atau sedang dibuka di aplikasi lain.", exception=e, context="pilih_target_file")

    def buka_popup_pilih_atau_tambah_sheet_pertama(self):
        popup = tk.Toplevel(self.root)
        popup.title("Pilih / Tambah Sheet Target")
        self.center_popup(popup, 360, 210)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: popup.destroy())

        ttk.Label(popup, text="Pilih Sheet untuk Transaksi:", font=("Segoe UI", 10, "bold")).pack(pady=(15, 5))
        sheets = self.wb_target.sheetnames if self.wb_target else []
        combo_chosen_sheet = ttk.Combobox(popup, values=sheets, width=28, state="readonly")
        if sheets:
            combo_chosen_sheet.current(0)
        combo_chosen_sheet.pack(pady=5)

        def aksi_tambah_sheet_di_popup():
            popup.destroy()
            self.tambah_sheet_baru()

        btn_frame = ttk.Frame(popup)
        btn_frame.pack(pady=15)
        btn_tambah_plus = ttk.Button(btn_frame, text="➕ Tambah Sheet Baru", command=aksi_tambah_sheet_di_popup)
        btn_tambah_plus.pack(side="left", padx=5)

        def aksi_pilih():
            selected_sheet = combo_chosen_sheet.get()
            if selected_sheet:
                self.combo_sheet.set(selected_sheet)
                self.on_sheet_changed()
                popup.destroy()
                self.ent_search.focus_force()
            else:
                messagebox.showwarning("Peringatan", "Pilih salah satu sheet terlebih dahulu!", parent=popup)

        btn_ok = ttk.Button(btn_frame, text="Gunakan Sheet Ini", command=aksi_pilih)
        btn_ok.pack(side="left", padx=5)
        popup.bind("<Return>", lambda e: aksi_pilih())

    def tambah_sheet_baru(self):
        if not self.target_file_path or not self.wb_target:
            messagebox.showwarning("Peringatan", "Pilih File Target Output terlebih dahulu!")
            return

        popup = tk.Toplevel(self.root)
        popup.title("Tambah Sheet Baru")
        self.center_popup(popup, 300, 130)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: popup.destroy())

        ttk.Label(popup, text="Nama Sheet Baru:").pack(pady=10)
        ent_sheet_name = ttk.Entry(popup, width=25)
        ent_sheet_name.pack(pady=5)
        ent_sheet_name.focus_force()

        def submit_sheet():
            name = ent_sheet_name.get().strip()
            if name:
                if name in self.wb_target.sheetnames:
                    messagebox.showwarning("Peringatan", "Nama sheet sudah ada!", parent=popup)
                else:
                    new_ws = self.wb_target.create_sheet(title=name)
                    self.cek_dan_buat_header_excel(new_ws)
                    self.wb_target.save(self.target_file_path)
                    self.combo_sheet['values'] = self.wb_target.sheetnames
                    self.combo_sheet.set(name)
                    self.cached_report_df = None
                    self.cached_sheet_name = None
                    self.hitung_nomor_transaksi_berikutnya()
                    popup.destroy()

        popup.bind("<Return>", lambda e: submit_sheet())
        ttk.Button(popup, text="Tambah", command=submit_sheet).pack(pady=5)

    def proses_scan_trigger(self, event=None):
        if self.is_loading:
            messagebox.showinfo("Proses Loading", "Database master sedang dimuat dari cloud. Harap tunggu sebentar!")
            return
        raw_input = self.ent_search.get().strip()
        if not raw_input:
            return
        if not self.target_file_path or not self.combo_sheet.get():
            messagebox.showwarning("File Output Belum Dipilih", "Harap pilih 'File Target Output' & 'Sheet Target' terlebih dahulu!")
            return

        match_barcode = self.df_master[self.df_master['barcode'].astype(str) == raw_input]
        if not match_barcode.empty:
            item = match_barcode.iloc[0]
            self.buka_popup_input_produk(item)
            self.ent_search.delete(0, tk.END)
            return

        match_manual = self.df_master[
            self.df_master['barcode'].astype(str).str.contains(raw_input, case=False, na=False, regex=False) |
            self.df_master['judul'].astype(str).str.contains(raw_input, case=False, na=False, regex=False)
        ]
        if not match_manual.empty:
            self.buka_popup_pilihan_produk(match_manual)
            self.ent_search.delete(0, tk.END)
        else:
            messagebox.showwarning("Tidak Ditemukan", f"Produk dengan kata kunci '{raw_input}' tidak ditemukan!")

    def buka_popup_pilihan_produk(self, df_pilihan):
        popup = tk.Toplevel(self.root)
        popup.title("Pilih Produk")
        self.center_popup(popup, 580, 380)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: (popup.destroy(), self.ent_search.focus_force()))

        ttk.Label(popup, text="Gunakan Panah Atas/Bawah lalu tekan ENTER untuk memilih:", font=("Segoe UI", 9, "bold")).pack(pady=8)
        frame_list = ttk.Frame(popup)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("barcode", "judul", "harga")
        tree_select = ttk.Treeview(frame_list, columns=cols, show="headings", selectmode="browse",
                                   style="Unified.Treeview")
        tree_select.heading("barcode", text="Barcode")
        tree_select.heading("judul", text="Judul")
        tree_select.heading("harga", text="Harga")
        tree_select.column("barcode", width=120)
        tree_select.column("judul", width=310)
        tree_select.column("harga", width=100, anchor="e")
        sb = ttk.Scrollbar(frame_list, orient="vertical", command=tree_select.yview)
        tree_select.configure(yscroll=sb.set)
        sb.pack(side="right", fill="y")
        tree_select.pack(side="left", fill="both", expand=True)

        for _, row in df_pilihan.iterrows():
            harga_val = float(row['harga'])
            tree_select.insert("", "end", values=(row['barcode'], row['judul'], f"Rp. {harga_val:,.0f}"))

        children = tree_select.get_children()
        if children:
            tree_select.selection_set(children[0])
            tree_select.focus_set()
            tree_select.focus(children[0])

        def pilih_item_terpilih(event=None):
            selected = tree_select.selection()
            if not selected:
                return
            idx = tree_select.index(selected[0])
            item_selected = df_pilihan.iloc[idx]
            popup.destroy()
            self.buka_popup_input_produk(item_selected)

        tree_select.bind("<Return>", pilih_item_terpilih)
        popup.bind("<Return>", pilih_item_terpilih)
        ttk.Button(popup, text="Pilih Produk Ini (ENTER)", command=pilih_item_terpilih).pack(pady=8)

    def buka_popup_input_produk(self, item):
        popup = tk.Toplevel(self.root)
        popup.title("Input Jumlah & Diskon")
        self.center_popup(popup, 400, 260)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: (popup.destroy(), self.ent_search.focus_force()))

        ttk.Label(popup, text=f"Produk: {item['judul']}", font=("Segoe UI", 10, "bold"), wraplength=380).pack(pady=(10, 2))
        harga_norm = float(item['harga'])
        ttk.Label(popup, text=f"Harga Normal: Rp. {harga_norm:,.0f}", foreground="gray").pack(pady=(0, 10))

        row_qty = ttk.Frame(popup)
        row_qty.pack(pady=5)
        ttk.Label(row_qty, text="Jumlah (Qty):").pack(side="left", padx=5)
        ent_qty = ttk.Entry(row_qty, width=10, font=("Segoe UI", 10, "bold"))
        ent_qty.insert(0, "1")
        ent_qty.pack(side="left", padx=5)
        ent_qty.focus_force()
        ent_qty.selection_range(0, tk.END)

        frame_diskon = ttk.LabelFrame(popup, text=" Opsi Diskon ")
        frame_diskon.pack(fill="x", padx=20, pady=5)
        diskon_var = tk.StringVar(value="0")
        row_radio = ttk.Frame(frame_diskon)
        row_radio.pack(pady=5)
        ttk.Radiobutton(row_radio, text="0%", variable=diskon_var, value="0").pack(side="left", padx=5)
        ttk.Radiobutton(row_radio, text="5%", variable=diskon_var, value="5").pack(side="left", padx=5)
        ttk.Radiobutton(row_radio, text="20%", variable=diskon_var, value="20").pack(side="left", padx=5)
        ttk.Radiobutton(row_radio, text="25%", variable=diskon_var, value="25").pack(side="left", padx=5)
        row_custom = ttk.Frame(frame_diskon)
        row_custom.pack(pady=5)
        ttk.Label(row_custom, text="Diskon Custom (%):").pack(side="left", padx=5)
        ent_custom_diskon = ttk.Entry(row_custom, width=8)
        ent_custom_diskon.insert(0, "0")
        ent_custom_diskon.pack(side="left", padx=5)

        def submit_item(event=None):
            try:
                added_qty = int(ent_qty.get().strip())
                if added_qty <= 0: raise ValueError()
            except Exception:
                messagebox.showwarning("Peringatan", "Jumlah Qty harus angka bulat positif!", parent=popup)
                return
            custom_val = ent_custom_diskon.get().strip()
            if custom_val != "0" and custom_val != "":
                try: diskon_percent = float(custom_val)
                except Exception: diskon_percent = float(diskon_var.get())
            else:
                diskon_percent = float(diskon_var.get())
            harga_akhir = harga_norm * (1 - (diskon_percent / 100))
            barcode_str = str(item['barcode'])
            existing_item = None
            for cart_item in self.cart:
                if cart_item['barcode'] == barcode_str and cart_item['diskon'] == diskon_percent:
                    existing_item = cart_item
                    break
            if existing_item:
                existing_item['jumlah'] += added_qty
                existing_item['total'] = existing_item['jumlah'] * harga_akhir
            else:
                total = harga_akhir * added_qty
                self.cart.append({"barcode": barcode_str, "judul": str(item['judul']), "jumlah": added_qty, "harga_normal": harga_norm, "diskon": diskon_percent, "harga_akhir": harga_akhir, "total": total})
            self.update_tabel_keranjang()
            popup.destroy()
            self.ent_search.focus_force()

        popup.bind("<Return>", submit_item)
        btn_row = ttk.Frame(popup)
        btn_row.pack(pady=10)
        ttk.Button(btn_row, text="OK (ENTER)", command=submit_item).pack(side="left", padx=5)
        ttk.Button(btn_row, text="Batal (ESC)", command=lambda: (popup.destroy(), self.ent_search.focus_force())).pack(side="left", padx=5)

    def update_tabel_keranjang(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        grand_total = 0
        total_qty = 0
        for idx, item in enumerate(self.cart, 1):
            diskon_str = f"{int(item['diskon'])}%" if float(item['diskon']).is_integer() else f"{item['diskon']}%"
            self.tree.insert("", "end", values=(idx, item['barcode'], item['judul'], item['jumlah'], f"Rp. {item['harga_normal']:,.0f}", diskon_str, f"Rp. {item['harga_akhir']:,.0f}", f"Rp. {item['total']:,.0f}"))
            grand_total += item['total']
            total_qty += item['jumlah']
        self.lbl_total.config(text=f"TOTAL BELANJA: Rp. {grand_total:,.0f}")
        self.lbl_total_qty.config(text=f"TOTAL ITEM / QTY: {total_qty:,} Pcs")

    def reset_transaksi(self):
        if messagebox.askyesno("Konfirmasi", "Bersihkan keranjang belanja saat ini?"):
            self.cart.clear()
            self.update_tabel_keranjang()
            self.ent_search.focus_force()

    def setup_global_keyboard(self):
        self._tree_nav_pos = {}
        self.root.bind_all("<Return>", self._global_enter, add="+")
        self.root.bind_all("<Escape>", self._global_escape, add="+")
        self.root.bind_all("<Down>", lambda event: self._global_arrow(event, 1), add="+")
        self.root.bind_all("<Up>", lambda event: self._global_arrow(event, -1), add="+")
        self.root.bind("<F2>", lambda event: self.buka_popup_tambah_produk())

    def _event_widget(self, event):
        try:
            w = event.widget
        except Exception:
            return None
        if w is None or isinstance(w, str):
            return None
        try:
            if not w.winfo_exists():
                return None
        except Exception:
            return None
        return w

    def _popup_aktif(self):
        for w in self.root.winfo_children():
            if not isinstance(w, tk.Toplevel):
                continue
            nama = w.winfo_name().lower()
            if "popdown" in nama or "tooltip" in nama:
                continue
            try:
                if w.winfo_viewable():
                    return w
            except tk.TclError:
                pass
        return None

    def _popdown_aktif(self):
        try:
            g = str(self.root.tk.call("grab", "current"))
        except tk.TclError:
            return False
        return "popdown" in g.lower()

    def _punya_binding(self, widget, seq):
        try:
            if widget.bind(seq):
                return True
        except tk.TclError:
            return True
        try:
            tl = widget.winfo_toplevel()
            if tl.bind(seq):
                return True
        except tk.TclError:
            return True
        return False

    def _fokuskan(self, widget):
        try:
            widget.focus_force()
        except tk.TclError:
            pass

    def _daftar_fokus(self, container):
        hasil = []
        def rekursif(w):
            for c in w.winfo_children():
                if isinstance(c, tk.Toplevel):
                    continue
                cls = c.winfo_class()
                ikut = cls in ("Entry", "TEntry", "TButton", "Button",
                               "Combobox", "TCombobox", "Treeview", "Canvas")
                if cls == "Canvas" and not hasattr(c, "command"):
                    ikut = False
                if ikut:
                    try:
                        if not c.winfo_viewable():
                            continue
                    except tk.TclError:
                        continue
                    nonaktif = False
                    try:
                        nonaktif = str(c.cget("state")) == "disabled"
                    except tk.TclError:
                        try:
                            nonaktif = c.instate(["disabled"])
                        except (tk.TclError, AttributeError):
                            nonaktif = False
                    if not nonaktif:
                        hasil.append(c)
                rekursif(c)
        rekursif(container)
        hasil.sort(key=lambda w: (w.winfo_rooty(), w.winfo_rootx()))
        return hasil

    def _geser_fokus(self, asal, arah):
        daftar = self._daftar_fokus(self.root)
        if not daftar:
            return
        try:
            idx = daftar.index(asal)
        except ValueError:
            idx = -1 if arah > 0 else 0
        n = len(daftar)
        for i in range(1, n + 1):
            kandidat = daftar[(idx + arah * i) % n]
            try:
                if kandidat.winfo_class() == "Treeview" and not kandidat.get_children():
                    continue
            except tk.TclError:
                continue
            self._fokuskan(kandidat)
            if kandidat.winfo_class() == "Treeview":
                anak = kandidat.get_children()
                if anak:
                    kandidat.selection_set(anak[0])
                    kandidat.focus(anak[0])
                    try:
                        kandidat.see(anak[0])
                    except tk.TclError:
                        pass
            return

    def _global_enter(self, event):
        w = self._event_widget(event)
        if w is None:
            return
        if self._punya_binding(w, "<Return>"):
            return
        cls = w.winfo_class()
        if cls in ("TButton", "Button"):
            try:
                w.invoke()
            except tk.TclError:
                pass
            return
        if cls == "Canvas" and hasattr(w, "command"):
            try:
                w.command()
            except Exception:
                pass
            return
        if self._popup_aktif() is None and w is getattr(self, "tree", None):
            self.buka_popup_edit_qty_langsung()

    def _global_escape(self, event):
        popup = self._popup_aktif()
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
        w = self._event_widget(event)
        if self._popdown_aktif():
            return
        try:
            if self.frame_transaksi.winfo_viewable():
                self._fokuskan(self.ent_search)
            elif self.frame_data_master.winfo_viewable():
                self._fokuskan(self.dm_ent_filter)
            else:
                self._fokuskan(self.combo_sheet_report)
        except tk.TclError:
            pass

    def _global_arrow(self, event, arah):
        if self._popup_aktif() is not None:
            return
        w = self._event_widget(event)
        if w is None:
            return
        cur = None
        try:
            cur = self.root.focus_get()
        except tk.TclError:
            cur = None
        if cur is not None and cur is not w:
            return
        cls = w.winfo_class()
        if cls in ("Combobox", "TCombobox", "Listbox", "Spinbox", "Text"):
            return
        if cls == "Treeview":
            self._tree_arrow(w, arah)
            return
        self._geser_fokus(w, arah)

    def _tree_arrow(self, tv, arah):
        anak = tv.get_children()
        if not anak:
            return
        item_fokus = None
        try:
            item_fokus = tv.focus()
        except tk.TclError:
            item_fokus = None
        if not item_fokus or item_fokus not in anak:
            tv.selection_set(anak[0])
            tv.focus(anak[0])
            try:
                tv.see(anak[0])
            except tk.TclError:
                pass
            self._tree_nav_pos[str(tv)] = 0
            return
        idx = anak.index(item_fokus)
        sebelumnya = self._tree_nav_pos.get(str(tv), idx)
        if arah > 0 and idx >= len(anak) - 1 and sebelumnya == idx:
            self._geser_fokus(tv, 1)
        elif arah < 0 and idx <= 0 and sebelumnya == idx:
            self._geser_fokus(tv, -1)
        self._tree_nav_pos[str(tv)] = idx

    def cetak_nota_thermal_80mm_custom(self, no_trx, metode, bayar, kembali, catatan_member, waktu_str=None, items_source=None, is_reprint=False):
        items_to_print = items_source if items_source is not None else self.cart
        grand_total = sum(item['total'] for item in items_to_print)
        if not waktu_str:
            waktu_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        lebar = int(self.config_nota.get("lebar_kertas", 48))
        nama_kasir = self.config_nota.get("nama_kasir", "Admin")

        def _fmt(text, align, width):
            clean = str(text).replace("\n", "").replace("\r", "")
            if align == "center":
                return clean.center(width)
            elif align == "right":
                return clean.rjust(width)
            elif align == "full":
                if len(clean) == 1:
                    return clean * width
                return clean.ljust(width)[:width]
            else:
                return clean.ljust(width)[:width]

        struk = []
        template = self.template_nota.get("rows", [])

        for row in template:
            rtype = row.get("type", "text")
            align = row.get("align", "center")
            content = row.get("content", "")

            if rtype == "empty":
                struk.append("")
            elif rtype == "text":
                struk.append(_fmt(content, align, lebar))
            elif rtype == "line":
                char = content if content else "-"
                struk.append(_fmt(char, "full", lebar))
            elif rtype == "header_trx":
                struk.append(f"No. Trx  : #{no_trx}".ljust(lebar // 2) +
                             f"Kasir : {nama_kasir}".rjust(lebar - (lebar // 2)))
                struk.append(f"Waktu   : {waktu_str}".ljust(lebar))
                if catatan_member:
                    struk.append(f"Member  : {catatan_member}".ljust(lebar))
            elif rtype == "products":
                for item in items_to_print:
                    barcode_str = str(item.get('barcode', '') or '').strip()
                    if barcode_str.lower() in ('nan', 'none'):
                        barcode_str = ''
                    if barcode_str.endswith('.0') and barcode_str[:-2].isdigit():
                        barcode_str = barcode_str[:-2]
                    judul = str(item.get('judul', ''))
                    baris_produk = f"{barcode_str} {judul}".strip() if barcode_str else judul
                    struk.append(baris_produk[:lebar])
                    qty_price = f"  {item['jumlah']} x @{item['harga_akhir']:,.0f}"
                    subtotal = f"Rp. {item['total']:,.0f}"
                    spasi = max(2, lebar - len(qty_price) - len(subtotal))
                    struk.append(qty_price + (" " * spasi) + subtotal)
                    if item.get('diskon', 0) > 0:
                        struk.append(f"  (Diskon {item['diskon']}%)".ljust(lebar))
            elif rtype == "total":
                struk.append(f"TOTAL BELANJA".ljust(lebar // 2) +
                             f"Rp. {grand_total:,.0f}".rjust(lebar - (lebar // 2)))
            elif rtype == "payment":
                struk.append(f"PEMBAYARAN ({metode})".ljust(lebar // 2) +
                             f"Rp. {bayar:,.0f}".rjust(lebar - (lebar // 2)))
                if metode == "TUNAI":
                    struk.append(f"KEMBALIAN".ljust(lebar // 2) +
                                 f"Rp. {kembali:,.0f}".rjust(lebar - (lebar // 2)))
            elif rtype == "footer_text":
                footer1 = self.config_nota.get("info_tambahan", "")
                footer2 = self.config_nota.get("pesan_penutup", "")
                if footer1:
                    struk.append(_fmt(footer1, align, lebar))
                if footer2:
                    struk.append(_fmt(footer2, align, lebar))

        if is_reprint:
            struk.insert(0, "*** RE-PRINT NOTA KASIR ***".center(lebar))

        PAPER_CUT = "\x1dV\x41\x00"
        text_struk_final = "\n".join(struk) + "\n\n\n\n" + PAPER_CUT

        try:
            if os.name == 'nt':
                try:
                    import win32print
                except ImportError:
                    messagebox.showerror("Error Printer", "Modul 'pywin32' tidak ditemukan.\nSilakan instal via: pip install pywin32")
                    return
                printer_name = win32print.GetDefaultPrinter()
                hPrinter = win32print.OpenPrinter(printer_name)
                try:
                    win32print.StartDocPrinter(hPrinter, 1, ("Nota Transaksi Kasir", None, "RAW"))
                    win32print.StartPagePrinter(hPrinter)
                    win32print.WritePrinter(hPrinter, text_struk_final.encode('utf-8', errors='ignore'))
                    win32print.EndPagePrinter(hPrinter)
                    win32print.EndDocPrinter(hPrinter)
                finally:
                    win32print.ClosePrinter(hPrinter)
            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as tmp:
                    tmp.write(text_struk_final)
                    tmp_path = tmp.name
                try:
                    subprocess.run(["lpr", "-o", "raw", tmp_path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10)
                except subprocess.CalledProcessError as e:
                    logging.error(f"lpr gagal dengan code {e.returncode}", exc_info=True)
                    raise
                except subprocess.TimeoutExpired:
                    logging.error("lpr timeout setelah 10 detik", exc_info=True)
                    raise
                finally:
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
        except Exception as e:
            logging.error(f"Gagal mencetak otomatis: {e}", exc_info=True)
            messagebox.showerror("Error Printer", "Gagal mencetak struk ke printer thermal. Silakan periksa koneksi printer.")

    def buka_popup_pembayaran(self):
        if not self.cart:
            messagebox.showwarning("Peringatan", "Keranjang belanja masih kosong!")
            return
        if not self.target_file_path or not self.combo_sheet.get():
            messagebox.showwarning("Peringatan", "Silakan pilih File Target Output & Sheet Target terlebih dahulu!")
            return

        grand_total = sum(item['total'] for item in self.cart)
        popup = tk.Toplevel(self.root)
        popup.title("Pembayaran")
        self.center_popup(popup, 380, 440)
        popup.transient(self.root)
        popup.grab_set()
        popup.bind("<Escape>", lambda e: (popup.destroy(), self.ent_search.focus_force()))

        ttk.Label(popup, text="Total Transaksi:", font=("Segoe UI", 11, "bold")).pack(pady=(10, 2))
        ttk.Label(popup, text=f"Rp. {grand_total:,.0f}", font=("Segoe UI", 16, "bold"), foreground="blue").pack(pady=(0, 5))

        ada_diskon = any(item['diskon'] > 0 for item in self.cart)
        frame_member_pay = ttk.LabelFrame(popup, text=" Catatan Member ")
        frame_member_pay.pack(fill="x", padx=30, pady=5)
        ent_catatan_pay = ttk.Entry(frame_member_pay, font=("Segoe UI", 10))
        ent_catatan_pay.pack(fill="x", padx=10, pady=5)
        if ada_diskon:
            ent_catatan_pay.config(state="normal")
        else:
            ent_catatan_pay.config(state="disabled")
            ent_catatan_pay.insert(0, "(Tidak ada diskon / Non-Member)")

        ttk.Label(popup, text="Pilih Jenis Pembayaran:").pack(anchor="w", padx=30, pady=(5, 0))
        pay_var = tk.StringVar(value="TUNAI")

        frame_tunai_input = ttk.Frame(popup)
        ttk.Label(frame_tunai_input, text="Nominal Bayar (Rp):").pack(anchor="w", pady=(5, 2))
        ent_bayar = ttk.Entry(frame_tunai_input, font=("Segoe UI", 11, "bold"))
        ent_bayar.pack(fill="x")
        lbl_kembalian = ttk.Label(frame_tunai_input, text="Kembalian: Rp. 0", font=("Segoe UI", 10, "bold"), foreground="green")
        lbl_kembalian.pack(anchor="w", pady=(5, 0))

        def hitung_kembalian(event=None):
            raw = ent_bayar.get().replace(".", "").replace(",", "").strip()
            if raw.isdigit():
                bayar = float(raw)
                kembali = bayar - grand_total
                if kembali >= 0:
                    lbl_kembalian.config(text=f"Kembalian: Rp. {kembali:,.0f}", foreground="green")
                else:
                    lbl_kembalian.config(text=f"Kurang: Rp. {abs(kembali):,.0f}", foreground="red")
            else:
                lbl_kembalian.config(text="Kembalian: Rp. 0", foreground="green")

        ent_bayar.bind("<KeyRelease>", hitung_kembalian)

        def toggle_pembayaran():
            if pay_var.get() == "TUNAI":
                frame_tunai_input.pack(fill="x", padx=30, pady=5)
                ent_bayar.focus_force()
            else:
                frame_tunai_input.pack_forget()

        row_radio = ttk.Frame(popup)
        row_radio.pack(fill="x", padx=30, pady=5)
        ttk.Radiobutton(row_radio, text="TUNAI", variable=pay_var, value="TUNAI", command=toggle_pembayaran).pack(side="left", padx=10)
        ttk.Radiobutton(row_radio, text="NON TUNAI", variable=pay_var, value="NON TUNAI", command=toggle_pembayaran).pack(side="left", padx=10)
        toggle_pembayaran()

        def simpan_dan_proses(event=None):
            catatan_val = ""
            if ada_diskon:
                catatan_val = ent_catatan_pay.get().strip()
                if not catatan_val:
                    messagebox.showwarning("Peringatan", "Catatan Member wajib diisi karena ada diskon!", parent=popup)
                    ent_catatan_pay.focus_force()
                    return
            metode = pay_var.get()
            tunai_val = 0
            nontunai_val = 0
            bayar_val = grand_total
            kembali_val = 0
            if metode == "TUNAI":
                raw_bayar = ent_bayar.get().replace(".", "").replace(",", "").strip()
                if not raw_bayar.isdigit() or float(raw_bayar) < grand_total:
                    messagebox.showwarning("Peringatan", "Nominal pembayaran Tunai kurang!", parent=popup)
                    return
                bayar_val = float(raw_bayar)
                tunai_val = grand_total
                kembali_val = bayar_val - grand_total
            else:
                nontunai_val = grand_total

            try:
                sheet_name = self.combo_sheet.get()
                ws = self.wb_target[sheet_name]
                self.cek_dan_buat_header_excel(ws)
                self.hitung_nomor_transaksi_berikutnya()
                no_trx_num = self.no_trx_counter
                timestamp_str = datetime.now().strftime("%H:%M:%S")
                currency_format = '"Rp" #,##0'
                thin_border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
                all_used_cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
                for idx_item, item in enumerate(self.cart):
                    max_r = ws.max_row + 1
                    if idx_item == 0: ws[f"A{max_r}"] = no_trx_num
                    ws[f"B{max_r}"] = item['barcode']
                    ws[f"C{max_r}"] = item['judul']
                    ws[f"D{max_r}"] = item['jumlah']
                    ws[f"E{max_r}"] = item['harga_akhir']
                    ws[f"E{max_r}"].number_format = currency_format
                    ws[f"F{max_r}"] = item['total']
                    ws[f"F{max_r}"].number_format = currency_format
                    if idx_item == 0:
                        if metode == "TUNAI":
                            ws[f"G{max_r}"] = tunai_val
                            ws[f"G{max_r}"].number_format = currency_format
                        else:
                            ws[f"H{max_r}"] = nontunai_val
                            ws[f"H{max_r}"].number_format = currency_format
                    ws[f"I{max_r}"] = item['diskon']
                    ws[f"J{max_r}"] = catatan_val if ada_diskon else ""
                    if idx_item == 0: ws[f"K{max_r}"] = timestamp_str
                    ws[f"L{max_r}"] = ""
                    for col_l in all_used_cols:
                        ws[f"{col_l}{max_r}"].border = thin_border

                self.wb_target.save(self.target_file_path)
                self.cached_report_df = None
                self.cached_sheet_name = None
                self.cetak_nota_thermal_80mm_custom(no_trx_num, metode, bayar_val, kembali_val, catatan_val if ada_diskon else "", is_reprint=False)
                messagebox.showinfo("Sukses", f"Transaksi No. {no_trx_num} berhasil disimpan & dicetak!")
                self.cart.clear()
                self.update_tabel_keranjang()
                self.hitung_nomor_transaksi_berikutnya()
                popup.destroy()
                self.ent_search.focus_force()
                self.refresh_data_laporan()
            except Exception as e:
                self._show_safe_error(popup, "Error Simpan", "Gagal menyimpan transaksi. Pastikan file Excel tidak sedang dibuka di aplikasi lain.", exception=e, context="simpan_dan_proses")

        popup.bind("<Return>", simpan_dan_proses)
        ttk.Button(popup, text="SIMPAN & CETAK NOTA (ENTER)", command=simpan_dan_proses).pack(side="bottom", pady=10)


if __name__ == "__main__":
    setup_app_logging()
    root = tk.Tk()
    app = KasirApp(root)
    root.mainloop()
