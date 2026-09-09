import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from datetime import datetime
import os
import threading
import json
import urllib.request

# Import modul sync master cloud/cache
from sync_master import MasterDataSync

# URL mentah ke version.json di GitHub Anda untuk mengambil changelog secara dinamis
VERSION_URL = "https://raw.githubusercontent.com/semogaberdampak/backup_app/main/version.json"

class KasirApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Takom Kasir v0.9.5")

        # AUTO LAUNCH FULL SCREEN (MAXIMIZED)
        self.root.state('zoomed')

        try:
            self.root.iconbitmap("logo.ico")
        except:
            pass

        self.sync_tool = MasterDataSync()
        self.df_master = pd.DataFrame()
        self.is_online = False
        self.is_loading = True
        
        # File & Sheet Target
        self.target_file_path = ""
        self.wb_target = None
        
        # Keranjang Transaksi
        self.cart = []
        self.no_trx_counter = 1

        # Ambil Changelog secara dinamis dari GitHub untuk Patch Log Notes
        self.dynamic_changelog = self.fetch_remote_changelog()

        self.setup_ui()
        threading.Thread(target=self.load_master_data_async, daemon=True).start()

    def fetch_remote_changelog(self):
        """Mengambil data changelog secara online dari version.json di GitHub"""
        default_log = "FITUR UTAMA\n- Sistem Kasir Berjalan Normal\n- Menggunakan modul pembaruan otomatis."
        try:
            req = urllib.request.urlopen(VERSION_URL, timeout=3)
            data = json.loads(req.read().decode('utf-8'))
            return data.get("changelog", default_log)
        except Exception as e:
            print("Gagal mengambil changelog online:", e)
            return "Mode Offline / Gagal memuat Changelog terbaru dari server."

    def center_popup(self, popup, width, height):
        popup.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()

        x = root_x + (root_w // 2) - (width // 2)
        y = root_y + (root_h // 2) - (height // 2)
        popup.geometry(f'{width}x{height}+{max(0, x)}+{max(0, y)}')

    def setup_ui(self):
        # =========================================================================
        # BARIS ATAS: PENGATURAN FILE & SHEET + PATCH LOG NOTES
        # =========================================================================
        top_container = ttk.Frame(self.root)
        top_container.pack(fill="x", padx=10, pady=5)

        # Panel Kiri: Pengaturan File & Sheet Target (Mapping sudah dipatenkan)
        frame_config = ttk.LabelFrame(top_container, text=" Pengaturan File & Sheet Target ")
        frame_config.pack(side="left", fill="both", expand=True, padx=(0, 5))

        row_master = ttk.Frame(frame_config)
        row_master.pack(fill="x", padx=5, pady=6)
        
        self.lbl_status = ttk.Label(row_master, text="[ 🔄 Connecting Database... Mohon Tunggu ]", font=("Arial", 9, "bold"), foreground="blue")
        self.lbl_status.pack(side="left", padx=5)

        row_target = ttk.Frame(frame_config)
        row_target.pack(fill="x", padx=5, pady=6)

        btn_target = ttk.Button(row_target, text="Pilih File Target Output", command=self.pilih_target_file)
        btn_target.pack(side="left", padx=5)

        self.lbl_target_path = ttk.Label(row_target, text="Belum dipilih", foreground="red", font=("Arial", 9, "bold"))
        self.lbl_target_path.pack(side="left", padx=5)

        ttk.Label(row_target, text="Pilih Sheet Target:").pack(side="left", padx=(20, 2))
        self.combo_sheet = ttk.Combobox(row_target, width=15, state="readonly")
        self.combo_sheet.pack(side="left", padx=2)
        self.combo_sheet.bind("<<ComboboxSelected>>", self.on_sheet_changed)

        btn_add_sheet = ttk.Button(row_target, text="+ Sheet", width=8, command=self.tambah_sheet_baru)
        btn_add_sheet.pack(side="left", padx=5)

        # Informasi Struktur Kolom Tetap (Standar Baku v0.9.2)
        info_struktur = ttk.Label(
            frame_config, 
            text="📌 Struktur Kolom Excel Baku (A s.d. K): [A] No Trx | [B] Kode Produk | [C] Judul | [D] Jumlah | [E] Harga Satuan | [F] Total Per Produk | [G-H] Metode (Tunai/Non-Tunai) | [I] Diskon | [J] Member | [K] Waktu", 
            font=("Arial", 8, "italic"),
            foreground="darkslategray"
        )
        info_struktur.pack(anchor="w", padx=10, pady=(4, 8))

        # Panel Kanan: Patch Log Notes (Dinamis dari GitHub)
        frame_notes = ttk.LabelFrame(top_container, text=" Patch Log Notes ")
        frame_notes.pack(side="right", fill="both", padx=(5, 0))

        self.txt_notes = tk.Text(frame_notes, width=35, height=6, font=("Consolas", 8))
        self.txt_notes.pack(fill="both", expand=True, padx=2, pady=2)
        
        # Masukkan teks dinamis hasil download dari version.json GitHub
        notes_content = f"[ PATCH LOG NOTES - ONLINE ]\n-----------------------------\n{self.dynamic_changelog}"
        self.txt_notes.insert("1.0", notes_content)
        self.txt_notes.config(state="disabled")

        # =========================================================================
        # AREA SCAN / CARI PRODUK
        # =========================================================================
        scan_frame = ttk.LabelFrame(self.root, text=" Area Scan / Cari Produk ")
        scan_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(scan_frame, text="Cari Barcode / Judul:", font=("Arial", 10, "bold")).pack(side="left", padx=5)
        self.ent_search = ttk.Entry(scan_frame, font=("Arial", 11))
        self.ent_search.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        self.ent_search.bind("<Return>", self.proses_scan_trigger)
        self.ent_search.bind("<Down>", self.fokus_ke_tabel)
        self.ent_search.focus()

        # =========================================================================
        # TABEL KERANJANG BELANJA
        # =========================================================================
        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("no", "barcode", "judul", "jumlah", "harga_normal", "diskon", "harga_akhir", "total")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("no", text="No")
        self.tree.heading("barcode", text="Barcode")
        self.tree.heading("judul", text="Judul")
        self.tree.heading("jumlah", text="Jumlah")
        self.tree.heading("harga_normal", text="Harga Normal")
        self.tree.heading("diskon", text="Diskon")
        self.tree.heading("harga_akhir", text="Harga Akhir")
        self.tree.heading("total", text="Total")

        self.tree.column("no", width=40, anchor="center")
        self.tree.column("barcode", width=120)
        self.tree.column("judul", width=380)
        self.tree.column("jumlah", width=60, anchor="center")
        self.tree.column("harga_normal", width=100, anchor="e")
        self.tree.column("diskon", width=80, anchor="center")
        self.tree.column("harga_akhir", width=100, anchor="e")
        self.tree.column("total", width=120, anchor="e")

        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Delete>", self.hapus_item_terpilih)
        self.tree.bind("<BackSpace>", self.hapus_item_terpilih)
        self.tree.bind("<Escape>", lambda event: self.ent_search.focus_force())

        self.root.bind("<F5>", lambda event: self.buka_popup_pembayaran())

        # =========================================================================
        # PANEL BOTTOM: TOTAL BELANJA, TOTAL QTY & TOMBOL UTAMA
        # =========================================================================
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=10, pady=10)

        info_summary_frame = ttk.Frame(bottom_frame)
        info_summary_frame.pack(side="top", fill="x", pady=5)

        self.lbl_total_qty = ttk.Label(
            info_summary_frame, 
            text="TOTAL ITEM / QTY: 0 Pcs", 
            font=("Arial", 12, "bold"), 
            foreground="darkgreen"
        )
        self.lbl_total_qty.pack(side="left", padx=10)

        self.lbl_total = ttk.Label(
            info_summary_frame, 
            text="TOTAL BELANJA: Rp. 0", 
            font=("Arial", 16, "bold"), 
            foreground="blue"
        )
        self.lbl_total.pack(side="right", padx=10)

        btn_row = ttk.Frame(bottom_frame)
        btn_row.pack(side="bottom", fill="x", pady=5)

        btn_delete_item = ttk.Button(btn_row, text="Hapus Item Selected (Del)", command=self.hapus_item_terpilih)
        btn_delete_item.pack(side="left", padx=5)

        btn_reset = ttk.Button(btn_row, text="Reset Transaksi Saat Ini", command=self.reset_transaksi)
        btn_reset.pack(side="right", padx=5)

        btn_finish = ttk.Button(btn_row, text="SELESAI & SIMPAN TRANSAKSI (F5)", command=self.buka_popup_pembayaran)
        btn_finish.pack(side="right", padx=5)

    def cek_dan_buat_header_excel(self, ws):
        val_a1 = ws["A1"].value
        val_g1 = ws["G1"].value

        if str(val_a1).strip().lower() not in ["no transaksi", "no"]:
            thin_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )
            
            header_font = Font(name="Calibri", size=11, bold=True)
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

            single_cols = [
                ("A", "No Transaksi"),
                ("B", "Kode Produk"),
                ("C", "Judul"),
                ("D", "Jumlah"),
                ("E", "Harga satuan"),
                ("F", "Total Per Produk"),
                ("I", "Diskon"),
                ("J", "Member"),
                ("K", "Waktu")
            ]

            for col, text in single_cols:
                ws.merge_cells(f"{col}1:{col}2")
                cell = ws[f"{col}1"]
                cell.value = text
                cell.font = header_font
                cell.alignment = align_center
                
                ws[f"{col}1"].border = thin_border
                ws[f"{col}2"].border = thin_border

            ws.merge_cells("G1:H1")
            cell_g1 = ws["G1"]
            cell_g1.value = "Metode"
            cell_g1.font = header_font
            cell_g1.alignment = align_center

            ws["G2"].value = "TUNAI"
            ws["G2"].font = header_font
            ws["G2"].alignment = align_center

            ws["H2"].value = "NON TUNAI"
            ws["H2"].font = header_font
            ws["H2"].alignment = align_center

            for r in [1, 2]:
                for c in ["G", "H"]:
                    ws[f"{c}{r}"].border = thin_border

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
                        last_trx = int(val)
                        break
                    except ValueError:
                        continue

            self.no_trx_counter = last_trx + 1 if last_trx > 0 else 1
        except Exception as e:
            print("Gagal membaca nomor transaksi terakhir:", e)
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
                self.lbl_status.config(text=f"[ 🟢 ONLINE Cloud: {len(self.df_master):,} Produk ]", foreground="green")
            else:
                self.lbl_status.config(text=f"[ 🟠 OFFLINE Cache: {len(self.df_master):,} Produk ]", foreground="orange")

        self.root.after(0, update_label)

    def pilih_target_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.target_file_path = file_path
            self.lbl_target_path.config(text=os.path.basename(file_path), foreground="green")
            
            try:
                self.wb_target = openpyxl.load_workbook(file_path)
                sheets = self.wb_target.sheetnames
                self.combo_sheet['values'] = sheets
                if sheets:
                    self.combo_sheet.current(0)
                self.hitung_nomor_transaksi_berikutnya()
            except Exception as e:
                messagebox.showerror("Error Excel", f"Gagal membaca file target:\n{str(e)}")

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
            messagebox.showwarning(
                "File Output Belum Dipilih", 
                "Harap pilih 'File Target Output' dan 'Sheet Target' terlebih dahulu di bagian atas sebelum melakukan transaksi!"
            )
            return

        match_barcode = self.df_master[self.df_master['barcode'].astype(str) == raw_input]

        if not match_barcode.empty:
            item = match_barcode.iloc[0]
            self.buka_popup_input_produk(item)
            self.ent_search.delete(0, tk.END)
            return

        match_manual = self.df_master[
            self.df_master['barcode'].astype(str).str.contains(raw_input, case=False, na=False) |
            self.df_master['judul'].astype(str).str.contains(raw_input, case=False, na=False)
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

        ttk.Label(
            popup, 
            text="Gunakan Panah Atas/Bawah lalu tekan ENTER untuk memilih (ESC untuk batal):", 
            font=("Arial", 9, "bold")
        ).pack(pady=8)

        frame_list = ttk.Frame(popup)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("barcode", "judul", "harga")
        tree_select = ttk.Treeview(frame_list, columns=cols, show="headings", selectmode="browse")

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
            tree_select.insert("", "end", values=(
                row['barcode'],
                row['judul'],
                f"Rp. {harga_val:,.0f}"
            ))

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

        btn_select = ttk.Button(popup, text="Pilih Produk Ini (ENTER)", command=pilih_item_terpilih)
        btn_select.pack(pady=8)

    def buka_popup_input_produk(self, item):
        popup = tk.Toplevel(self.root)
        popup.title("Input Jumlah & Diskon")
        self.center_popup(popup, 400, 260)
        popup.transient(self.root)
        popup.grab_set()

        popup.bind("<Escape>", lambda e: (popup.destroy(), self.ent_search.focus_force()))

        ttk.Label(popup, text=f"Produk: {item['judul']}", font=("Arial", 10, "bold"), wraplength=380).pack(pady=(10, 2))
        
        harga_norm = float(item['harga'])
        ttk.Label(popup, text=f"Harga Normal: Rp. {harga_norm:,.0f}", foreground="gray").pack(pady=(0, 10))

        row_qty = ttk.Frame(popup)
        row_qty.pack(pady=5)
        ttk.Label(row_qty, text="Jumlah (Qty):").pack(side="left", padx=5)
        ent_qty = ttk.Entry(row_qty, width=10, font=("Arial", 10, "bold"))
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
            except:
                messagebox.showwarning("Peringatan", "Jumlah Qty harus angka bulat positif!", parent=popup)
                return

            custom_val = ent_custom_diskon.get().strip()
            if custom_val != "0" and custom_val != "":
                try:
                    diskon_percent = float(custom_val)
                except:
                    diskon_percent = float(diskon_var.get())
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
                self.cart.append({
                    "barcode": barcode_str,
                    "judul": str(item['judul']),
                    "jumlah": added_qty,
                    "harga_normal": harga_norm,
                    "diskon": diskon_percent,
                    "harga_akhir": harga_akhir,
                    "total": total
                })

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
            diskon_str = f"{int(item['diskon'])}%" if item['diskon'].is_integer() else f"{item['diskon']}%"
            
            self.tree.insert("", "end", values=(
                idx,
                item['barcode'],
                item['judul'],
                item['jumlah'],
                f"Rp. {item['harga_normal']:,.0f}",
                diskon_str,
                f"Rp. {item['harga_akhir']:,.0f}",
                f"Rp. {item['total']:,.0f}"
            ))
            grand_total += item['total']
            total_qty += item['jumlah']

        self.lbl_total.config(text=f"TOTAL BELANJA: Rp. {grand_total:,.0f}")
        self.lbl_total_qty.config(text=f"TOTAL ITEM / QTY: {total_qty:,} Pcs")

    def reset_transaksi(self):
        if messagebox.askyesno("Konfirmasi", "Bersihkan keranjang belanja saat ini?"):
            self.cart.clear()
            self.update_tabel_keranjang()
            self.ent_search.focus_force()

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

        ttk.Label(popup, text="Total Transaksi:", font=("Arial", 11, "bold")).pack(pady=(10, 2))
        ttk.Label(popup, text=f"Rp. {grand_total:,.0f}", font=("Arial", 16, "bold"), foreground="blue").pack(pady=(0, 5))

        ada_diskon = any(item['diskon'] > 0 for item in self.cart)

        frame_member_pay = ttk.LabelFrame(popup, text=" Catatan Member ")
        frame_member_pay.pack(fill="x", padx=30, pady=5)

        ent_catatan_pay = ttk.Entry(frame_member_pay, font=("Arial", 10))
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
        ent_bayar = ttk.Entry(frame_tunai_input, font=("Arial", 11, "bold"))
        ent_bayar.pack(fill="x")

        lbl_kembalian = ttk.Label(frame_tunai_input, text="Kembalian: Rp. 0", font=("Arial", 10, "bold"), foreground="green")
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

        ttk.Radiobutton(
            row_radio, text="TUNAI", variable=pay_var, value="TUNAI", command=toggle_pembayaran
        ).pack(side="left", padx=10)
        
        ttk.Radiobutton(
            row_radio, text="NON TUNAI", variable=pay_var, value="NON TUNAI", command=toggle_pembayaran
        ).pack(side="left", padx=10)

        toggle_pembayaran()

        def simpan_dan_proses(event=None):
            catatan_val = ""
            if ada_diskon:
                catatan_val = ent_catatan_pay.get().strip()
                if not catatan_val:
                    messagebox.showwarning("Peringatan", "Catatan Member wajib diisi karena ada produk yang mendapat diskon!", parent=popup)
                    ent_catatan_pay.focus_force()
                    return

            metode = pay_var.get()
            tunai_val = 0
            nontunai_val = 0

            if metode == "TUNAI":
                raw_bayar = ent_bayar.get().replace(".", "").replace(",", "").strip()
                if not raw_bayar.isdigit() or float(raw_bayar) < grand_total:
                    messagebox.showwarning("Peringatan", "Nominal pembayaran Tunai kurang dari Total Transaksi!", parent=popup)
                    return
                tunai_val = grand_total
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

                thin_border = Border(
                    left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000')
                )

                c_no_trx = "A"
                c_kode = "B"
                c_judul = "C"
                c_qty = "D"
                c_harga = "E"
                c_tot_item = "F"
                c_tunai = "G"
                c_nontunai = "H"
                c_diskon = "I"
                c_member = "J"
                c_time = "K"

                all_used_cols = [c_no_trx, c_kode, c_judul, c_qty, c_harga, c_tot_item, c_tunai, c_nontunai, c_diskon, c_member, c_time]

                for idx_item, item in enumerate(self.cart):
                    max_r = ws.max_row + 1

                    if idx_item == 0:
                        ws[f"{c_no_trx}{max_r}"] = no_trx_num

                    ws[f"{c_kode}{max_r}"] = item['barcode']
                    ws[f"{c_judul}{max_r}"] = item['judul']
                    ws[f"{c_qty}{max_r}"] = item['jumlah']
                    
                    cell_harga = ws[f"{c_harga}{max_r}"]
                    cell_harga.value = item['harga_akhir']
                    cell_harga.number_format = currency_format

                    cell_tot_item = ws[f"{c_tot_item}{max_r}"]
                    cell_tot_item.value = item['total']
                    cell_tot_item.number_format = currency_format

                    if idx_item == 0:
                        if metode == "TUNAI":
                            cell_tunai = ws[f"{c_tunai}{max_r}"]
                            cell_tunai.value = tunai_val
                            cell_tunai.number_format = currency_format
                        else:
                            cell_nontunai = ws[f"{c_nontunai}{max_r}"]
                            cell_nontunai.value = nontunai_val
                            cell_nontunai.number_format = currency_format

                    ws[f"{c_diskon}{max_r}"] = item['diskon']
                    ws[f"{c_member}{max_r}"] = catatan_val if ada_diskon else ""

                    if idx_item == 0:
                        ws[f"{c_time}{max_r}"] = timestamp_str

                    for col_l in all_used_cols:
                        if col_l:
                            ws[f"{col_l}{max_r}"].border = thin_border

                self.wb_target.save(self.target_file_path)

                messagebox.showinfo("Sukses", f"Transaksi No. {no_trx_num} berhasil disimpan ke Sheet '{sheet_name}'!")
                
                self.cart.clear()
                self.update_tabel_keranjang()
                self.hitung_nomor_transaksi_berikutnya()
                popup.destroy()
                self.ent_search.focus_force()

            except Exception as e:
                messagebox.showerror("Error Simpan", f"Gagal menyimpan transaksi:\n{str(e)}", parent=popup)

        popup.bind("<Return>", simpan_dan_proses)

        btn_save = ttk.Button(popup, text="SIMPAN & PROSES (ENTER)", command=simpan_dan_proses)
        btn_save.pack(side="bottom", pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = KasirApp(root)
    root.mainloop()
