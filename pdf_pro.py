import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
import fitz  # PyMuPDF
import os
import threading
import queue

PAPER_SIZES = {
    "A4": (595, 842), "Letter": (612, 792), "Legal": (612, 1008),
    "A3": (842, 1191), "Ukuran Asli Gambar": None
}

class PDFEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 PDF Editor Pro ")
        self.root.geometry("1000x900")
        self.root.minsize(800, 600)
        self.pdf_document = None
        self.file_path = None
        self.previews = []
        self.preview_frames = []
        self.selected_page_index = None
        
        self.image_queue = queue.Queue()
        self.save_queue = queue.Queue() # Untuk status penyimpanan
        
        self.loading_frame = None
        self.saving_dialog = None # Untuk dialog 'Menyimpan...'
        self.is_loading = False
        
        self.root.configure(bg="#f8f9fa")
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "iconPDF.ico")
            if os.path.exists(icon_path): self.root.iconbitmap(icon_path)
        except Exception as e: print(f"Gagal mengatur ikon: {e}")
        self.setup_ui()
        self.display_previews()

    # --- Bagian Setup UI (Tidak Ada Perubahan) ---
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=X, pady=(0, 20))
        title_label = ttk.Label(header_frame, text="📄 PDF Editor Pro", font=("Segoe UI", 24, "bold"), foreground="#2c3e50")
        title_label.pack(side=LEFT)
        subtitle_label = ttk.Label(header_frame, text="Edit, gabungkan, dan kelola PDF", font=("Segoe UI", 11), foreground="#7f8c8d")
        subtitle_label.pack(side=LEFT, padx=(20, 0), pady=(8, 0))
        toolbar_frame = ttk.LabelFrame(main_frame, text="🛠️ Toolbar", padding=15)
        toolbar_frame.pack(fill=X, pady=(0, 15))
        file_ops_frame = ttk.Frame(toolbar_frame)
        file_ops_frame.pack(fill=X, pady=(0, 10))
        self.btn_open = ttk.Button(file_ops_frame, text="📂 Buka PDF", command=self.open_pdf, bootstyle="primary", width=15)
        self.btn_open.pack(side=LEFT, padx=(0, 8))
        self.btn_new = ttk.Button(file_ops_frame, text="📄 PDF Baru", command=self.new_empty_pdf, bootstyle="info", width=15)
        self.btn_new.pack(side=LEFT, padx=8)
        self.btn_save = ttk.Button(file_ops_frame, text="💾 Simpan Sebagai", state=DISABLED, command=self.save_pdf, bootstyle="success", width=15)
        self.btn_save.pack(side=RIGHT)
        page_ops_frame = ttk.Frame(toolbar_frame)
        page_ops_frame.pack(fill=X)
        self.btn_add = ttk.Button(page_ops_frame, text="➕ Tambah Halaman", state=DISABLED, command=self.add_pages, bootstyle="info-outline", width=18)
        self.btn_add.pack(side=LEFT, padx=(0, 8))
        self.btn_delete = ttk.Button(page_ops_frame, text="🗑️ Hapus Halaman", state=DISABLED, command=self.delete_page, bootstyle="danger-outline", width=18)
        self.btn_delete.pack(side=LEFT, padx=8)
        self.btn_delete_range = ttk.Button(page_ops_frame, text="🗂️ Hapus Rentang", state=DISABLED, command=self.delete_page_range, bootstyle="warning-outline", width=18)
        self.btn_delete_range.pack(side=LEFT, padx=8)
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=X, pady=(0, 15))
        status_container = ttk.LabelFrame(status_frame, text="📊 Status Dokumen", padding=10)
        status_container.pack(fill=X)
        self.info_label = ttk.Label(status_container, text="🎯 Buka file PDF atau buat PDF baru", font=("Segoe UI", 11), foreground="#34495e")
        self.info_label.pack(fill=X)
        ttk.Separator(main_frame, orient=HORIZONTAL).pack(fill=X, pady=(0, 15))
        self.setup_pdf_mode(main_frame)
        
    def setup_pdf_mode(self, parent):
        self.pdf_frame = ttk.LabelFrame(parent, text="🔍 Pratinjau Halaman PDF", padding=15)
        self.pdf_frame.pack(fill=BOTH, expand=True)
        info_panel = ttk.Frame(self.pdf_frame)
        info_panel.pack(fill=X, pady=(0, 10))
        preview_info = ttk.Label(info_panel, text="💡 Klik halaman untuk memilih. Gunakan tombol untuk memutar.", font=("Segoe UI", 9), foreground="#7f8c8d")
        preview_info.pack(side=LEFT)
        canvas_container = ttk.Frame(self.pdf_frame)
        canvas_container.pack(fill=BOTH, expand=True)
        self.canvas = tk.Canvas(canvas_container, highlightthickness=0, bg="#ecf0f1", relief="flat")
        self.scrollbar = ttk.Scrollbar(canvas_container, orient=VERTICAL, command=self.canvas.yview, bootstyle="round-primary")
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="n")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)

    # --- Bagian Loading dan Helper (Ada Penambahan) ---
    def _show_saving_indicator(self):
        """Menampilkan dialog modal 'Menyimpan...'."""
        self.saving_dialog = tk.Toplevel(self.root)
        self.saving_dialog.title("Menyimpan")
        self.saving_dialog.geometry("300x120")
        self.saving_dialog.resizable(False, False)
        self.saving_dialog.transient(self.root)
        self.saving_dialog.grab_set()

        main_frame = ttk.Frame(self.saving_dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        ttk.Label(main_frame, text="📄 Menyimpan PDF...", font=("Segoe UI", 11)).pack(pady=(0, 10))
        progress = ttk.Progressbar(main_frame, mode='indeterminate', length=250)
        progress.pack()
        progress.start(10)

        self._center_dialog(self.saving_dialog)

    def _hide_saving_indicator(self):
        """Menyembunyikan dialog 'Menyimpan...'."""
        if self.saving_dialog:
            self.saving_dialog.destroy()
            self.saving_dialog = None

    def show_loading_indicator(self, total_pages):
        self.hide_loading_indicator()
        self.is_loading = True
        self.loading_frame = ttk.Frame(self.pdf_frame, bootstyle="light")
        self.loading_frame.place(relx=0.5, rely=0.5, anchor=CENTER)
        ttk.Label(self.loading_frame, text="⏳ Memuat Halaman...", font=("Segoe UI", 14, "bold")).pack(pady=(15,10))
        self.loading_progress = ttk.Progressbar(self.loading_frame, mode='determinate', length=300, maximum=total_pages, bootstyle="striped-primary")
        self.loading_progress.pack(pady=5, padx=20)
        self.loading_label = ttk.Label(self.loading_frame, text=f"Mempersiapkan...", font=("Segoe UI", 10))
        self.loading_label.pack(pady=(5, 15))

    def hide_loading_indicator(self):
        self.is_loading = False
        if self.loading_frame:
            self.loading_frame.destroy()
            self.loading_frame = None

    def _center_dialog(self, dialog):
        dialog.update_idletasks()
        root_x, root_y = self.root.winfo_x(), self.root.winfo_y()
        root_w, root_h = self.root.winfo_width(), self.root.winfo_height()
        dialog_w, dialog_h = dialog.winfo_width(), dialog.winfo_height()
        pos_x = root_x + (root_w // 2) - (dialog_w // 2)
        pos_y = root_y + (root_h // 2) - (dialog_h // 2)
        dialog.geometry(f"+{pos_x}+{pos_y}")
        dialog.transient(self.root)
        dialog.grab_set()

    def _on_mousewheel(self, event):
        if not self.is_loading:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            
    # --- Bagian Pratinjau dan Manipulasi UI ---
    def display_previews(self):
        self.is_loading = False
        self.hide_loading_indicator()
        for widget in self.scrollable_frame.winfo_children(): widget.destroy()
        self.previews.clear()
        self.preview_frames.clear()
        if not self.pdf_document:
            empty_frame = ttk.Frame(self.scrollable_frame)
            empty_frame.pack(expand=True, fill=BOTH, pady=50)
            ttk.Label(empty_frame, text="📁\n\nTidak ada dokumen PDF yang terbuka", font=("Segoe UI", 14), foreground="#bdc3c7", justify=CENTER).pack(expand=True)
            self.update_info_label()
            return
        page_count = len(self.pdf_document)
        self.update_info_label()
        if page_count > 0:
            self.show_loading_indicator(page_count)
            threading.Thread(target=self._load_pages_worker, daemon=True).start()
            self.root.after(100, self._process_image_queue)
        self.selected_page_index = None
        self.btn_delete.config(state=DISABLED)
        self.btn_delete_range.config(state=NORMAL if page_count > 0 else DISABLED)

    def _load_pages_worker(self):
        if not self.pdf_document: return
        for i, page in enumerate(self.pdf_document):
            if not self.is_loading: break
            pix = page.get_pixmap(dpi=96)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail((280, 380))
            self.image_queue.put((i, img))
        self.image_queue.put(None)
        
    def _process_image_queue(self):
        try:
            item = self.image_queue.get_nowait()
            if item is None:
                self.hide_loading_indicator()
                return
            page_index, img_data = item
            new_widget_data = self._create_preview_widget(page=None, index=page_index, img_data=img_data)
            new_widget_data['frame'].pack(pady=12, padx=15, fill=X)
            self.previews.append(new_widget_data['photo'])
            self.preview_frames.append({'frame': new_widget_data['frame'], 'img_label': new_widget_data['img_label']})
            if self.loading_frame:
                self.loading_progress['value'] = page_index + 1
                self.loading_label.config(text=f"Memuat halaman {page_index + 1} dari {len(self.pdf_document)}...")
        except queue.Empty:
            pass
        if self.is_loading:
            self.root.after(50, self._process_image_queue)

    def _find_index_from_widget(self, frame_widget):
        return next((i for i, item in enumerate(self.preview_frames) if item['frame'] == frame_widget), None)

    def handle_page_click(self, frame_widget):
        index = self._find_index_from_widget(frame_widget)
        if index is not None: self.select_page(index)
    
    def handle_rotate(self, frame_widget, angle):
        index = self._find_index_from_widget(frame_widget)
        if index is not None: self.rotate_page(index, angle)

    def _create_preview_widget(self, page, index, img_data=None):
        if img_data is None:
            pix = page.get_pixmap(dpi=96)
            img_data = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_data.thumbnail((280, 380))
        photo = ImageTk.PhotoImage(img_data)
        page_frame = ttk.Labelframe(self.scrollable_frame, text=f"📃 Halaman {index + 1}", bootstyle="secondary", padding=12)
        img_container = ttk.Frame(page_frame)
        img_container.pack()
        img_label = ttk.Label(img_container, image=photo, cursor="hand2")
        img_label.pack(padx=8, pady=8)
        action_frame = ttk.Frame(page_frame)
        action_frame.pack(fill=X, pady=(10, 0))
        btn_rotate_left = ttk.Button(action_frame, text="↺ Putar Kiri", bootstyle="secondary-outline", command=lambda f=page_frame: self.handle_rotate(f, -90))
        btn_rotate_left.pack(side=LEFT, expand=True, padx=5)
        btn_rotate_right = ttk.Button(action_frame, text="Putar Kanan ↻", bootstyle="secondary-outline", command=lambda f=page_frame: self.handle_rotate(f, 90))
        btn_rotate_right.pack(side=RIGHT, expand=True, padx=5)
        def create_hover_bindings(frame):
            def on_enter(e):
                index = self._find_index_from_widget(frame)
                if index != self.selected_page_index: frame.configure(bootstyle="info")
            def on_leave(e):
                index = self._find_index_from_widget(frame)
                if index != self.selected_page_index: frame.configure(bootstyle="secondary")
            frame.bind("<Enter>", on_enter)
            frame.bind("<Leave>", on_leave)
            for widget in [frame, img_container, img_label]:
                widget.bind("<Button-1>", lambda e, f=frame: self.handle_page_click(f))
        create_hover_bindings(page_frame)
        return {'frame': page_frame, 'img_label': img_label, 'photo': photo}

    def _update_single_preview(self, index):
        try:
            page = self.pdf_document[index]
            pix = page.get_pixmap(dpi=96)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail((280, 380))
            new_photo = ImageTk.PhotoImage(img)
            target_label = self.preview_frames[index]['img_label']
            target_label.config(image=new_photo)
            self.previews[index] = new_photo
        except Exception as e:
            self.display_previews()

    def _remove_preview_at(self, index):
        if not (0 <= index < len(self.preview_frames)): return
        self.preview_frames[index]['frame'].destroy()
        del self.preview_frames[index]
        del self.previews[index]
        self._renumber_previews_after(index)
        self.selected_page_index = None
        self.btn_delete.config(state=DISABLED)
        self.update_info_label()

    def _renumber_previews_after(self, start_index):
        for i in range(start_index, len(self.preview_frames)):
            new_page_num = i + 1
            self.preview_frames[i]['frame'].config(text=f"📃 Halaman {new_page_num}")

    def update_info_label(self, custom_message=None):
        if custom_message:
            self.info_label.config(text=custom_message)
            return
        if self.pdf_document:
            page_count = len(self.pdf_document)
            file_name = os.path.basename(self.file_path) if self.file_path else "PDF Baru"
            if self.selected_page_index is not None:
                self.info_label.config(text=f"📁 {file_name}  |  📄 {page_count} hal.  |  ✅ Hal. {self.selected_page_index + 1} dipilih")
            else:
                self.info_label.config(text=f"📁 {file_name}  |  📄 {page_count} hal.  |  ✨ Klik halaman untuk memilih")
        else:
            self.info_label.config(text="🎯 Buka file PDF atau buat PDF baru")

    def select_page(self, index):
        if self.selected_page_index is not None:
            try:
                self.preview_frames[self.selected_page_index]['frame'].config(bootstyle="secondary")
            except (IndexError, KeyError): pass
        self.selected_page_index = index
        self.preview_frames[index]['frame'].config(bootstyle="primary")
        self.btn_delete.config(state=NORMAL)
        self.update_info_label()

    # --- Bagian Aksi Utama (File, Halaman, dll) ---
    def open_pdf(self):
        if self.is_loading:
            self.is_loading = False
            messagebox.showinfo("Proses Dibatalkan", "Proses pemuatan PDF sebelumnya telah dihentikan.", parent=self.root)
        path = filedialog.askopenfilename(title="Pilih File PDF", filetypes=[("PDF Files", "*.pdf")])
        if not path: return
        try:
            if self.pdf_document: self.pdf_document.close()
            self.file_path = path
            self.pdf_document = fitz.open(self.file_path)
            self.update_ui_after_load()
            self.display_previews()
            messagebox.showinfo("✅ Berhasil", f"PDF berhasil dibuka: {os.path.basename(path)}", parent=self.root)
        except Exception as e:
            messagebox.showerror("❌ Error", f"Gagal membuka file PDF:\n\n{str(e)}", parent=self.root)
            self.reset_state()

    def new_empty_pdf(self):
        if self.pdf_document and messagebox.askyesno("🤔 Konfirmasi", "Tutup PDF saat ini dan buat baru?", parent=self.root):
            self.reset_state()
        try:
            self.pdf_document = fitz.open()
            self.file_path = "PDF_Baru.pdf"
            self.update_ui_after_load()
            self.display_previews()
            self.update_info_label("✨ PDF kosong baru telah dibuat.")
        except Exception as e:
            messagebox.showerror("❌ Error", f"Gagal membuat PDF baru:\n\n{str(e)}", parent=self.root)

    def add_pages(self):
        choice = messagebox.askyesnocancel("📎 Pilih Sumber", "Tambah halaman dari:\n\n- [Ya]   = File PDF lain\n- [Tidak] = File Gambar", parent=self.root)
        if choice is None: return
        elif choice: self.add_pages_from_pdf()
        else: self.add_pages_from_images()

    def add_pages_from_pdf(self):
        add_path = filedialog.askopenfilename(title="📂 Pilih PDF untuk ditambahkan", filetypes=[("PDF Files", "*.pdf")])
        if not add_path: return
        try:
            with fitz.open(add_path) as pdf_to_add:
                pages_added_count = len(pdf_to_add)
                insert_at = simpledialog.askinteger("📍 Posisi Penyisipan", f"Sisipkan setelah halaman ke:\n(0 = di awal, {len(self.pdf_document)} = di akhir)", minvalue=0, maxvalue=len(self.pdf_document), parent=self.root)
                if insert_at is None: return
                self.pdf_document.insert_pdf(pdf_to_add, start_at=insert_at)
                sibling_widget = self.preview_frames[insert_at]['frame'] if insert_at < len(self.preview_frames) else None
                for i in range(pages_added_count):
                    current_index = insert_at + i
                    page = self.pdf_document[current_index]
                    new_widget_data = self._create_preview_widget(page, current_index)
                    if sibling_widget:
                        new_widget_data['frame'].pack(before=sibling_widget)
                    else:
                        new_widget_data['frame'].pack(pady=12, padx=15, fill=X)
                    self.previews.insert(current_index, new_widget_data['photo'])
                    self.preview_frames.insert(current_index, {'frame': new_widget_data['frame'], 'img_label': new_widget_data['img_label']})
                self._renumber_previews_after(insert_at)
                self.update_info_label()
            messagebox.showinfo("✅ Berhasil", f"🎉 {pages_added_count} halaman berhasil ditambahkan!", parent=self.root)
        except Exception as e:
            messagebox.showerror("❌ Error", f"Gagal menambahkan halaman:\n\n{str(e)}", parent=self.root)

    def _ask_paper_size_for_images(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("📏 Pilih Ukuran Halaman")
        dialog.resizable(False, False)
        dialog.configure(bg="#f8f9fa")
        result = tk.StringVar(value="A4")
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main_frame, text="📏 Pilih Ukuran Kertas", font=("Segoe UI", 16, "bold"), foreground="#2c3e50").pack(anchor='w')
        ttk.Label(main_frame, text="Tentukan ukuran halaman untuk gambar:", font=("Segoe UI", 10), foreground="#7f8c8d").pack(anchor='w', pady=(5, 0))
        option_frame = ttk.LabelFrame(main_frame, text="Pilihan Ukuran", padding=15)
        option_frame.pack(fill=X, pady=15)
        for size_name in PAPER_SIZES.keys():
            size_info = f" ({PAPER_SIZES[size_name][0]}x{PAPER_SIZES[size_name][1]} pt)" if PAPER_SIZES[size_name] else " (Sesuai ukuran asli)"
            ttk.Radiobutton(option_frame, text=size_name + size_info, variable=result, value=size_name, bootstyle="primary").pack(anchor='w', pady=3)
        final_choice = [None]
        def on_ok():
            final_choice[0] = result.get()
            dialog.destroy()
        btn_frame = ttk.Frame(main_frame, padding=(0, 10, 0, 0))
        btn_frame.pack(fill=X, side=BOTTOM)
        ttk.Button(btn_frame, text="❌ Batal", command=dialog.destroy, bootstyle="secondary-outline", width=12).pack(side=RIGHT)
        ttk.Button(btn_frame, text="✅ OK", command=on_ok, bootstyle="primary", width=12).pack(side=RIGHT, padx=(0, 10))
        self._center_dialog(dialog)
        self.root.wait_window(dialog)
        return final_choice[0]

    def add_pages_from_images(self):
        selected_paper_size_name = self._ask_paper_size_for_images()
        if not selected_paper_size_name: return
        image_paths = filedialog.askopenfilenames(title="🖼️ Pilih Gambar", filetypes=[("Image Files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff")])
        if not image_paths: return
        try:
            insert_at = simpledialog.askinteger("📍 Posisi Penyisipan", f"Sisipkan setelah halaman ke:\n(0 = di awal, {len(self.pdf_document)} = di akhir)", minvalue=0, maxvalue=len(self.pdf_document), parent=self.root)
            if insert_at is None: return
            page_dim = PAPER_SIZES[selected_paper_size_name]
            sibling_widget = self.preview_frames[insert_at]['frame'] if insert_at < len(self.preview_frames) else None
            for i, img_path in enumerate(image_paths):
                current_index = insert_at + i
                with Image.open(img_path) as img:
                    if img.mode in ('RGBA', 'LA'):
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, (0, 0), img)
                        processed_img = background
                    else:
                        processed_img = img.convert('RGB')
                    if page_dim:
                        pw, ph = page_dim
                        page = self.pdf_document.new_page(pno=current_index, width=pw, height=ph)
                        iw, ih = processed_img.size
                        r_img, r_page = iw / ih, pw / ph
                        if r_img > r_page:
                            rw, rh = pw * 0.95, (pw * 0.95) / r_img
                        else:
                            rw, rh = (ph * 0.95) * r_img, ph * 0.95
                        x, y = (pw - rw) / 2, (ph - rh) / 2
                        img_rect = fitz.Rect(x, y, x + rw, y + rh)
                    else:
                        pw, ph = processed_img.size
                        page = self.pdf_document.new_page(pno=current_index, width=pw, height=ph)
                        img_rect = fitz.Rect(0, 0, pw, ph)
                    page.insert_image(img_rect, filename=img_path)
                    new_widget_data = self._create_preview_widget(page, current_index)
                    if sibling_widget:
                        new_widget_data['frame'].pack(before=sibling_widget)
                    else:
                        new_widget_data['frame'].pack(pady=12, padx=15, fill=X)
                    self.previews.insert(current_index, new_widget_data['photo'])
                    self.preview_frames.insert(current_index, {'frame': new_widget_data['frame'], 'img_label': new_widget_data['img_label']})
            self._renumber_previews_after(insert_at)
            self.update_info_label()
            messagebox.showinfo("✅ Berhasil", f"🎉 {len(image_paths)} gambar ditambahkan!", parent=self.root)
        except Exception as e:
            messagebox.showerror("❌ Error", f"Gagal menambahkan gambar:\n\n{str(e)}", parent=self.root)

    def rotate_page(self, page_index, angle):
        if not self.pdf_document or not (0 <= page_index < len(self.pdf_document)): return
        try:
            page = self.pdf_document[page_index]
            new_rotation = (page.rotation + angle) % 360
            page.set_rotation(new_rotation)
            self._update_single_preview(page_index)
        except Exception as e:
            messagebox.showerror("❌ Error", f"Gagal memutar halaman:\n\n{str(e)}", parent=self.root)

    def delete_page(self):
        if self.selected_page_index is None:
            messagebox.showwarning("⚠️ Peringatan", "Pilih halaman yang ingin dihapus.", parent=self.root)
            return
        page_num_to_delete = self.selected_page_index + 1
        if messagebox.askyesno("🗑️ Konfirmasi Hapus", f"Yakin ingin menghapus Halaman {page_num_to_delete}?", parent=self.root):
            try:
                index_to_delete = self.selected_page_index
                self.pdf_document.delete_page(index_to_delete)
                self._remove_preview_at(index_to_delete)
                messagebox.showinfo("✅ Berhasil", f"Halaman {page_num_to_delete} berhasil dihapus.", parent=self.root)
            except Exception as e:
                messagebox.showerror("❌ Error", f"Gagal menghapus halaman:\n\n{str(e)}", parent=self.root)

    def delete_page_range(self):
        if not self.pdf_document or len(self.pdf_document) < 2:
            messagebox.showwarning("⚠️ Peringatan", "Perlu minimal 2 halaman untuk menghapus rentang.", parent=self.root)
            return
        dialog_result = self._ask_delete_range()
        if dialog_result is None: return
        start_page, end_page = dialog_result
        start_index, end_index = start_page - 1, end_page - 1
        num_to_delete = (end_index - start_index) + 1
        if messagebox.askyesno("🗑️ Konfirmasi", f"Hapus halaman {start_page} hingga {end_page} ({num_to_delete} halaman)?", parent=self.root):
            try:
                self.pdf_document.delete_pages(range(start_index, end_index + 1))
                for i in range(end_index, start_index - 1, -1):
                    self.preview_frames[i]['frame'].destroy()
                    del self.preview_frames[i]
                    del self.previews[i]
                self._renumber_previews_after(start_index)
                self.selected_page_index = None
                self.btn_delete.config(state=DISABLED)
                self.update_info_label()
                messagebox.showinfo("✅ Berhasil", f"🎉 {num_to_delete} halaman dihapus!", parent=self.root)
            except Exception as e:
                messagebox.showerror("❌ Error", f"Gagal menghapus rentang:\n\n{str(e)}", parent=self.root)

    def _ask_delete_range(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("🗂️ Hapus Rentang Halaman")
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        total_pages = len(self.pdf_document)
        ttk.Label(main_frame, text="🗂️ Hapus Rentang Halaman", font=("Segoe UI", 16, "bold")).pack(anchor=W)
        ttk.Label(main_frame, text=f"📊 Total halaman: {total_pages}", font=("Segoe UI", 10)).pack(anchor=W, pady=(0, 15))
        input_frame = ttk.LabelFrame(main_frame, text="Pilih Rentang", padding=15)
        input_frame.pack(fill=X, pady=(0, 15))
        range_grid_frame = ttk.Frame(input_frame)
        range_grid_frame.pack(pady=5)
        ttk.Label(range_grid_frame, text="Dari:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky=W)
        entry_from = ttk.Entry(range_grid_frame, width=8, font=("Segoe UI", 12))
        entry_from.grid(row=0, column=1, padx=5, pady=5)
        entry_from.insert(0, "1")
        ttk.Label(range_grid_frame, text="Hingga:", font=("Segoe UI", 10)).grid(row=0, column=2, sticky=W, padx=(15, 0))
        entry_to = ttk.Entry(range_grid_frame, width=8, font=("Segoe UI", 12))
        entry_to.grid(row=0, column=3, padx=5, pady=5)
        entry_to.insert(0, str(total_pages))
        result = [None]
        def on_ok():
            try:
                start, end = int(entry_from.get()), int(entry_to.get())
                if not (1 <= start <= end <= total_pages and (start, end) != (1, total_pages)):
                    raise ValueError
                result[0] = (start, end)
                dialog.destroy()
            except ValueError:
                messagebox.showerror("❌ Input Tidak Valid", f"Rentang tidak valid. Pastikan Awal ≤ Akhir, antara 1-{total_pages}, dan tidak menghapus semua halaman.", parent=dialog)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=X, side=BOTTOM)
        ttk.Button(btn_frame, text="❌ Batal", command=dialog.destroy, bootstyle="secondary-outline").pack(side=RIGHT)
        ttk.Button(btn_frame, text="🗑️ Hapus", command=on_ok, bootstyle="danger").pack(side=RIGHT, padx=(0, 10))
        self._center_dialog(dialog)
        self.root.wait_window(dialog)
        return result[0]
        
    def save_pdf(self):
        if not self.pdf_document: return
        dialog = tk.Toplevel(self.root)
        dialog.title("💾 Opsi Penyimpanan")
        self._center_dialog(dialog)
        save_option = tk.StringVar(value="all")
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        ttk.Label(main_frame, text="💾 Opsi Penyimpanan", font=("Segoe UI", 16, "bold")).pack(anchor=W)
        def toggle():
            state = NORMAL if save_option.get() == "range" else DISABLED
            entry_from.config(state=state)
            entry_to.config(state=state)
        options_frame = ttk.LabelFrame(main_frame, text="Pilihan", padding=15)
        options_frame.pack(fill=X, pady=15)
        ttk.Radiobutton(options_frame, text=f"Simpan Semua Halaman ({len(self.pdf_document)})", variable=save_option, value="all", command=toggle).pack(anchor=W)
        ttk.Radiobutton(options_frame, text="Simpan Rentang Halaman:", variable=save_option, value="range", command=toggle).pack(anchor=W, pady=(5,0))
        range_input_frame = ttk.Frame(options_frame)
        range_input_frame.pack(fill=X, padx=(25, 0), pady=(5, 0))
        ttk.Label(range_input_frame, text="Dari:").pack(side=LEFT)
        entry_from = ttk.Entry(range_input_frame, width=6, state=DISABLED)
        entry_from.pack(side=LEFT, padx=(5, 10))
        ttk.Label(range_input_frame, text="Hingga:").pack(side=LEFT)
        entry_to = ttk.Entry(range_input_frame, width=6, state=DISABLED)
        entry_to.pack(side=LEFT, padx=5)
        def on_ok():
            start_page, end_page = None, None
            if save_option.get() == "range":
                try:
                    start_page, end_page = int(entry_from.get()), int(entry_to.get())
                    if not (1 <= start_page <= end_page <= len(self.pdf_document)): raise ValueError()
                except (ValueError, TypeError):
                    messagebox.showerror("❌ Error", f"Rentang tidak valid (1-{len(self.pdf_document)}).", parent=dialog)
                    return
            dialog.destroy()
            self._execute_save(start_page, end_page)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=X, side=BOTTOM, pady=(10, 0))
        ttk.Button(btn_frame, text="❌ Batal", command=dialog.destroy, bootstyle="secondary-outline").pack(side=RIGHT)
        ttk.Button(btn_frame, text="💾 Simpan", command=on_ok, bootstyle="success").pack(side=RIGHT, padx=(0, 10))

    def _execute_save(self, start_page=None, end_page=None):
        save_path = filedialog.asksaveasfilename(title="💾 Simpan PDF Sebagai", defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not save_path: return
        self._show_saving_indicator()
        threading.Thread(
            target=self._save_worker,
            args=(save_path, start_page, end_page),
            daemon=True
        ).start()
        self.root.after(100, self._process_save_queue)

    def _save_worker(self, save_path, start_page=None, end_page=None):
        try:
            if start_page is None:
                self.pdf_document.save(save_path, garbage=4, deflate=True, clean=True)
                pages_saved = len(self.pdf_document)
            else:
                start_index, end_index = start_page - 1, end_page - 1
                new_doc = fitz.open()
                new_doc.insert_pdf(self.pdf_document, from_page=start_index, to_page=end_index)
                new_doc.save(save_path, garbage=4, deflate=True, clean=True)
                new_doc.close()
                pages_saved = (end_index - start_index) + 1
            self.save_queue.put(('success', save_path, pages_saved))
        except Exception as e:
            self.save_queue.put(('error', e))

    def _process_save_queue(self):
        try:
            item = self.save_queue.get_nowait()
            self._hide_saving_indicator()
            status, data = item[0], item[1:]
            if status == 'success':
                save_path, pages_saved = data
                messagebox.showinfo("✅ Berhasil", f"🎉 File berhasil disimpan!\nLokasi: {os.path.basename(save_path)}\nHalaman: {pages_saved}", parent=self.root)
            else:
                error_exception = data[0]
                messagebox.showerror("❌ Error", f"Gagal menyimpan file:\n\n{str(error_exception)}", parent=self.root)
        except queue.Empty:
            self.root.after(100, self._process_save_queue)

    def update_ui_after_load(self):
        self.btn_add.config(state=NORMAL)
        self.btn_save.config(state=NORMAL)
        self.btn_delete.config(state=DISABLED)
        self.btn_delete_range.config(state=NORMAL if self.pdf_document and len(self.pdf_document) > 0 else DISABLED)
    
    def reset_state(self):
        self.is_loading = False
        if self.pdf_document: self.pdf_document.close()
        self.pdf_document, self.file_path, self.selected_page_index = None, None, None
        self.previews.clear()
        self.preview_frames.clear()
        self.btn_add.config(state=DISABLED)
        self.btn_save.config(state=DISABLED) 
        self.btn_delete.config(state=DISABLED)
        self.btn_delete_range.config(state=DISABLED)
        self.update_info_label()
        self.display_previews()

if __name__ == "__main__":
    root = ttk.Window(themename="litera")
    app = PDFEditorApp(root)
    root.mainloop()