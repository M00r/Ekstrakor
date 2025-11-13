import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from pathlib import Path
import threading
import os
from PIL import Image
import io
import tempfile
import cv2
from docx import Document
from docx.shared import Mm

# Helper functions

def is_valid_image(path):
    try:
        with Image.open(path) as img:
            img.verify()
        return True
    except:
        return False

def extract_first_frame_gif(path):
    try:
        with Image.open(path) as im:
            im.seek(0)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                im.save(tmp.name, format='PNG')
                return Path(tmp.name)
    except:
        return None

def extract_middle_frame_video(path):
    cap = cv2.VideoCapture(str(path))
    if not cap.isOpened():
        return None
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    middle_frame = frame_count // 2
    cap.set(cv2.CAP_PROP_POS_FRAMES, middle_frame)
    ret, frame = cap.read()
    cap.release()
    if ret:
        # Convert BGR to RGB
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            # Save the RGB frame as PNG after converting back to BGR for cv2.imwrite
            cv2.imwrite(tmp.name, cv2.cvtColor(frame_rgb, cv2.COLOR_RGB2BGR))
            return Path(tmp.name)
    return None

def get_doc_size_in_mb(doc, doc_number):
    temp_path = Path(f"temp_{doc_number}.docx")
    doc.save(str(temp_path))
    size = temp_path.stat().st_size / (1024 * 1024)
    temp_path.unlink()
    return size

def gather_media_files_from_folder(folder_path: Path):
    media_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            path = Path(root) / file
            if path.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp", ".gif",
                                       ".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv"}:
                media_files.append(path)
    return media_files

def process_files(selected_files, output_file_path, progress_callback):
    media_files = selected_files
    media_files.sort()

    if not media_files:
        messagebox.showinfo("Info", "Brak wybranych plików do przetwarzania.")
        progress_callback(100)
        return

    max_size_mb = 400
    doc_counter = 1
    current_doc = Document()
    section = current_doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(15)
    section.right_margin = Mm(15)
    section.top_margin = Mm(15)
    section.bottom_margin = Mm(15)
    current_doc.add_heading(f"Załącznik do protokołu - część {doc_counter}", level=1)

    total_files = len(media_files)
    for idx, media_path in enumerate(media_files):
        # Check current document size
        current_size = get_doc_size_in_mb(current_doc, doc_counter)
        if current_size >= max_size_mb:
            # Save current doc
            output_path = Path(str(output_file_path).replace('.docx', f'_part{doc_counter}.docx'))
            current_doc.save(str(output_path))
            doc_counter += 1
            # Start new document
            current_doc = Document()
            section = current_doc.sections[0]
            section.page_width = Mm(210)
            section.page_height = Mm(297)
            section.left_margin = Mm(15)
            section.right_margin = Mm(15)
            section.top_margin = Mm(15)
            section.bottom_margin = Mm(15)
            current_doc.add_heading(f"Galeria multimediów - część {doc_counter}", level=1)

        suffix = media_path.suffix.lower()
        img_to_add = None

        try:
            if suffix in {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp", ".gif"}:
                if suffix == ".gif":
                    img_to_add = extract_first_frame_gif(media_path)
                elif is_valid_image(media_path):
                    img_to_add = media_path
                else:
                    continue
            elif suffix in {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv"}:
                img_to_add = extract_middle_frame_video(media_path)
                if not img_to_add:
                    continue

            if img_to_add:
                rel_path = media_path
                current_doc.add_paragraph(str(rel_path))
                current_doc.add_picture(str(img_to_add), width=Mm(120))
                current_doc.add_paragraph()
        except:
            continue

        # Update progress
        progress = int(((idx + 1) / total_files) * 100)
        progress_callback(progress)

    # Save the last document
    output_path = Path(str(output_file_path).replace('.docx', f'_part{doc_counter}.docx'))
    current_doc.save(str(output_path))
    messagebox.showinfo("Done", f"Pliki zapisano jako: {output_path}")

# GUI Application
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Media Gallery Processor")

        self.selected_files = []
        self.selected_folders = []
        self.output_file = None

        # UI Elements
        self.btn_select_files = tk.Button(root, text="Wybierz pliki", command=self.select_files)
        self.btn_select_files.pack(pady=5)

        self.btn_select_folder = tk.Button(root, text="Wybierz folder", command=self.select_folder)
        self.btn_select_folder.pack(pady=5)

        self.lbl_files = tk.Label(root, text="Nie wybrano plików")
        self.lbl_files.pack()

        self.lbl_folders = tk.Label(root, text="Nie wybrano folderów")
        self.lbl_folders.pack()

        self.btn_select_output = tk.Button(root, text="Wybierz plik wyjściowy", command=self.select_output)
        self.btn_select_output.pack(pady=5)

        self.lbl_output = tk.Label(root, text="Nie wybrano pliku wyjściowego")
        self.lbl_output.pack()

        self.size_label = tk.Label(root, text="Rozmiar wybranych plików: 0 MB")
        self.size_label.pack(pady=5)

        self.progress = Progressbar(root, orient='horizontal', length=300, mode='determinate')
        self.progress.pack(pady=10)

        self.btn_start = tk.Button(root, text="Rozpocznij przetwarzanie", command=self.start_processing)
        self.btn_start.pack(pady=5)

        # Quote label
        self.quote_label = tk.Label(root, text="Made by Iwo Szczeciński", font=("Arial", 10, "italic"))
        self.quote_label.pack(pady=10)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Wybierz pliki",
            filetypes=[
                ("Media files", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp *.gif *.mp4 *.mov *.avi *.mkv *.wmv *.flv")
            ]
        )
        if files:
            self.selected_files = [Path(f) for f in files]
            self.lbl_files.config(text=f"Wybrano {len(self.selected_files)} plików")
            total_size = sum(p.stat().st_size for p in self.selected_files) / (1024 * 1024)
            self.size_label.config(text=f"Rozmiar wybranych plików: {total_size:.2f} MB")
        else:
            self.selected_files = []
            self.lbl_files.config(text="Nie wybrano plików")
            self.size_label.config(text="Rozmiar wybranych plików: 0 MB")

    def select_folder(self):
        folder = filedialog.askdirectory(title="Wybierz folder")
        if folder:
            folder_path = Path(folder)
            self.selected_folders.append(folder_path)
            # Gather media files from this folder
            media_files_in_folder = gather_media_files_from_folder(folder_path)
            self.selected_files.extend(media_files_in_folder)
            self.lbl_folders.config(text=f"Wybrano {len(self.selected_folders)} folderów")
            total_size = sum(p.stat().st_size for p in self.selected_files) / (1024 * 1024)
            self.size_label.config(text=f"Rozmiar wybranych plików: {total_size:.2f} MB")
        else:
            # No folder selected
            pass

    def select_output(self):
        file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if file:
            self.output_file = Path(file)
            self.lbl_output.config(text=str(self.output_file))

    def start_processing(self):
        if not self.selected_files or not self.output_file:
            messagebox.showwarning("Warning", "Wybierz pliki, folder i plik wyjściowy.")
            return
        threading.Thread(target=self.run_processing).start()

    def run_processing(self):
        self.progress['value'] = 0
        process_files(self.selected_files, self.output_file, self.update_progress)

    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
