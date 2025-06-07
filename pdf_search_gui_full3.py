import os
import sys
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import subprocess
import platform


CONFIG_DIR = os.path.join(os.getenv('APPDATA'), "DeepIndexPDF")
os.makedirs(CONFIG_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.txt")

def save_last_folder(path):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(path)
    except Exception as e:
        print("Gagal menyimpan folder:", e)

def load_last_folder():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            print("Gagal memuat folder:", e)
    return ""

def search_pdfs(folder_path, keyword):
    results = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            try:
                doc = fitz.open(file_path)
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    text = page.get_text()
                    if keyword.lower() in text.lower():
                        lines = text.split('\n')
                        for line in lines:
                            if keyword.lower() in line.lower():
                                results.append((filename, page_num + 1, line.strip(), file_path))
                doc.close()
            except Exception as e:
                results.append((filename, "Error", str(e), file_path))
    return results

def browse_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_var.set(folder)
        save_last_folder(folder)

def run_search():
    folder = folder_var.get()
    keyword = keyword_entry.get().strip()
    if not folder or not keyword:
        messagebox.showwarning("Input Salah", "Pilih folder dan masukkan kata kunci.")
        return

    results = search_pdfs(folder, keyword)
    for row in tree.get_children():
        tree.delete(row)

    for file, page, snippet, path in results:
        tree.insert("", "end", values=(file, page, snippet, path))

    if results:
        save_button.config(state=tk.NORMAL)
    else:
        messagebox.showinfo("Hasil", "Tidak ditemukan hasil yang cocok.")

def save_results():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV Files", "*.csv")])
    if file_path:
        with open(file_path, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["File", "Halaman", "Cuplikan", "Path"])
            for row_id in tree.get_children():
                writer.writerow(tree.item(row_id)["values"])
        messagebox.showinfo("Selesai", f"Hasil disimpan di:\n{file_path}")

def open_pdf_from_tree(event):
    selected = tree.selection()
    if selected:
        values = tree.item(selected[0])["values"]
        file_path = values[3]
        page_number = values[1]
        open_pdf_at_page(file_path, page_number)

def open_pdf_at_page(file_path, page_number):
    try:
        base_dir = os.path.dirname(sys.executable)
        sumatra_path = os.path.join(base_dir, "SumatraPDF", "SumatraPDF.exe")
        keyword = keyword_entry.get().strip()

        if platform.system() == "Windows" and os.path.exists(sumatra_path):
            subprocess.Popen([sumatra_path, "-page", str(page_number), "-search", keyword, file_path])
        else:
            os.startfile(file_path)  # fallback: hanya buka file
    except Exception as e:
        messagebox.showerror("Gagal Membuka PDF", f"Tidak bisa membuka file:\n{e}")

# GUI setup
root = tk.Tk()
root.title("Deep Index PDF - by: gungratz")
root.geometry("900x550")

folder_var = tk.StringVar()

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Label(frame, text="Folder PDF:").grid(row=0, column=0, sticky="w")
tk.Entry(frame, textvariable=folder_var, width=60).grid(row=0, column=1, padx=5)
tk.Button(frame, text="Pilih", command=browse_folder).grid(row=0, column=2)

tk.Label(frame, text="Kata kunci:").grid(row=1, column=0, sticky="w")
keyword_entry = tk.Entry(frame, width=30)
keyword_entry.grid(row=1, column=1, sticky="w")

tk.Button(frame, text="Cari", command=run_search, width=20).grid(row=2, column=1, pady=10)

# Treeview + Scrollbars
tree_frame = tk.Frame(root)
tree_frame.pack(fill="both", expand=True)

tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
tree_scroll_y.pack(side="right", fill="y")

tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
tree_scroll_x.pack(side="bottom", fill="x")

tree = ttk.Treeview(tree_frame, columns=("File", "Halaman", "Cuplikan", "Path"),
                    show="headings", yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

tree.heading("File", text="File")
tree.heading("Halaman", text="Halaman")
tree.heading("Cuplikan", text="Cuplikan")
tree.heading("Path", text="Path")

tree.column("File", width=200)
tree.column("Halaman", width=80)
tree.column("Cuplikan", width=500)
tree.column("Path", width=0, stretch=False)  # Sembunyikan kolom path

tree.pack(fill="both", expand=True)

tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)

tree.bind("<Double-1>", open_pdf_from_tree)

save_button = tk.Button(root, text="Simpan Hasil ke CSV", command=save_results, state=tk.DISABLED)
save_button.pack(pady=10)

# Load last folder on startup
last_folder = load_last_folder()
if last_folder:
    folder_var.set(last_folder)

root.mainloop()
