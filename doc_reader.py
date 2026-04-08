"""
Developer: https://github.com/zrnge
Document Reader - A GUI application to read .docx and .doc files
Requirements:
    pip install python-docx textract antiword
    
    For .doc support, you need one of:
    - antiword (Linux: sudo apt install antiword)
    - libreoffice (for fallback conversion)
    
Quick install:
    pip install python-docx
    sudo apt install antiword  # Linux .doc support
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
import threading


# ─── .docx reader ───────────────────────────────────────────────────────────
def read_docx(filepath):
    """Read a .docx file using python-docx."""
    try:
        from docx import Document
    except ImportError:
        raise ImportError(
            "python-docx is not installed.\n"
            "Install it with: pip install python-docx"
        )

    doc = Document(filepath)
    paragraphs = []
    for para in doc.paragraphs:
        paragraphs.append(para.text)
    return "\n".join(paragraphs)


# ─── .doc reader (legacy format) ───────────────────────────────────────────
def read_doc(filepath):
    """
    Read a legacy .doc file.
    Tries antiword first, then falls back to LibreOffice conversion.
    """
    # Attempt 1: antiword
    try:
        result = subprocess.run(
            ["antiword", filepath],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0:
            return result.stdout
    except FileNotFoundError:
        pass  # antiword not installed

    # Attempt 2: LibreOffice headless conversion to txt
    try:
        tmp_dir = os.path.dirname(filepath)
        result = subprocess.run(
            [
                "libreoffice", "--headless", "--convert-to", "txt:Text",
                "--outdir", tmp_dir, filepath
            ],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0:
            txt_path = os.path.splitext(filepath)[0] + ".txt"
            if os.path.exists(txt_path):
                with open(txt_path, "r", encoding="utf-8", errors="replace") as f:
                    content = f.read()
                os.remove(txt_path)  # clean up
                return content
    except FileNotFoundError:
        pass  # libreoffice not installed

    raise RuntimeError(
        "Cannot read .doc files.\n\n"
        "Install one of the following:\n"
        "  • antiword  →  sudo apt install antiword\n"
        "  • LibreOffice  →  sudo apt install libreoffice\n\n"
        "On Windows, install LibreOffice and add it to PATH."
    )


# ─── Unified reader ────────────────────────────────────────────────────────
def read_document(filepath):
    """Route to the correct reader based on file extension."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".docx":
        return read_docx(filepath)
    elif ext == ".doc":
        return read_doc(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


# ═══════════════════════════════════════════════════════════════════════════
#  GUI Application
# ═══════════════════════════════════════════════════════════════════════════
class DocReaderApp:
    BG = "#1e1e2e"
    FG = "#cdd6f4"
    ACCENT = "#89b4fa"
    SURFACE = "#313244"
    BORDER = "#45475a"
    RED = "#f38ba8"
    GREEN = "#a6e3a1"

    def __init__(self, root):
        self.root = root
        self.root.title("Document Reader")
        self.root.geometry("900x650")
        self.root.configure(bg=self.BG)
        self.root.minsize(600, 400)

        self.current_file = None
        self._build_ui()

    # ── UI Construction ─────────────────────────────────────────────────
    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Configure styles
        style.configure("App.TFrame", background=self.BG)
        style.configure("App.TLabel", background=self.BG, foreground=self.FG,
                        font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=self.BG, foreground=self.FG,
                        font=("Segoe UI", 16, "bold"))
        style.configure("Status.TLabel", background=self.SURFACE,
                        foreground=self.FG, font=("Segoe UI", 9))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

        # ── Header ──────────────────────────────────────────────────────
        header = ttk.Frame(self.root, style="App.TFrame")
        header.pack(fill=tk.X, padx=20, pady=(15, 5))

        ttk.Label(header, text="📄 Document Reader", style="Title.TLabel"
                  ).pack(side=tk.LEFT)

        btn_frame = ttk.Frame(header, style="App.TFrame")
        btn_frame.pack(side=tk.RIGHT)

        self.open_btn = tk.Button(
            btn_frame, text="Open File", command=self.open_file,
            bg=self.ACCENT, fg="#1e1e2e", font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
            activebackground="#b4d0fb", activeforeground="#1e1e2e"
        )
        self.open_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.clear_btn = tk.Button(
            btn_frame, text="Clear", command=self.clear_text,
            bg=self.SURFACE, fg=self.FG, font=("Segoe UI", 10),
            relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
            activebackground=self.BORDER, activeforeground=self.FG
        )
        self.clear_btn.pack(side=tk.LEFT)

        # ── File info bar ───────────────────────────────────────────────
        self.file_label = ttk.Label(
            self.root, text="No file loaded",
            style="App.TLabel", font=("Segoe UI", 9, "italic")
        )
        self.file_label.pack(fill=tk.X, padx=22, pady=(5, 2))

        # ── Separator ──────────────────────────────────────────────────
        sep = tk.Frame(self.root, height=1, bg=self.BORDER)
        sep.pack(fill=tk.X, padx=20, pady=(2, 8))

        # ── Text area with scrollbar ────────────────────────────────────
        text_frame = tk.Frame(self.root, bg=self.BG)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 8))

        self.scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL,
                                       bg=self.SURFACE, troughcolor=self.BG)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_area = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Consolas", 11),
            bg=self.SURFACE,
            fg=self.FG,
            insertbackground=self.ACCENT,
            selectbackground=self.ACCENT,
            selectforeground="#1e1e2e",
            relief=tk.FLAT,
            padx=14,
            pady=10,
            yscrollcommand=self.scrollbar.set,
            state=tk.DISABLED,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=self.BORDER,
            highlightcolor=self.ACCENT,
        )
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.text_area.yview)

        # ── Status bar ──────────────────────────────────────────────────
        status_frame = tk.Frame(self.root, bg=self.SURFACE, height=28)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)

        self.status_label = tk.Label(
            status_frame, text="Ready",
            bg=self.SURFACE, fg=self.FG,
            font=("Segoe UI", 9), anchor=tk.W, padx=12
        )
        self.status_label.pack(fill=tk.X, expand=True)

        self.word_count_label = tk.Label(
            status_frame, text="Words: 0  |  Lines: 0",
            bg=self.SURFACE, fg=self.FG,
            font=("Segoe UI", 9), anchor=tk.E, padx=12
        )
        self.word_count_label.pack(side=tk.RIGHT)

        # ── Keyboard shortcuts ──────────────────────────────────────────
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-q>", lambda e: self.root.quit())

    # ── Actions ─────────────────────────────────────────────────────────
    def open_file(self):
        filepath = filedialog.askopenfilename(
            title="Select a Document",
            filetypes=[
                ("Word Documents", "*.docx *.doc"),
                ("DOCX files", "*.docx"),
                ("DOC files (legacy)", "*.doc"),
                ("All files", "*.*"),
            ]
        )
        if not filepath:
            return

        self.current_file = filepath
        filename = os.path.basename(filepath)
        size_kb = os.path.getsize(filepath) / 1024

        self.file_label.config(
            text=f"📁  {filename}   ({size_kb:.1f} KB)",
            font=("Segoe UI", 9)
        )
        self.status_label.config(text=f"Loading {filename}...", fg=self.ACCENT)
        self.root.update_idletasks()

        # Load in a thread to keep UI responsive for large files
        threading.Thread(target=self._load_file, args=(filepath,), daemon=True).start()

    def _load_file(self, filepath):
        try:
            content = read_document(filepath)
            self.root.after(0, self._display_content, content)
        except Exception as e:
            self.root.after(0, self._show_error, str(e))

    def _display_content(self, content):
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, content)
        self.text_area.config(state=tk.DISABLED)

        # Update stats
        words = len(content.split())
        lines = content.count("\n") + 1
        self.word_count_label.config(text=f"Words: {words:,}  |  Lines: {lines:,}")
        self.status_label.config(text="File loaded successfully", fg=self.GREEN)

    def _show_error(self, message):
        self.status_label.config(text="Error loading file", fg=self.RED)
        messagebox.showerror("Error", message)

    def clear_text(self):
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete("1.0", tk.END)
        self.text_area.config(state=tk.DISABLED)
        self.file_label.config(text="No file loaded",
                               font=("Segoe UI", 9, "italic"))
        self.word_count_label.config(text="Words: 0  |  Lines: 0")
        self.status_label.config(text="Ready", fg=self.FG)
        self.current_file = None


# ═══════════════════════════════════════════════════════════════════════════
#  Entry point
# ═══════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    app = DocReaderApp(root)
    root.mainloop()
