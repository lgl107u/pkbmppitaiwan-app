"""
PKBM PPI Taiwan - Generator App
Modern GUI untuk generate Rapot, SKHUPK, Kartu UPK, dan Transkrip Nilai
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import shutil
from typing import Optional, Dict, Callable, List
import threading

# Try to import PIL for logo support
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


# =============================================================================
# HELPER FUNCTIONS FOR EXECUTABLE SUPPORT
# =============================================================================
def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Running in normal Python environment
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)


# =============================================================================
# CONSTANTS
# =============================================================================
APP_TITLE = "PKBM PPI Taiwan - Generator App"
APP_VERSION = "2.4"

# Responsive window sizing
def get_window_dimensions():
    """Calculate responsive window dimensions based on screen size."""
    import tkinter as tk
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.destroy()
    
    # Use 80% of screen with min/max constraints
    width = min(1100, max(780, int(screen_width * 0.80)))
    height = min(820, max(620, int(screen_height * 0.80)))
    return width, height

WINDOW_WIDTH, WINDOW_HEIGHT = get_window_dimensions()

# Colors
COLOR_PRIMARY = "#2563eb"
COLOR_PRIMARY_HOVER = "#1d4ed8"
COLOR_SUCCESS = "#16a34a"
COLOR_SUCCESS_HOVER = "#15803d"
COLOR_WARNING = "#ea580c"
COLOR_WARNING_HOVER = "#c2410c"
COLOR_DANGER = "#dc2626"
COLOR_BG = "#f8fafc"
COLOR_CARD = "#ffffff"
COLOR_TEXT = "#1e293b"
COLOR_TEXT_SECONDARY = "#64748b"
COLOR_BORDER = "#e2e8f0"


# =============================================================================
# GENERATOR CONFIGURATIONS
# =============================================================================
GENERATORS = {
    "rapot": {
        "name": "Rapot",
        "description": "Generate Laporan Hasil Belajar Siswa (butuh 2 file Excel)",
        "color": COLOR_PRIMARY,
        "hover": COLOR_PRIMARY_HOVER,
        "files_needed": [
            {"key": "data_siswa", "label": "File Data Siswa (.xlsx)", "required": True},
            {"key": "nilai_siswa", "label": "File Nilai Siswa (.xlsx)", "required": True},
        ],
        "has_wali_signatures": True,
        "output_folder": "./output/rapot",
        "template_files": ["data/rapot/Template Data Siswa Rapot.xlsx", "data/rapot/Template Nilai Rapot.xlsx"]
    },
    "skhupk": {
        "name": "SKHUPK", 
        "description": "Generate Surat Keterangan Hasil UPK",
        "color": COLOR_SUCCESS,
        "hover": COLOR_SUCCESS_HOVER,
        "files_needed": [
            {"key": "excel_skhupk", "label": "File Database SKHUPK (.xlsx)", "required": True},
            {"key": "ttd_kepsek", "label": "Tanda Tangan Kepsek (opsional)", "required": False, "file_type": "image"},
        ],
        "output_folder": "./output/skhupk",
        "template_files": ["data/skhupk/Database  SKHUPK 2025.xlsx"]
    },
    "kartu_upk": {
        "name": "Kartu UPK",
        "description": "Generate Kartu Peserta UPK Paket C (85.60x53.98mm)",
        "color": COLOR_WARNING,
        "hover": COLOR_WARNING_HOVER,
        "files_needed": [
            {"key": "excel_kartu", "label": "File Data Kartu UPK (.xlsx)", "required": True},
            {"key": "foto_folder_upk", "label": "Folder Foto Siswa (opsional)", "required": False, "file_type": "folder"},
        ],
        "has_tahun_ajaran": True,
        "output_folder": "./output/kartu_upk",
        "template_files": ["data/kartu_upk/template_kartu_upk.xlsx"]
    },
    "kartu_siswa": {
        "name": "Kartu Pelajar",
        "description": "Generate Kartu Pelajar / Identitas Siswa (85.60x53.98mm)",
        "color": "#ec4899",
        "hover": "#db2777",
        "files_needed": [
            {"key": "excel_kartu_siswa", "label": "File Data Kartu Siswa (.xlsx)", "required": True},
            {"key": "foto_folder_siswa", "label": "Folder Foto Siswa (opsional)", "required": False, "file_type": "folder"},
        ],
        "has_tahun_ajaran": True,
        "output_folder": "./output/kartu_siswa",
        "template_files": ["data/kartu_siswa/template_kartu_siswa.xlsx"]
    },
    "transkrip": {
        "name": "Transkrip Nilai",
        "description": "Generate Transkrip Nilai Siswa",
        "color": "#8b5cf6",
        "hover": "#7c3aed",
        "files_needed": [
            {"key": "excel_transkrip", "label": "File Database Transkrip (.xlsx)", "required": True},
            {"key": "ttd_kepsek", "label": "Tanda Tangan Kepsek (opsional)", "required": False, "file_type": "image"},
        ],
        "output_folder": "./output/transkrip",
        "template_files": ["data/transkrip/Database Transkrip Nilai.xlsx"]
    }
}


# =============================================================================
# MODERN BUTTON COMPONENT
# =============================================================================
class ModernButton(tk.Frame):
    """Modern styled button with hover effects."""
    
    def __init__(
        self, 
        parent, 
        text: str, 
        command: Optional[Callable] = None,
        bg_color: str = COLOR_PRIMARY,
        hover_color: str = COLOR_PRIMARY_HOVER,
        fg_color: str = "white",
        width: int = 200,
        height: int = 45,
        font_size: int = 11,
        corner_radius: int = 8
    ):
        super().__init__(parent, bg=parent.cget("bg"))
        
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.fg_color = fg_color
        self.command = command
        self.enabled = True
        
        # Main button frame
        self.btn_frame = tk.Frame(self, bg=bg_color, width=width, height=height)
        self.btn_frame.pack(fill=tk.BOTH, expand=True)
        self.btn_frame.pack_propagate(False)
        
        # Text label
        self.label = tk.Label(
            self.btn_frame, 
            text=text, 
            font=("Segoe UI", font_size, "bold"),
            bg=bg_color, 
            fg=fg_color,
            cursor="hand2"
        )
        self.label.pack(expand=True)
        
        # Bind events
        self._bind_events()
    
    def _bind_events(self):
        for widget in [self.btn_frame, self.label]:
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
            widget.bind("<Button-1>", self._on_click)
    
    def _on_enter(self, event):
        if self.enabled:
            self.btn_frame.config(bg=self.hover_color)
            self.label.config(bg=self.hover_color)
    
    def _on_leave(self, event):
        if self.enabled:
            self.btn_frame.config(bg=self.bg_color)
            self.label.config(bg=self.bg_color)
    
    def _on_click(self, event):
        if self.enabled and self.command:
            self.command()
    
    def config(self, text=None, bg_color=None, hover_color=None, fg_color=None):
        """Update button configuration dynamically."""
        if text is not None:
            self.label.config(text=text)
        if bg_color is not None:
            self.bg_color = bg_color
            self.btn_frame.config(bg=bg_color)
            self.label.config(bg=bg_color)
        if hover_color is not None:
            self.hover_color = hover_color
        if fg_color is not None:
            self.fg_color = fg_color
            self.label.config(fg=fg_color)
    
    def set_enabled(self, enabled: bool):
        self.enabled = enabled
        if enabled:
            self.btn_frame.config(bg=self.bg_color)
            self.label.config(bg=self.bg_color, fg=self.fg_color)
            self.label.config(cursor="hand2")
        else:
            self.btn_frame.config(bg="#94a3b8")
            self.label.config(bg="#94a3b8", fg="#cbd5e1")
            self.label.config(cursor="arrow")


# =============================================================================
# FILE SELECTOR COMPONENT
# =============================================================================
class FileSelector(tk.Frame):
    """File selector with browse button and path display."""
    
    def __init__(
        self, 
        parent, 
        label: str,
        file_types: tuple = (("Excel files", "*.xlsx *.xls"), ("All files", "*.*")),
        on_file_selected: Optional[Callable] = None,
        is_image: bool = False,
        is_folder: bool = False
    ):
        if is_image:
            file_types = (("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*"))
        super().__init__(parent, bg=COLOR_CARD)
        
        self.file_path = tk.StringVar()
        self.file_types = file_types
        self.on_file_selected = on_file_selected
        self.is_folder = is_folder
        
        # Label
        self.lbl = tk.Label(
            self, 
            text=label, 
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD, 
            fg=COLOR_TEXT,
            anchor="w"
        )
        self.lbl.pack(fill=tk.X, pady=(2, 2))
        
        # Container for entry and button
        container = tk.Frame(self, bg=COLOR_CARD)
        container.pack(fill=tk.X, pady=(0, 5))
        
        # Entry
        self.entry = tk.Entry(
            container,
            textvariable=self.file_path,
            font=("Segoe UI", 9),
            state="readonly",
            readonlybackground="#f1f5f9",
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_PRIMARY
        )
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        
        # Browse button
        self.browse_btn = tk.Button(
            container,
            text="Browse...",
            font=("Segoe UI", 9),
            bg=COLOR_PRIMARY,
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=4,
            command=self._browse
        )
        self.browse_btn.pack(side=tk.RIGHT, padx=(8, 0))
        
        # Hover effect for browse button
        self.browse_btn.bind("<Enter>", lambda e: self.browse_btn.config(bg=COLOR_PRIMARY_HOVER))
        self.browse_btn.bind("<Leave>", lambda e: self.browse_btn.config(bg=COLOR_PRIMARY))
    
    def _browse(self):
        if self.is_folder:
            folder_path = filedialog.askdirectory(title="Pilih Folder Foto Siswa")
            if folder_path:
                self.file_path.set(folder_path)
                if self.on_file_selected:
                    self.on_file_selected(folder_path)
        else:
            file_path = filedialog.askopenfilename(
                title="Pilih File",
                filetypes=self.file_types
            )
            if file_path:
                self.file_path.set(file_path)
                if self.on_file_selected:
                    self.on_file_selected(file_path)
    
    def get_path(self) -> str:
        return self.file_path.get()
    
    def clear(self):
        self.file_path.set("")
    
    def set_enabled(self, enabled: bool):
        """Enable or disable the file selector."""
        state = "normal" if enabled else "disabled"
        self.browse_btn.config(state=state)


# =============================================================================
# GENERATOR CARD COMPONENT
# =============================================================================
class GeneratorCard(tk.Frame):
    """Card component for each generator type."""
    
    def __init__(
        self, 
        parent, 
        generator_key: str,
        config: dict,
        on_select: Callable
    ):
        super().__init__(parent, bg=COLOR_CARD, relief="flat", bd=0)
        
        self.generator_key = generator_key
        self.config = config
        self.on_select = on_select
        self.selected = False
        
        # Configure padding
        self.configure(highlightthickness=2, highlightbackground=COLOR_BORDER)
        
        # Inner container
        inner = tk.Frame(self, bg=COLOR_CARD, padx=12, pady=10)
        inner.pack(fill=tk.BOTH, expand=True)
        
        # Color indicator
        indicator = tk.Frame(inner, bg=config["color"], width=5)
        indicator.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))
        
        # Content
        content = tk.Frame(inner, bg=COLOR_CARD)
        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Title
        self.title_lbl = tk.Label(
            content,
            text=config["name"],
            font=("Segoe UI", 12, "bold"),
            bg=COLOR_CARD,
            fg=COLOR_TEXT,
            anchor="w"
        )
        self.title_lbl.pack(fill=tk.X)
        
        # Bind click events
        for widget in [self, inner, indicator, content, self.title_lbl]:
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
            widget.configure(cursor="hand2")
    
    def _on_click(self, event):
        self.on_select(self.generator_key)
    
    def _on_enter(self, event):
        if not self.selected:
            self.configure(highlightbackground=self.config["color"])
    
    def _on_leave(self, event):
        if not self.selected:
            self.configure(highlightbackground=COLOR_BORDER)
    
    def set_selected(self, selected: bool):
        self.selected = selected
        if selected:
            self.configure(highlightbackground=self.config["color"], highlightthickness=3)
        else:
            self.configure(highlightbackground=COLOR_BORDER, highlightthickness=2)


# =============================================================================
# MAIN APPLICATION
# =============================================================================
class PKBMGeneratorApp:
    """Main application class."""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.configure(bg=COLOR_BG)
        self.root.resizable(True, True)  # Allow resizing for smaller screens
        
        # Set minimum and maximum window size — ensures menu+Aksi always visible
        self.root.minsize(780, 720)
        self.root.maxsize(1200, 900)
        
        # State for generation
        self.is_generating = False
        self.stop_requested = False
        
        # State
        self.selected_generator: Optional[str] = None
        self.file_selectors: Dict[str, FileSelector] = {}
        self.generator_cards: Dict[str, GeneratorCard] = {}
        
        # Wali kelas signatures: {kelas: path} — 'KEPSEK' key reserved for kepala sekolah
        self.wali_signatures: Dict[str, str] = {}
        self.guru_data: List[Dict] = []  # List of {kelas, nama} from Excel
        self.kepsek_data: Optional[Dict] = None  # {kelas: 'KEPSEK', nama: '...'}
        
        # Build UI — footer packed first so it is never hidden during resize
        self._create_header()
        self._create_footer()
        self._create_main_content()
    
    def _create_header(self):
        """Create header with logo and title."""
        header = tk.Frame(self.root, bg=COLOR_CARD, height=70)
        header.pack(fill=tk.X, padx=15, pady=(10, 0))
        header.pack_propagate(False)
        
        # Inner container
        inner = tk.Frame(header, bg=COLOR_CARD)
        inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # Logo
        logo_path = get_resource_path(os.path.join("assets", "logo.png"))
        if PIL_AVAILABLE and os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                img = img.resize((50, 50), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                logo_lbl = tk.Label(inner, image=self.logo_img, bg=COLOR_CARD)
                logo_lbl.pack(side=tk.LEFT, padx=(0, 10))
            except Exception:
                pass
        
        # Title section
        title_frame = tk.Frame(inner, bg=COLOR_CARD)
        title_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        title = tk.Label(
            title_frame,
            text="PKBM PPI TAIWAN",
            font=("Segoe UI", 16, "bold"),
            bg=COLOR_CARD,
            fg=COLOR_TEXT
        )
        title.pack(anchor="w", pady=(2, 0))
        
        subtitle = tk.Label(
            title_frame,
            text="Aplikasi Generator Dokumen Pendidikan",
            font=("Segoe UI", 9),
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY
        )
        subtitle.pack(anchor="w", pady=(2, 0))
        
        # Version label on the right
        version = tk.Label(
            inner,
            text=f"v{APP_VERSION}",
            font=("Segoe UI", 9),
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY
        )
        version.pack(side=tk.RIGHT, padx=(10, 0))
    
    def _create_main_content(self):
        """Create main content area."""
        main = tk.Frame(self.root, bg=COLOR_BG)
        main.pack(fill=tk.BOTH, expand=True, padx=15, pady=(15, 4))
        
        # Left panel with fixed width and minimum height to keep Aksi visible
        left_panel = tk.Frame(main, bg=COLOR_BG, width=300, height=550)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))
        left_panel.pack_propagate(False)
        
        # Section title
        section_title = tk.Label(
            left_panel,
            text="Pilih Jenis Dokumen",
            font=("Segoe UI", 13, "bold"),
            bg=COLOR_BG,
            fg=COLOR_TEXT
        )
        section_title.pack(anchor="w", pady=(0, 12))
        
        # Generator cards
        for key, config in GENERATORS.items():
            card = GeneratorCard(
                left_panel,
                generator_key=key,
                config=config,
                on_select=self._on_generator_select
            )
            card.pack(fill=tk.X, pady=(0, 10))
            self.generator_cards[key] = card
        
        # =====================================================================
        # Action Buttons Group (below document type selector)
        # =====================================================================
        action_groupbox = tk.LabelFrame(
            left_panel,
            text="⚡ Aksi",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_BG,
            fg=COLOR_TEXT,
            padx=8,
            pady=8,
            height=120
        )
        action_groupbox.pack(fill=tk.X, pady=(8, 0))
        action_groupbox.pack_propagate(False)
        
        # Generate PDF button (will toggle to Stop during generation)
        self.generate_btn = ModernButton(
            action_groupbox,
            text="🚀 Generate",
            command=self._toggle_generate,
            bg_color=COLOR_SUCCESS,
            hover_color=COLOR_SUCCESS_HOVER,
            width=260,
            height=38,
            font_size=10
        )
        self.generate_btn.pack(fill=tk.X, pady=(0, 6))
        
        # Row for Clear and Download Template
        btn_row = tk.Frame(action_groupbox, bg=COLOR_BG)
        btn_row.pack(fill=tk.X)
        
        self.clear_btn = ModernButton(
            btn_row,
            text="🗑 Clear",
            command=self._clear_files,
            bg_color="#64748b",
            hover_color="#475569",
            width=120,
            height=34,
            font_size=9
        )
        self.clear_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        
        self.template_btn = ModernButton(
            btn_row,
            text="📥 Template",
            command=self._download_template,
            bg_color="#8b5cf6",
            hover_color="#7c3aed",
            width=120,
            height=34,
            font_size=9
        )
        self.template_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        
        # Right panel - File selection and actions
        right_panel = tk.Frame(main, bg=COLOR_CARD)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Right panel inner padding
        right_inner = tk.Frame(right_panel, bg=COLOR_CARD, padx=20, pady=18)
        right_inner.pack(fill=tk.BOTH, expand=True)
        
        # Instruction label
        self.instruction_lbl = tk.Label(
            right_inner,
            text="← Pilih jenis dokumen yang ingin di-generate",
            font=("Segoe UI", 11),
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY
        )
        self.instruction_lbl.pack(expand=True)
        
        # File selectors container (hidden initially)
        self.file_container = tk.Frame(right_inner, bg=COLOR_CARD)
        
        # Selected generator title
        self.selected_title = tk.Label(
            self.file_container,
            text="",
            font=("Segoe UI", 14, "bold"),
            bg=COLOR_CARD,
            fg=COLOR_TEXT
        )
        self.selected_title.pack(anchor="w", pady=(0, 4))
        
        # Selected generator description
        self.selected_desc = tk.Label(
            self.file_container,
            text="",
            font=("Segoe UI", 9),
            bg=COLOR_CARD,
            fg=COLOR_TEXT_SECONDARY
        )
        self.selected_desc.pack(anchor="w", pady=(0, 12))
        
        # =====================================================================
        # GROUPBOX: File Input Section
        # =====================================================================
        self.file_groupbox = tk.LabelFrame(
            self.file_container,
            text="📁 Pilih File Input",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD,
            fg=COLOR_TEXT,
            padx=12,
            pady=8
        )
        self.file_groupbox.pack(fill=tk.X, pady=(0, 10))
        self.file_groupbox.pack_propagate(True)
        
        # File inputs container inside groupbox
        self.file_inputs_container = tk.Frame(self.file_groupbox, bg=COLOR_CARD)
        self.file_inputs_container.pack(fill=tk.BOTH, expand=True)
        
        # =====================================================================
        # GROUPBOX: Output Folder Section
        # =====================================================================
        output_groupbox = tk.LabelFrame(
            self.file_container,
            text="📂 Folder Output",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD,
            fg=COLOR_TEXT,
            padx=12,
            pady=8
        )
        output_groupbox.pack(fill=tk.X, pady=(0, 10))
        
        output_container = tk.Frame(output_groupbox, bg=COLOR_CARD)
        output_container.pack(fill=tk.X, pady=(4, 4))
        
        self.output_path = tk.StringVar()
        self.output_entry = tk.Entry(
            output_container,
            textvariable=self.output_path,
            font=("Segoe UI", 9),
            state="readonly",
            readonlybackground="#f1f5f9",
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLOR_BORDER
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        
        output_btn = tk.Button(
            output_container,
            text="Ubah...",
            font=("Segoe UI", 9),
            bg="#64748b",
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=4,
            command=self._browse_output
        )
        output_btn.pack(side=tk.RIGHT, padx=(8, 0))
        
        # Keep reference for output_frame compatibility
        self.output_frame = output_groupbox
        
        # =====================================================================
        # GROUPBOX: Console Log Section
        # =====================================================================
        log_groupbox = tk.LabelFrame(
            self.file_container,
            text="📋 Status Log",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD,
            fg=COLOR_TEXT,
            padx=10,
            pady=8
        )
        log_groupbox.pack(fill=tk.BOTH, expand=True, pady=(0, 2))
        
        self.log_frame = log_groupbox
        
        # Scrollable log text
        log_container = tk.Frame(log_groupbox, bg=COLOR_BORDER)
        log_container.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(
            log_container,
            font=("Consolas", 8),
            bg="#1e293b",
            fg="#e2e8f0",
            wrap=tk.WORD,
            state="disabled",
            relief="flat",
            padx=8,
            pady=8,
            height=10
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        log_scrollbar = ttk.Scrollbar(log_container, orient="vertical", command=self.log_text.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        
        # Configure text tags for colors
        self.log_text.tag_configure("info", foreground="#60a5fa")
        self.log_text.tag_configure("success", foreground="#4ade80")
        self.log_text.tag_configure("error", foreground="#f87171")
        self.log_text.tag_configure("warning", foreground="#fbbf24")
        
        # Progress bar (inside log groupbox)
        self.progress = ttk.Progressbar(
            log_groupbox,
            mode="indeterminate",
            length=400
        )
    
    def _create_footer(self):
        """Create footer."""
        footer = tk.Frame(self.root, bg=COLOR_BG, height=28)
        footer.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=(0, 8))
        
        footer_text = tk.Label(
            footer,
            text="© 2025 PKBM PPI Taiwan | Developed by Boby HP",
            font=("Segoe UI", 10),
            bg=COLOR_BG,
            fg=COLOR_TEXT_SECONDARY
        )
        footer_text.pack(side=tk.RIGHT)
    
    def _log(self, message: str, level: str = "info"):
        """Add message to log panel."""
        self.log_text.config(state="normal")
        
        # Add timestamp
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        
        # Insert with appropriate tag
        prefix = {"info": "ℹ️", "success": "✅", "error": "❌", "warning": "⚠️"}.get(level, "")
        self.log_text.insert(tk.END, f"[{timestamp}] {prefix} {message}\n", level)
        
        # Auto-scroll to bottom
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update_idletasks()
    
    def _clear_log(self):
        """Clear the log panel."""
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")
    
    def _set_ui_enabled(self, enabled: bool):
        """Enable or disable UI elements during generation."""
        self.is_generating = not enabled
        
        # Disable/enable generator cards
        for card in self.generator_cards.values():
            for child in card.winfo_children():
                if hasattr(child, 'config'):
                    try:
                        child.config(state="normal" if enabled else "disabled")
                    except:
                        pass
        
        # Disable/enable file selectors
        for selector in self.file_selectors.values():
            selector.set_enabled(enabled)
        
        # Disable/enable buttons (keep generate_btn enabled so Stop is clickable)
        # self.generate_btn.set_enabled(enabled)  # Keep enabled for Stop functionality
        self.clear_btn.set_enabled(enabled)
        self.template_btn.set_enabled(enabled)
    
    def _on_generator_select(self, generator_key: str):
        """Handle generator selection."""
        # Don't allow selection during generation
        if self.is_generating:
            return
            
        # Update selection
        self.selected_generator = generator_key
        config = GENERATORS[generator_key]
        
        # Update card states
        for key, card in self.generator_cards.items():
            card.set_selected(key == generator_key)
        
        # Hide instruction, show file container
        self.instruction_lbl.pack_forget()
        self.file_container.pack(fill=tk.BOTH, expand=True)
        
        # Update title and description
        self.selected_title.config(text=config["name"])
        self.selected_desc.config(text=config["description"])
        
        # Clear and rebuild file inputs
        for widget in self.file_inputs_container.winfo_children():
            widget.destroy()
        self.file_selectors.clear()
        
        # Create file selectors - group optional files in rows of 2
        required_files = [f for f in config["files_needed"] if f.get("required")]
        optional_files = [f for f in config["files_needed"] if not f.get("required")]
        
        # Required files - full width
        for file_config in required_files:
            is_image = file_config.get("file_type") == "image"
            is_folder = file_config.get("file_type") == "folder"
            selector = FileSelector(
                self.file_inputs_container,
                label=file_config["label"] + " *",
                on_file_selected=self._on_file_selected,
                is_image=is_image,
                is_folder=is_folder
            )
            selector.pack(fill=tk.X, pady=(0, 2))
            self.file_selectors[file_config["key"]] = selector
        
        # Optional files - dynamic responsive layout
        if optional_files:
            opt_frame = tk.Frame(self.file_inputs_container, bg=COLOR_CARD)
            opt_frame.pack(fill=tk.X, pady=(5, 0))
            
            num_files = len(optional_files)
            
            if num_files == 1:
                # Single file - full width
                opt_frame.columnconfigure(0, weight=1)
                file_config = optional_files[0]
                is_image = file_config.get("file_type") == "image"
                is_folder = file_config.get("file_type") == "folder"
                selector = FileSelector(
                    opt_frame,
                    label=file_config["label"],
                    on_file_selected=self._on_file_selected,
                    is_image=is_image,
                    is_folder=is_folder
                )
                selector.grid(row=0, column=0, sticky="ew", pady=(0, 2))
                self.file_selectors[file_config["key"]] = selector
                
            elif num_files == 2:
                # 2 files - 1 row, 2 columns
                opt_frame.columnconfigure(0, weight=1)
                opt_frame.columnconfigure(1, weight=1)
                for idx, file_config in enumerate(optional_files):
                    is_image = file_config.get("file_type") == "image"
                    is_folder = file_config.get("file_type") == "folder"
                    selector = FileSelector(
                        opt_frame,
                        label=file_config["label"],
                        on_file_selected=self._on_file_selected,
                        is_image=is_image,
                        is_folder=is_folder
                    )
                    selector.grid(row=0, column=idx, sticky="ew", padx=(0 if idx == 0 else 5, 0), pady=(0, 2))
                    self.file_selectors[file_config["key"]] = selector
                    
            elif num_files == 3:
                # 3 files - first row full width, second row 2 columns
                opt_frame.columnconfigure(0, weight=1)
                opt_frame.columnconfigure(1, weight=1)
                
                # First file - full width
                file_config = optional_files[0]
                is_image = file_config.get("file_type") == "image"
                is_folder = file_config.get("file_type") == "folder"
                selector = FileSelector(
                    opt_frame,
                    label=file_config["label"],
                    on_file_selected=self._on_file_selected,
                    is_image=is_image,
                    is_folder=is_folder
                )
                selector.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 2))
                self.file_selectors[file_config["key"]] = selector
                
                # Remaining 2 files - second row
                for idx, file_config in enumerate(optional_files[1:]):
                    is_image = file_config.get("file_type") == "image"
                    is_folder = file_config.get("file_type") == "folder"
                    selector = FileSelector(
                        opt_frame,
                        label=file_config["label"],
                        on_file_selected=self._on_file_selected,
                        is_image=is_image,
                        is_folder=is_folder
                    )
                    selector.grid(row=1, column=idx, sticky="ew", padx=(0 if idx == 0 else 5, 0), pady=(0, 2))
                    self.file_selectors[file_config["key"]] = selector
                    
            elif num_files == 4:
                # 4 files - 2x2 grid
                opt_frame.columnconfigure(0, weight=1)
                opt_frame.columnconfigure(1, weight=1)
                for idx, file_config in enumerate(optional_files):
                    is_image = file_config.get("file_type") == "image"
                    is_folder = file_config.get("file_type") == "folder"
                    selector = FileSelector(
                        opt_frame,
                        label=file_config["label"],
                        on_file_selected=self._on_file_selected,
                        is_image=is_image,
                        is_folder=is_folder
                    )
                    row = idx // 2
                    col = idx % 2
                    selector.grid(row=row, column=col, sticky="ew", padx=(0 if col == 0 else 5, 0), pady=(0, 2))
                    self.file_selectors[file_config["key"]] = selector
                    
            else:
                # More than 4 files - default 2 per row
                opt_frame.columnconfigure(0, weight=1)
                opt_frame.columnconfigure(1, weight=1)
                for idx, file_config in enumerate(optional_files):
                    is_image = file_config.get("file_type") == "image"
                    is_folder = file_config.get("file_type") == "folder"
                    selector = FileSelector(
                        opt_frame,
                        label=file_config["label"],
                        on_file_selected=self._on_file_selected,
                        is_image=is_image,
                        is_folder=is_folder
                    )
                    row = idx // 2
                    col = idx % 2
                    selector.grid(row=row, column=col, sticky="ew", padx=(0 if col == 0 else 5, 0), pady=(0, 2))
                    self.file_selectors[file_config["key"]] = selector
        
        # File groupbox will auto-size based on content
        
        # Add wali kelas signature management button for rapot generator
        if config.get("has_wali_signatures"):
            wali_frame = tk.Frame(self.file_inputs_container, bg=COLOR_CARD)
            wali_frame.pack(fill=tk.X, pady=(8, 0))
            
            wali_label = tk.Label(
                wali_frame,
                text="Tanda Tangan Wali Kelas",
                font=("Segoe UI", 9, "bold"),
                bg=COLOR_CARD,
                fg=COLOR_TEXT,
                anchor="w"
            )
            wali_label.pack(fill=tk.X, pady=(0, 4))
            
            self.wali_btn = tk.Button(
                wali_frame,
                text="📝 Kelola TTD Wali Kelas...",
                font=("Segoe UI", 9),
                bg="#8b5cf6",
                fg="white",
                relief="flat",
                cursor="hand2",
                padx=12,
                pady=6,
                command=self._open_wali_signature_manager
            )
            self.wali_btn.pack(fill=tk.X)
            self.wali_btn.bind("<Enter>", lambda e: self.wali_btn.config(bg="#7c3aed"))
            self.wali_btn.bind("<Leave>", lambda e: self.wali_btn.config(bg="#8b5cf6"))
            
            # Status label for wali signatures
            self.wali_status_label = tk.Label(
                wali_frame,
                text="Belum ada TTD wali kelas (klik tombol di atas untuk menambahkan)",
                font=("Segoe UI", 8),
                bg=COLOR_CARD,
                fg=COLOR_TEXT_SECONDARY,
                anchor="w"
            )
            self.wali_status_label.pack(fill=tk.X, pady=(4, 0))
            self._update_wali_status_label()
        
        # Tahun Ajaran will be loaded from Excel DataTetap automatically
        self.tahun_ajaran_var = None
        
        # Add Kepsek signature management for kartu generators (like rapot)
        if self.selected_generator in ["kartu_upk", "kartu_siswa"]:
            kepsek_frame = tk.Frame(self.file_inputs_container, bg=COLOR_CARD)
            kepsek_frame.pack(fill=tk.X, pady=(8, 0))
            
            kepsek_label = tk.Label(
                kepsek_frame,
                text="Tanda Tangan Kepala PKBM",
                font=("Segoe UI", 9, "bold"),
                bg=COLOR_CARD,
                fg=COLOR_TEXT,
                anchor="w"
            )
            kepsek_label.pack(fill=tk.X, pady=(0, 4))
            
            self.kepsek_btn = tk.Button(
                kepsek_frame,
                text="📝 Kelola TTD Kepala PKBM...",
                font=("Segoe UI", 9),
                bg="#8b5cf6",
                fg="white",
                relief="flat",
                cursor="hand2",
                padx=12,
                pady=6,
                command=self._open_kepsek_signature_manager
            )
            self.kepsek_btn.pack(fill=tk.X)
            self.kepsek_btn.bind("<Enter>", lambda e: self.kepsek_btn.config(bg="#7c3aed"))
            self.kepsek_btn.bind("<Leave>", lambda e: self.kepsek_btn.config(bg="#8b5cf6"))
            
            # Status label for kepsek signature
            self.kepsek_status_label = tk.Label(
                kepsek_frame,
                text="Belum ada TTD Kepala PKBM (klik tombol di atas untuk menambahkan)",
                font=("Segoe UI", 8),
                bg=COLOR_CARD,
                fg=COLOR_TEXT_SECONDARY,
                anchor="w"
            )
            self.kepsek_status_label.pack(fill=tk.X, pady=(4, 0))
            self._update_kepsek_status_label()
        
        # Clear output folder (user must select)
        self.output_path.set("")
        
        # Clear log and reset progress
        self._clear_log()
        self.progress.pack_forget()
    
    def _update_kepsek_status_label(self):
        """Update the kepsek signature status label."""
        if not hasattr(self, 'kepsek_status_label'):
            return
        
        ttd_kepsek = self.wali_signatures.get('KEPSEK')
        if ttd_kepsek:
            self.kepsek_status_label.config(
                text=f"✓ TTD Kepala PKBM: {os.path.basename(ttd_kepsek)}",
                fg=COLOR_SUCCESS
            )
        else:
            self.kepsek_status_label.config(
                text="Belum ada TTD Kepala PKBM (klik tombol di atas untuk menambahkan)",
                fg=COLOR_TEXT_SECONDARY
            )
    
    def _open_kepsek_signature_manager(self):
        """Open popup window to manage Kepsek signature."""
        # Load kepsek name from Excel first
        excel_key = None
        if self.selected_generator == "kartu_upk":
            excel_key = "excel_kartu"
        elif self.selected_generator == "kartu_siswa":
            excel_key = "excel_kartu_siswa"
        
        kepsek_name = "Kepala PKBM"
        if excel_key and excel_key in self.file_selectors:
            excel_path = self.file_selectors[excel_key].get_path()
            if excel_path and os.path.exists(excel_path):
                try:
                    import pandas as pd
                    import warnings
                    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
                    
                    df = pd.read_excel(excel_path, sheet_name="DataTetap")
                    df.columns = df.columns.str.strip()
                    
                    # Look for 'Kepsek' in variabel column (first column)
                    for _, row in df.iterrows():
                        var_name = str(row.iloc[0]).strip()
                        if var_name.lower() == 'kepsek':
                            kepsek_name = str(row.iloc[1]).strip() if len(row) > 1 else "Kepala PKBM"
                            break
                except Exception as e:
                    print(f"Error loading kepsek name: {e}")
        
        # Create popup dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Kelola Tanda Tangan Kepala PKBM")
        dialog.geometry("600x300")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 600) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 300) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Title
        tk.Label(
            dialog,
            text="Tanda Tangan Kepala PKBM",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=(15, 5))
        
        tk.Label(
            dialog,
            text="Pilih file gambar tanda tangan untuk Kepala PKBM",
            font=("Segoe UI", 9),
            fg=COLOR_TEXT_SECONDARY
        ).pack(pady=(0, 15))
        
        # Kepsek row
        row_frame = tk.Frame(dialog, bg="white", relief="groove", bd=1)
        row_frame.pack(fill=tk.X, pady=10, padx=20)
        
        tk.Label(
            row_frame,
            text="Kepala PKBM",
            font=("Segoe UI", 10, "bold"),
            bg="white",
            width=14,
            anchor="w"
        ).pack(side=tk.LEFT, padx=(10, 5), pady=8)
        
        tk.Label(
            row_frame,
            text=kepsek_name,
            font=("Segoe UI", 9),
            bg="white",
            width=22,
            anchor="w"
        ).pack(side=tk.LEFT, padx=5, pady=8)
        
        path_var = tk.StringVar(value=self.wali_signatures.get('KEPSEK', ""))
        
        tk.Entry(
            row_frame,
            textvariable=path_var,
            font=("Segoe UI", 8),
            state="readonly",
            readonlybackground="#f1f5f9",
            width=20
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=8)
        
        def browse():
            fp = filedialog.askopenfilename(
                title="Pilih TTD Kepala PKBM",
                filetypes=(("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*"))
            )
            if fp:
                path_var.set(fp)
        
        tk.Button(
            row_frame,
            text="Browse",
            font=("Segoe UI", 8),
            bg=COLOR_PRIMARY,
            fg="white",
            relief="flat",
            padx=8,
            pady=2,
            command=browse
        ).pack(side=tk.LEFT, padx=(5, 4), pady=8)
        
        def clear():
            path_var.set("")
        
        tk.Button(
            row_frame,
            text="×",
            font=("Segoe UI", 10, "bold"),
            bg="#dc2626",
            fg="white",
            relief="flat",
            width=2,
            command=clear
        ).pack(side=tk.LEFT, padx=(0, 10), pady=8)
        
        # Buttons frame
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        def save_and_close():
            path = path_var.get()
            if path:
                self.wali_signatures['KEPSEK'] = path
            elif 'KEPSEK' in self.wali_signatures:
                del self.wali_signatures['KEPSEK']
            
            self._update_kepsek_status_label()
            self._log(f"TTD Kepala PKBM diperbarui", "success")
            dialog.destroy()
        
        tk.Button(
            btn_frame,
            text="💾 Simpan",
            font=("Segoe UI", 10),
            bg=COLOR_SUCCESS,
            fg="white",
            padx=20,
            pady=5,
            relief="flat",
            command=save_and_close
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Batal",
            font=("Segoe UI", 10),
            bg="#6b7280",
            fg="white",
            padx=20,
            pady=5,
            relief="flat",
            command=dialog.destroy
        ).pack(side=tk.LEFT, padx=5)
    
    def _on_file_selected(self, path: str):
        """Handle file selection."""
        self._log(f"File dipilih: {os.path.basename(path)}", "info")
        
        # If nilai_siswa file is selected, try to load guru data for wali signatures
        if self.selected_generator == "rapot" and "nilai_siswa" in self.file_selectors:
            nilai_path = self.file_selectors["nilai_siswa"].get_path()
            if nilai_path and os.path.exists(nilai_path):
                self._load_guru_data(nilai_path)
    
    def _load_guru_data(self, nilai_path: str):
        """Load guru/wali kelas data from nilai Excel file."""
        # Create loading popup
        loading_dialog = tk.Toplevel(self.root)
        loading_dialog.title("Memuat Data Guru")
        loading_dialog.geometry("350x120")
        loading_dialog.resizable(False, False)
        loading_dialog.transient(self.root)
        loading_dialog.configure(bg=COLOR_BG)
        
        # Center dialog
        loading_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 350) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 120) // 2
        loading_dialog.geometry(f"+{x}+{y}")
        
        tk.Label(
            loading_dialog,
            text="📋 Membaca data guru & wali kelas...",
            font=("Segoe UI", 12, "bold"),
            bg=COLOR_BG,
            fg=COLOR_TEXT
        ).pack(pady=(20, 10))
        
        tk.Label(
            loading_dialog,
            text="Harap tunggu sebentar",
            font=("Segoe UI", 9),
            bg=COLOR_BG,
            fg=COLOR_TEXT_SECONDARY
        ).pack(pady=(0, 10))
        
        progress = ttk.Progressbar(
            loading_dialog,
            mode="indeterminate",
            length=280
        )
        progress.pack(pady=(0, 15))
        progress.start(10)
        
        loading_dialog.update()
        
        try:
            import pandas as pd
            df = pd.read_excel(nilai_path, sheet_name=None)
            
            # Normalize sheet names
            df = {"".join(k.upper().strip().split()): v for k, v in df.items()}
            
            if 'GURU' in df:
                guru_df = df['GURU']
                guru_df.columns = guru_df.columns.str.upper().str.strip()
                
                # Extract kelas and nama columns
                self.guru_data = []
                self.kepsek_data = None
                for _, row in guru_df.iterrows():
                    kelas = str(row.get('KELAS', '')).strip().upper()
                    nama = str(row.get('NAMA', '')).strip()
                    if kelas == 'KEPSEK' and nama:
                        self.kepsek_data = {'kelas': 'KEPSEK', 'nama': nama}
                    elif kelas and nama:
                        self.guru_data.append({'kelas': kelas, 'nama': nama})
                
                self._log(f"Loaded {len(self.guru_data)} wali kelas dari file nilai", "info")
                self._update_wali_status_label()
            
            loading_dialog.destroy()
        except Exception as e:
            loading_dialog.destroy()
            print(f"Error loading guru data: {e}")
    
    def _update_wali_status_label(self):
        """Update the wali signature status label."""
        if not hasattr(self, 'wali_status_label'):
            return
            
        count = len([k for k, v in self.wali_signatures.items() if v])
        total_wali = len(self.guru_data) if self.guru_data else 0
        total = total_wali + (1 if self.kepsek_data else 0)
        
        if count > 0:
            self.wali_status_label.config(
                text=f"✓ {count} dari {total} TTD (wali kelas + kepsek) sudah diatur",
                fg=COLOR_SUCCESS
            )
        elif total > 0:
            self.wali_status_label.config(
                text=f"Ditemukan {total_wali} wali kelas + kepala sekolah (klik tombol di atas untuk TTD)",
                fg=COLOR_TEXT_SECONDARY
            )
        else:
            self.wali_status_label.config(
                text="Pilih file nilai siswa terlebih dahulu untuk melihat daftar guru",
                fg=COLOR_TEXT_SECONDARY
            )
    
    def _open_wali_signature_manager(self):
        """Open popup window to manage wali kelas signatures."""
        # Check if nilai file is selected first
        if "nilai_siswa" not in self.file_selectors or not self.file_selectors["nilai_siswa"].get_path():
            messagebox.showwarning(
                "File Nilai Diperlukan",
                "Pilih file nilai siswa terlebih dahulu untuk melihat daftar wali kelas."
            )
            return
        
        # Load guru data if not loaded yet
        nilai_path = self.file_selectors["nilai_siswa"].get_path()
        if not self.guru_data and not self.kepsek_data:
            self._load_guru_data(nilai_path)
        
        if not self.guru_data and not self.kepsek_data:
            messagebox.showwarning(
                "Data Guru Tidak Ditemukan",
                "Tidak dapat menemukan data guru/wali kelas di file nilai.\n\nPastikan file nilai memiliki sheet 'GURU' dengan kolom 'KELAS' dan 'NAMA'."
            )
            return
        
        # Create popup dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Kelola Tanda Tangan Guru & Kepala Sekolah")
        dialog.geometry("640x500")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 640) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 500) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Title
        tk.Label(
            dialog,
            text="Tanda Tangan Guru & Kepala Sekolah",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=(15, 5))
        
        tk.Label(
            dialog,
            text="Pilih file gambar tanda tangan untuk kepala sekolah dan masing-masing wali kelas",
            font=("Segoe UI", 9),
            fg=COLOR_TEXT_SECONDARY
        ).pack(pady=(0, 10))
        
        # Scrollable frame for list
        canvas_frame = tk.Frame(dialog)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))
        
        canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Store entry widgets for later access
        path_vars = {}
        
        def _add_signature_row(parent, role_label, nama, sig_key, bg_color="white"):
            """Helper to build one signature row."""
            row_frame = tk.Frame(parent, bg=bg_color, relief="groove", bd=1)
            row_frame.pack(fill=tk.X, pady=3, padx=5)
            
            tk.Label(
                row_frame,
                text=role_label,
                font=("Segoe UI", 10, "bold"),
                bg=bg_color,
                width=14,
                anchor="w"
            ).pack(side=tk.LEFT, padx=(10, 5), pady=8)
            
            tk.Label(
                row_frame,
                text=nama,
                font=("Segoe UI", 9),
                bg=bg_color,
                width=22,
                anchor="w"
            ).pack(side=tk.LEFT, padx=5, pady=8)
            
            path_var = tk.StringVar(value=self.wali_signatures.get(sig_key, ""))
            path_vars[sig_key] = path_var
            
            tk.Entry(
                row_frame,
                textvariable=path_var,
                font=("Segoe UI", 8),
                state="readonly",
                readonlybackground="#f1f5f9",
                width=20
            ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=8)
            
            def make_browse(k, pv):
                def browse():
                    fp = filedialog.askopenfilename(
                        title=f"Pilih TTD untuk {k}",
                        filetypes=(("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*"))
                    )
                    if fp:
                        pv.set(fp)
                return browse
            
            tk.Button(
                row_frame,
                text="Browse",
                font=("Segoe UI", 8),
                bg=COLOR_PRIMARY,
                fg="white",
                relief="flat",
                padx=8,
                pady=2,
                command=make_browse(sig_key, path_var)
            ).pack(side=tk.LEFT, padx=(5, 4), pady=8)
            
            def make_clear(pv):
                def clear(): pv.set("")
                return clear
            
            tk.Button(
                row_frame,
                text="×",
                font=("Segoe UI", 10, "bold"),
                bg="#dc2626",
                fg="white",
                relief="flat",
                width=2,
                command=make_clear(path_var)
            ).pack(side=tk.LEFT, padx=(0, 10), pady=8)
        
        # ── Kepala Sekolah row (at top, highlighted) ─────────────────────────
        if self.kepsek_data:
            tk.Label(
                scrollable_frame,
                text="Kepala Sekolah",
                font=("Segoe UI", 9, "bold"),
                fg=COLOR_TEXT_SECONDARY,
                anchor="w"
            ).pack(fill=tk.X, padx=8, pady=(6, 0))
            _add_signature_row(
                scrollable_frame,
                role_label="Kepala Sekolah",
                nama=self.kepsek_data['nama'],
                sig_key="KEPSEK",
                bg_color="#eff6ff"
            )
        
        # ── Wali Kelas rows ───────────────────────────────────────────────────
        if self.guru_data:
            tk.Label(
                scrollable_frame,
                text="Wali Kelas",
                font=("Segoe UI", 9, "bold"),
                fg=COLOR_TEXT_SECONDARY,
                anchor="w"
            ).pack(fill=tk.X, padx=8, pady=(10, 0))
            for guru in self.guru_data:
                _add_signature_row(
                    scrollable_frame,
                    role_label=f"Kelas {guru['kelas']}",
                    nama=guru['nama'],
                    sig_key=guru['kelas']
                )
        
        # Buttons frame
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        def save_and_close():
            # Save all paths to wali_signatures
            for kelas, path_var in path_vars.items():
                path = path_var.get()
                if path:
                    self.wali_signatures[kelas] = path
                elif kelas in self.wali_signatures:
                    del self.wali_signatures[kelas]
            
            self._update_wali_status_label()
            self._log(f"TTD wali kelas diperbarui: {len(self.wali_signatures)} file", "success")
            dialog.destroy()
        
        tk.Button(
            btn_frame,
            text="💾 Simpan",
            font=("Segoe UI", 10),
            bg=COLOR_SUCCESS,
            fg="white",
            padx=20,
            pady=5,
            relief="flat",
            command=save_and_close
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="Batal",
            font=("Segoe UI", 10),
            padx=20,
            pady=5,
            command=dialog.destroy
        ).pack(side=tk.LEFT, padx=5)
    
    def _download_template(self):
        """Download/copy template files with selection dialog."""
        # Create template selection dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Download Template")
        dialog.geometry("400x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        
        # Center dialog first, then grab
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 400) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 350) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Wait for dialog to be visible before grabbing
        dialog.after(100, lambda: dialog.grab_set())
        
        tk.Label(
            dialog,
            text="Pilih Template yang Ingin Didownload:",
            font=("Segoe UI", 12, "bold")
        ).pack(pady=(20, 15))
        
        # Checkboxes for each template
        selected_templates = {}
        for key, config in GENERATORS.items():
            if config.get("disabled"):
                continue
            var = tk.BooleanVar(value=True)
            selected_templates[key] = var
            cb = tk.Checkbutton(
                dialog,
                text=config["name"],
                variable=var,
                font=("Segoe UI", 11)
            )
            cb.pack(anchor="w", padx=40, pady=3)
        
        def do_download():
            dialog.destroy()
            
            # Ask for destination folder
            dest_folder = filedialog.askdirectory(title="Pilih Folder Tujuan untuk Template")
            if not dest_folder:
                return
            
            copied_files = []
            errors = []
            
            # Copy selected template files
            for gen_key, var in selected_templates.items():
                if not var.get():
                    continue
                
                # Get template files from GENERATORS config
                config = GENERATORS.get(gen_key, {})
                template_files = config.get("template_files", [])
                
                for template_file in template_files:
                    # Use get_resource_path to support both dev and executable
                    src_path = get_resource_path(template_file)
                    if os.path.exists(src_path):
                        gen_folder = os.path.join(dest_folder, f"template_{gen_key}")
                        os.makedirs(gen_folder, exist_ok=True)
                        
                        dest_path = os.path.join(gen_folder, os.path.basename(template_file))
                        try:
                            shutil.copy2(src_path, dest_path)
                            copied_files.append(os.path.basename(template_file))
                        except Exception as e:
                            errors.append(f"{template_file}: {e}")
                    else:
                        errors.append(f"{template_file}: File tidak ditemukan")
            
            if copied_files:
                msg = f"✅ {len(copied_files)} file template berhasil di-copy!\n\nLokasi: {dest_folder}"
                if errors:
                    msg += f"\n\n⚠️ Beberapa file gagal:\n" + "\n".join(errors)
                messagebox.showinfo("Template Downloaded", msg)
            else:
                messagebox.showerror("Error", "Tidak ada file template yang bisa di-copy.\n\n" + "\n".join(errors))
        
        # Buttons
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        tk.Button(
            btn_frame,
            text="Download",
            font=("Segoe UI", 10),
            bg=COLOR_SUCCESS,
            fg="white",
            padx=20,
            pady=5,
            command=do_download
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="Batal",
            font=("Segoe UI", 10),
            padx=20,
            pady=5,
            command=dialog.destroy
        ).pack(side=tk.LEFT, padx=5)
    
    def _browse_output(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(title="Pilih Folder Output")
        if folder:
            self.output_path.set(folder)
    
    def _clear_files(self):
        """Clear all file selections."""
        for selector in self.file_selectors.values():
            selector.clear()
        self._clear_log()
        self._log("File dan log dibersihkan", "info")
    
    def _validate_files(self) -> bool:
        """Validate that required files are selected."""
        if not self.selected_generator:
            messagebox.showwarning("Peringatan", "Pilih jenis dokumen terlebih dahulu!")
            return False
        
        config = GENERATORS[self.selected_generator]
        for file_config in config["files_needed"]:
            if file_config.get("required"):
                selector = self.file_selectors.get(file_config["key"])
                if not selector or not selector.get_path():
                    messagebox.showwarning(
                        "File Diperlukan",
                        f"Mohon pilih file untuk: {file_config['label']}"
                    )
                    return False
        
        # Validate output folder is selected
        if not self.output_path.get():
            messagebox.showwarning(
                "Folder Output Diperlukan",
                "Mohon pilih folder output terlebih dahulu!"
            )
            return False
        
        return True
    
    def _validate_excel_format(self) -> tuple:
        """Validate Excel file format matches expected structure.
        Returns (is_valid, error_message)
        """
        import pandas as pd
        
        generator_key = self.selected_generator
        
        # Create loading popup
        loading_dialog = tk.Toplevel(self.root)
        loading_dialog.title("Memuat File")
        loading_dialog.geometry("350x120")
        loading_dialog.resizable(False, False)
        loading_dialog.transient(self.root)
        loading_dialog.configure(bg=COLOR_BG)
        
        # Center dialog
        loading_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 350) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 120) // 2
        loading_dialog.geometry(f"+{x}+{y}")
        
        tk.Label(
            loading_dialog,
            text="📂 Sedang membaca file Excel...",
            font=("Segoe UI", 12, "bold"),
            bg=COLOR_BG,
            fg=COLOR_TEXT
        ).pack(pady=(20, 10))
        
        tk.Label(
            loading_dialog,
            text="Harap tunggu, jangan tutup aplikasi",
            font=("Segoe UI", 9),
            bg=COLOR_BG,
            fg=COLOR_TEXT_SECONDARY
        ).pack(pady=(0, 10))
        
        progress = ttk.Progressbar(
            loading_dialog,
            mode="indeterminate",
            length=280
        )
        progress.pack(pady=(0, 15))
        progress.start(10)
        
        loading_dialog.update()
        
        try:
            # Only check if files can be read as Excel files
            # PDF generators will handle their own validation
            if generator_key == "skhupk":
                excel_path = self.file_selectors["excel_skhupk"].get_path()
                pd.read_excel(excel_path, sheet_name=None)  # Just check if readable
                    
            elif generator_key == "transkrip":
                excel_path = self.file_selectors["excel_transkrip"].get_path()
                pd.read_excel(excel_path, sheet_name=None)  # Just check if readable
                    
            elif generator_key == "rapot":
                file1_path = self.file_selectors["data_siswa"].get_path()
                file2_path = self.file_selectors["nilai_siswa"].get_path()
                
                data_df = pd.read_excel(file1_path, sheet_name=None)
                nilai_df = pd.read_excel(file2_path, sheet_name=None)
                
                # Normalize sheet names: uppercase, strip, remove spaces
                def normalize_sheet(name):
                    return "".join(name.upper().strip().split())
                
                data_sheets = {normalize_sheet(k): k for k in data_df.keys()}
                nilai_sheets = {normalize_sheet(k): k for k in nilai_df.keys()}
                
                # Find KP class sheets in nilai file
                nilai_kp = {k: v for k, v in nilai_sheets.items() if "KP" in k}
                data_kp  = {k: v for k, v in data_sheets.items() if "KP" in k}
                
                matched = []
                missing = []
                for norm_key, orig_name in nilai_kp.items():
                    if norm_key in data_kp:
                        matched.append(f"{orig_name} ↔ {data_kp[norm_key]}")
                    else:
                        missing.append(orig_name)
                
                self._log("File Excel dapat dibaca ✓", "success")
                
                if matched:
                    self._log(f"Sheet cocok ({len(matched)}): {', '.join(matched)}", "success")
                if missing:
                    self._log(
                        f"Sheet nilai tidak ditemukan di Data Siswa: {', '.join(missing)}",
                        "warning"
                    )
                
                extra_data = [v for k, v in data_kp.items() if k not in nilai_kp]
                if extra_data:
                    self._log(
                        f"Sheet di Data Siswa tidak ada di Nilai: {', '.join(extra_data)}",
                        "warning"
                    )
            
            elif generator_key == "kartu_upk":
                excel_path = self.file_selectors["excel_kartu"].get_path()
                df_check = pd.read_excel(excel_path, sheet_name=None)
                if 'DataSiswa' not in df_check:
                    loading_dialog.destroy()
                    return False, "Sheet 'DataSiswa' tidak ditemukan di file Excel.\nPastikan file memiliki sheet DataSiswa dan DataTetap."
                self._log(f"File Excel valid - {len(df_check.get('DataSiswa', []))} siswa ditemukan", "success")
            
            elif generator_key == "kartu_siswa":
                excel_path = self.file_selectors["excel_kartu_siswa"].get_path()
                df_check = pd.read_excel(excel_path, sheet_name=None)
                if 'DataSiswa' not in df_check:
                    loading_dialog.destroy()
                    return False, "Sheet 'DataSiswa' tidak ditemukan di file Excel.\nPastikan file memiliki sheet DataSiswa dan DataTetap."
                self._log(f"File Excel valid - {len(df_check.get('DataSiswa', []))} siswa ditemukan", "success")
            
            loading_dialog.destroy()
            return True, ""
            
        except Exception as e:
            loading_dialog.destroy()
            return False, f"Error membaca file Excel: {str(e)}\nPastikan file adalah format Excel yang valid (.xlsx)"
    
    def _toggle_generate(self):
        """Toggle between generate and stop."""
        if self.is_generating:
            # Currently generating - show stop confirmation
            response = messagebox.askyesno(
                "Konfirmasi Stop",
                "Yakin ingin menghentikan proses generate?"
            )
            if response:
                self.stop_requested = True
                self._log("Menghentikan proses generate...", "warning")
        else:
            # Start generation
            self._generate()
    
    def _generate(self):
        """Generate PDF documents."""
        if not self._validate_files():
            return
        
        # Clear log and start fresh
        self._clear_log()
        self._log("Memulai proses generate...", "info")
        
        # Validate Excel format first
        self._log("Memeriksa format file Excel...", "info")
        is_valid, error_msg = self._validate_excel_format()
        if not is_valid:
            self._log(f"Format file tidak sesuai!", "error")
            self._log(error_msg, "error")
            messagebox.showerror(
                "Format File Tidak Sesuai",
                f"Data yang dimasukkan tidak sesuai template!\n\n{error_msg}"
            )
            return
        
        self._log("Format file valid ✓", "success")
        
        # Reset stop flag
        self.stop_requested = False
        
        # Disable UI during generation
        self._set_ui_enabled(False)
        
        # Change button to Stop
        self.generate_btn.config(
            text="⛔ Stop Generate",
            bg_color="#dc2626",
            hover_color="#b91c1c"
        )
        
        # Show progress
        self.progress.pack(fill=tk.X, pady=(10, 0))
        self.progress.start(10)
        
        # Run generation in thread
        thread = threading.Thread(target=self._run_generation)
        thread.daemon = True
        thread.start()
    
    def _run_generation(self):
        """Run the actual generation process."""
        try:
            generator_key = self.selected_generator
            output_folder = self.output_path.get()
            
            # Create output folder
            os.makedirs(output_folder, exist_ok=True)
            
            self.root.after(0, lambda: self._log(f"Generator: {GENERATORS[generator_key]['name']}", "info"))
            self.root.after(0, lambda: self._log(f"Output folder: {output_folder}", "info"))
            
            if generator_key == "rapot":
                self._generate_rapot(output_folder)
            elif generator_key == "skhupk":
                self._generate_skhupk(output_folder)
            elif generator_key == "kartu_upk":
                self._generate_kartu_upk(output_folder)
            elif generator_key == "kartu_siswa":
                self._generate_kartu_siswa(output_folder)
            elif generator_key == "transkrip":
                self._generate_transkrip(output_folder)
            
            # Success
            self.root.after(0, lambda: self._on_generation_complete(True))
            
        except InterruptedError as e:
            # User stopped generation
            self.root.after(0, lambda: self._on_generation_complete(False, str(e)))
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda err=str(e): self._on_generation_complete(False, err))
    
    def _generate_rapot(self, output_folder: str):
        """Generate rapot PDFs for all students."""
        import sys
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from lib.pdf_generators.rapot_generator import generate_all_rapots
        
        data_siswa_path = self.file_selectors["data_siswa"].get_path()
        nilai_siswa_path = self.file_selectors["nilai_siswa"].get_path()
        
        self.root.after(0, lambda: self._log(f"File data siswa: {os.path.basename(data_siswa_path)}", "info"))
        self.root.after(0, lambda: self._log(f"File nilai: {os.path.basename(nilai_siswa_path)}", "info"))
        
        # Change to app directory for resources
        original_dir = os.getcwd()
        try:
            # For PyInstaller executable
            app_dir = sys._MEIPASS
        except Exception:
            # For normal Python environment
            app_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(app_dir)
        
        try:
            self.root.after(0, lambda: self._log("Memuat data dari Excel...", "info"))
            
            # Get kepsek signature from combined wali_signatures dict (key='KEPSEK')
            ttd_kepsek_path = self.wali_signatures.get('KEPSEK') or None
            if ttd_kepsek_path:
                self.root.after(0, lambda p=ttd_kepsek_path: self._log(f"TTD Kepsek: {os.path.basename(p)}", "info"))
            
            # Get wali kelas signatures dict (exclude KEPSEK key)
            ttd_wali_dict = {k: v for k, v in self.wali_signatures.items() if k != 'KEPSEK'} or None
            if ttd_wali_dict:
                self.root.after(0, lambda: self._log(f"TTD Wali Kelas: {len(ttd_wali_dict)} file", "info"))
            
            # Progress callback for logging and stop checking
            def progress_callback(current, total, kelas, student_name):
                if self.stop_requested:
                    raise InterruptedError("Generate dihentikan oleh user")
                self.root.after(0, lambda c=current, t=total, k=kelas, n=student_name: 
                    self._log(f"[{k}] [{c}/{t}] {n}", "info"))
            
            # Generate all rapots using the helper function
            result = generate_all_rapots(
                data_siswa_path=data_siswa_path,
                nilai_path=nilai_siswa_path,
                output_folder=output_folder,
                ttd_kepsek_path=ttd_kepsek_path,
                ttd_wali_dict=ttd_wali_dict,
                progress_callback=progress_callback
            )
            
            # Handle both dict (new) and list (legacy) return types
            if isinstance(result, dict):
                generated_files = result.get('generated_files', [])
                gen_warnings = result.get('warnings', [])
                gen_errors = result.get('errors', [])
            else:
                generated_files = result
                gen_warnings = []
                gen_errors = []
            
            # Display warnings to user
            for w in gen_warnings:
                self.root.after(0, lambda msg=w: self._log(f"⚠ {msg}", "warning"))
            
            # Display errors to user
            for e in gen_errors:
                self.root.after(0, lambda msg=e: self._log(f"❌ {msg}", "error"))
            
            self.root.after(0, lambda: self._log(f"{len(generated_files)} file rapot berhasil dibuat!", "success"))
            
            if gen_warnings or gen_errors:
                summary_parts = []
                if gen_warnings:
                    summary_parts.append(f"{len(gen_warnings)} warning")
                if gen_errors:
                    summary_parts.append(f"{len(gen_errors)} error")
                self.root.after(0, lambda: self._log(
                    f"Selesai dengan {', '.join(summary_parts)}. Periksa log di atas untuk detail.", "warning"))
        finally:
            os.chdir(original_dir)
    
    def _generate_skhupk(self, output_folder: str):
        """Generate SKHUPK PDFs."""
        import sys
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from lib.pdf_generators.skhupk_generator import generate_all_skhupk
        
        excel_path = self.file_selectors["excel_skhupk"].get_path()
        self.root.after(0, lambda: self._log(f"File Excel: {os.path.basename(excel_path)}", "info"))
        
        # Get optional signature path
        ttd_kepsek_path = None
        if "ttd_kepsek" in self.file_selectors:
            path = self.file_selectors["ttd_kepsek"].get_path()
            if path:
                ttd_kepsek_path = path
                self.root.after(0, lambda: self._log(f"TTD Kepsek: {os.path.basename(path)}", "info"))
        
        # Change to app directory for resources (logo, fonts)
        original_dir = os.getcwd()
        app_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(app_dir)
        
        try:
            # Progress callback for logging
            def progress_callback(current, total, student_name):
                self.root.after(0, lambda c=current, t=total, n=student_name: 
                    self._log(f"[{c}/{t}] Generating: {n}", "info"))
            
            # Generate all SKHUPK using the wrapper function
            generated_files = generate_all_skhupk(
                excel_file_path=excel_path,
                output_folder=output_folder,
                ttd_kepsek_path=ttd_kepsek_path,
                progress_callback=progress_callback
            )
            
            self.root.after(0, lambda: self._log(f"{len(generated_files)} file SKHUPK berhasil dibuat!", "success"))
        finally:
            os.chdir(original_dir)
    
    def _generate_kartu_upk(self, output_folder: str):
        """Generate Kartu UPK PDFs."""
        import sys
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from lib.pdf_generators.kartu_upk_generator import generate_all_kartu_upk
        
        excel_path = self.file_selectors["excel_kartu"].get_path()
        self.root.after(0, lambda: self._log(f"File Excel: {os.path.basename(excel_path)}", "info"))
        
        # Get optional photo folder
        photo_folder = None
        if "foto_folder_upk" in self.file_selectors:
            path = self.file_selectors["foto_folder_upk"].get_path()
            if path:
                photo_folder = path
                self.root.after(0, lambda p=path: self._log(f"Folder Foto: {p}", "info"))
        
        # Load tahun ajaran from Excel DataTetap
        tahun_ajaran = None
        try:
            import pandas as pd
            import warnings
            warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
            
            df = pd.read_excel(excel_path, sheet_name="DataTetap")
            df.columns = df.columns.str.strip()
            
            for _, row in df.iterrows():
                var_name = str(row.iloc[0]).strip()
                if var_name.lower() == 'tahun':
                    tahun_ajaran = str(row.iloc[1]).strip() if len(row) > 1 else None
                    break
            
            if tahun_ajaran:
                self.root.after(0, lambda t=tahun_ajaran: self._log(f"Tahun Ajaran: {t} (dari Excel)", "info"))
        except Exception as e:
            self.root.after(0, lambda: self._log(f"Warning: Tidak dapat load Tahun Ajaran dari Excel", "warning"))
        
        # Get kepsek signature (from wali_signatures if available)
        ttd_kepsek_path = self.wali_signatures.get('KEPSEK') or None
        if ttd_kepsek_path:
            self.root.after(0, lambda p=ttd_kepsek_path: self._log(f"TTD Kepsek: {os.path.basename(p)}", "info"))
        
        # Change to app directory for resources
        original_dir = os.getcwd()
        app_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(app_dir)
        
        try:
            self.root.after(0, lambda: self._log("Memuat data dari Excel...", "info"))
            
            def progress_callback(current, total, student_name):
                if self.stop_requested:
                    raise InterruptedError("Generate dihentikan oleh user")
                self.root.after(0, lambda c=current, t=total, n=student_name: 
                    self._log(f"[{c}/{t}] Generating: {n}", "info"))
            
            generated_files = generate_all_kartu_upk(
                excel_file_path=excel_path,
                output_folder=output_folder,
                photo_folder=photo_folder,
                tahun_ajaran=tahun_ajaran,
                ttd_kepsek_path=ttd_kepsek_path,
                progress_callback=progress_callback
            )
            
            self.root.after(0, lambda: self._log(f"{len(generated_files)} file Kartu UPK berhasil dibuat!", "success"))
        finally:
            os.chdir(original_dir)
    
    def _generate_kartu_siswa(self, output_folder: str):
        """Generate Kartu Siswa/Pelajar PDFs."""
        import sys
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from lib.pdf_generators.kartu_siswa_generator import generate_all_kartu_siswa
        
        excel_path = self.file_selectors["excel_kartu_siswa"].get_path()
        self.root.after(0, lambda: self._log(f"File Excel: {os.path.basename(excel_path)}", "info"))
        
        # Get optional photo folder
        photo_folder = None
        if "foto_folder_siswa" in self.file_selectors:
            path = self.file_selectors["foto_folder_siswa"].get_path()
            if path:
                photo_folder = path
                self.root.after(0, lambda p=path: self._log(f"Folder Foto: {p}", "info"))
        
        # Load tahun ajaran from Excel DataTetap
        tahun_ajaran = None
        try:
            import pandas as pd
            import warnings
            warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
            
            df = pd.read_excel(excel_path, sheet_name="DataTetap")
            df.columns = df.columns.str.strip()
            
            for _, row in df.iterrows():
                var_name = str(row.iloc[0]).strip()
                if var_name.lower() == 'tahun':
                    tahun_ajaran = str(row.iloc[1]).strip() if len(row) > 1 else None
                    break
            
            if tahun_ajaran:
                self.root.after(0, lambda t=tahun_ajaran: self._log(f"Tahun Ajaran: {t} (dari Excel)", "info"))
        except Exception as e:
            self.root.after(0, lambda: self._log(f"Warning: Tidak dapat load Tahun Ajaran dari Excel", "warning"))
        
        # Get kepsek signature (from wali_signatures if available)
        ttd_kepsek_path = self.wali_signatures.get('KEPSEK') or None
        if ttd_kepsek_path:
            self.root.after(0, lambda p=ttd_kepsek_path: self._log(f"TTD Kepsek: {os.path.basename(p)}", "info"))
        
        # Change to app directory for resources
        original_dir = os.getcwd()
        app_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(app_dir)
        
        try:
            self.root.after(0, lambda: self._log("Memuat data dari Excel...", "info"))
            
            def progress_callback(current, total, student_name):
                if self.stop_requested:
                    raise InterruptedError("Generate dihentikan oleh user")
                self.root.after(0, lambda c=current, t=total, n=student_name: 
                    self._log(f"[{c}/{t}] Generating: {n}", "info"))
            
            generated_files = generate_all_kartu_siswa(
                excel_file_path=excel_path,
                output_folder=output_folder,
                photo_folder=photo_folder,
                tahun_ajaran=tahun_ajaran,
                ttd_kepsek_path=ttd_kepsek_path,
                progress_callback=progress_callback
            )
            
            self.root.after(0, lambda: self._log(f"{len(generated_files)} file Kartu Pelajar berhasil dibuat!", "success"))
        finally:
            os.chdir(original_dir)
    
    def _generate_transkrip(self, output_folder: str):
        """Generate Transkrip Nilai PDFs."""
        import sys
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from lib.pdf_generators.transcript_generator import generate_all_transcripts
        
        excel_path = self.file_selectors["excel_transkrip"].get_path()
        self.root.after(0, lambda: self._log(f"File Excel: {os.path.basename(excel_path)}", "info"))
        
        # Get optional signature path
        ttd_kepsek_path = None
        if "ttd_kepsek" in self.file_selectors:
            path = self.file_selectors["ttd_kepsek"].get_path()
            if path:
                ttd_kepsek_path = path
                self.root.after(0, lambda: self._log(f"TTD Kepsek: {os.path.basename(path)}", "info"))
        
        # Change to app directory (for font and other resources)
        original_dir = os.getcwd()
        app_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(app_dir)
        
        try:
            self.root.after(0, lambda: self._log("Memuat data dari Excel...", "info"))
            
            # Progress callback for logging
            def progress_callback(current, total, student_name):
                self.root.after(0, lambda c=current, t=total, n=student_name: 
                    self._log(f"[{c}/{t}] Generating: {n}", "info"))
            
            # Generate all transcripts using the wrapper function
            generated_files = generate_all_transcripts(
                excel_file_path=excel_path,
                output_folder=output_folder,
                ttd_kepsek_path=ttd_kepsek_path,
                progress_callback=progress_callback
            )
            
            self.root.after(0, lambda: self._log(f"{len(generated_files)} file transkrip berhasil dibuat!", "success"))
        finally:
            os.chdir(original_dir)
    
    def _on_generation_complete(self, success: bool, error: str = None):
        """Handle generation completion."""
        self.progress.stop()
        self.progress.pack_forget()
        
        # Reset button to Generate (green)
        self.generate_btn.config(
            text="🚀 Generate ",
            bg_color=COLOR_SUCCESS,
            hover_color=COLOR_SUCCESS_HOVER
        )
        
        # Re-enable UI
        self._set_ui_enabled(True)
        self.stop_requested = False
        
        if success:
            self._log(f"Selesai! File tersimpan di: {self.output_path.get()}", "success")
            messagebox.showinfo(
                "Berhasil",
                f"Dokumen berhasil di-generate!\n\nLokasi: {self.output_path.get()}"
            )
        else:
            self._log(f"Error: {error}", "error")
            messagebox.showerror("Error", f"Gagal generate dokumen:\n\n{error}")


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================
def main():
    # Create main window first (hidden)
    root = tk.Tk()
    root.withdraw()
    
    # Create splash screen as Toplevel
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg=COLOR_BG)
    
    # Get screen dimensions
    screen_width = splash.winfo_screenwidth()
    screen_height = splash.winfo_screenheight()
    
    # Splash dimensions
    splash_width = 400
    splash_height = 300
    x = (screen_width - splash_width) // 2
    y = (screen_height - splash_height) // 2
    
    splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
    
    # Main frame
    frame = tk.Frame(splash, bg=COLOR_BG, highlightthickness=2, highlightbackground=COLOR_PRIMARY)
    frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
    
    # Logo
    logo_path = get_resource_path(os.path.join("assets", "logo.png"))
    if PIL_AVAILABLE and os.path.exists(logo_path):
        try:
            img = Image.open(logo_path)
            img = img.resize((120, 120), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            logo_label = tk.Label(frame, image=photo, bg=COLOR_BG)
            logo_label.image = photo
            logo_label.pack(pady=(30, 10))
        except:
            pass
    
    # Title
    tk.Label(
        frame,
        text="PKBM PPI TAIWAN",
        font=("Segoe UI", 18, "bold"),
        bg=COLOR_BG,
        fg=COLOR_TEXT
    ).pack(pady=(10, 5))
    
    tk.Label(
        frame,
        text="Generator Dokumen Pendidikan",
        font=("Segoe UI", 10),
        bg=COLOR_BG,
        fg=COLOR_TEXT_SECONDARY
    ).pack(pady=(0, 20))
    
    # Loading message
    loading_label = tk.Label(
        frame,
        text="Memuat aplikasi...",
        font=("Segoe UI", 9),
        bg=COLOR_BG,
        fg=COLOR_TEXT_SECONDARY
    )
    loading_label.pack(pady=(10, 10))
    
    # Progress bar
    progress = ttk.Progressbar(
        frame,
        mode="indeterminate",
        length=300
    )
    progress.pack(pady=(0, 30))
    progress.start(10)
    
    # Version
    tk.Label(
        frame,
        text=f"v{APP_VERSION}",
        font=("Segoe UI", 8),
        bg=COLOR_BG,
        fg=COLOR_TEXT_SECONDARY
    ).pack(side=tk.BOTTOM, pady=(0, 10))
    
    # Update splash to show it
    splash.update()

    # Load app in background while splash is showing
    def load_app():
        # Set window icon using logo.png
        logo_path = get_resource_path(os.path.join("assets", "logo.png"))
        if PIL_AVAILABLE and os.path.exists(logo_path):
            try:
                # Load and set icon from PNG
                img = Image.open(logo_path)
                photo = ImageTk.PhotoImage(img)
                root.iconphoto(True, photo)
            except Exception as e:
                print(f"Could not set icon: {e}")
        
        # Initialize app (heavy imports happen here)
        app = PKBMGeneratorApp(root)
        
        # Close splash and show main window
        splash.destroy()
        root.deiconify()
        
        return app
    
    # Schedule app loading after splash is visible
    app_instance = [None]  # Use list to store app reference for closure
    
    def load_and_setup():
        app = load_app()
        app_instance[0] = app
        
        # Setup window close handler after app is loaded
        def on_closing():
            if app.is_generating:
                # Show confirmation dialog with two options
                response = messagebox.askyesnocancel(
                    "Konfirmasi Keluar",
                    "Proses generate sedang berjalan!\n\n"
                    "• Klik YES untuk STOP generate dan KELUAR aplikasi\n"
                    "• Klik NO untuk STOP generate tapi TETAP di aplikasi\n"
                    "• Klik CANCEL untuk LANJUTKAN generate"
                )
                
                if response is True:  # YES - stop and exit
                    app.stop_requested = True
                    # Wait a bit for thread to detect stop, then force exit
                    root.after(500, root.destroy)
                elif response is False:  # NO - stop but stay
                    app.stop_requested = True
                    # Generation thread will catch InterruptedError and call _on_generation_complete
                # else: CANCEL - do nothing, continue generating
            else:
                root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
    
    root.after(100, load_and_setup)
    root.mainloop()


if __name__ == "__main__":
    main()
