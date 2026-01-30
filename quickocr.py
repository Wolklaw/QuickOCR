import sys
import os
import json
import ctypes
import tkinter as tk
from tkinter import messagebox
from typing import Optional, Dict, Any

import pytesseract
import pyperclip
from PIL import ImageGrab, Image, ImageOps

APP_NAME = "QuickOCR"
VERSION = "1.0.1"
CONFIG_FILENAME = "config.json"

THEME = {
    "bg_main":    "#1e1e1e",
    "bg_header":  "#252526",
    "fg_text":    "#cccccc",
    "accent":     "#007acc",
    "accent_hov": "#0098ff",
    "close_hov":  "#e81123",
    "min_hov":    "#3e3e42",
    "success":    "#4cc790",
    "instruction":"#888888"
}

GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080

def enable_high_dpi_awareness():
    if os.name != 'nt': return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

def get_resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def force_taskbar_visibility(root_window: tk.Tk):
    try:
        hwnd = ctypes.windll.user32.GetParent(root_window.winfo_id())
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        style = style & ~WS_EX_TOOLWINDOW
        style = style | WS_EX_APPWINDOW
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
        
        root_window.withdraw()
        root_window.after(10, root_window.deiconify)
    except Exception:
        pass

class ConfigManager:
    @staticmethod
    def _get_path() -> str:
        appdata = os.getenv('APPDATA')
        folder = os.path.join(appdata, APP_NAME)
        if not os.path.exists(folder):
            os.makedirs(folder)
        return os.path.join(folder, CONFIG_FILENAME)

    @staticmethod
    def load() -> Dict[str, Any]:
        try:
            path = ConfigManager._get_path()
            if os.path.exists(path):
                with open(path, 'r') as f:
                    return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            pass
        return {}

    @staticmethod
    def save(data: Dict[str, Any]):
        try:
            with open(ConfigManager._get_path(), 'w') as f:
                json.dump(data, f)
        except IOError:
            pass

class OCREngine:
    def __init__(self):
        self.tesseract_cmd = get_resource_path(os.path.join('Tesseract-OCR', 'tesseract.exe'))
        if sys.platform.startswith('win'):
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd

    def extract_text(self, img: Image.Image) -> str:
        w, h = img.size
        img = img.resize((w * 3, h * 3), Image.Resampling.LANCZOS)
        img = img.convert('L')
        img = ImageOps.invert(img)
        img = img.point(lambda x: 0 if x < 140 else 255, '1')
        return pytesseract.image_to_string(img, lang='eng+fra', config='--psm 6')

class SnippingOverlay(tk.Toplevel):
    def __init__(self, parent, on_complete):
        super().__init__(parent)
        self.on_complete = on_complete
        self.start_pos = None
        self.rect = None

        user32 = ctypes.windll.user32
        self.v_width = user32.GetSystemMetrics(78)
        self.v_height = user32.GetSystemMetrics(79)
        self.v_x = user32.GetSystemMetrics(76)
        self.v_y = user32.GetSystemMetrics(77)

        self.geometry(f"{self.v_width}x{self.v_height}+{self.v_x}+{self.v_y}")
        self.overrideredirect(True)
        self.attributes('-alpha', 0.3)
        self.attributes('-topmost', True)
        self.configure(bg="black")

        self.canvas = tk.Canvas(self, cursor="cross", bg="grey11", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.bind('<Escape>', lambda e: self.destroy())

    def _on_press(self, event):
        self.start_pos = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.rect = self.canvas.create_rectangle(
            *self.start_pos, *self.start_pos, 
            outline=THEME['accent'], width=2
        )

    def _on_drag(self, event):
        cur_x, cur_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_pos[0], self.start_pos[1], cur_x, cur_y)

    def _on_release(self, event):
        if not self.start_pos: return
        
        x1, y1 = self.start_pos
        x2, y2 = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.destroy()
        
        x_min, x_max = min(x1, x2), max(x1, x2)
        y_min, y_max = min(y1, y2), max(y1, y2)

        if (x_max - x_min) < 5 or (y_max - y_min) < 5:
            self.on_complete(None)
            return

        # Map logical coordinates to physical virtual screen coordinates
        capture_bbox = (
            int(x_min + self.v_x),
            int(y_min + self.v_y),
            int(x_max + self.v_x),
            int(y_max + self.v_y)
        )
        
        try:
            img = ImageGrab.grab(bbox=capture_bbox, all_screens=True)
            self.on_complete(img)
        except Exception as e:
            messagebox.showerror("Capture Error", str(e))
            self.on_complete(None)

class ResultPopup(tk.Toplevel):
    def __init__(self, parent, text: str, timeout: int = 10):
        super().__init__(parent)
        self.overrideredirect(True)
        self.configure(bg=THEME["bg_main"])
        self.attributes('-topmost', True)
        
        w, h = 400, 240
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        
        frame = tk.Frame(self, bg=THEME["bg_main"], highlightbackground=THEME["accent"], highlightthickness=1)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            frame, text="✓ COPIED TO CLIPBOARD", 
            font=("Segoe UI", 11, "bold"), bg=THEME["bg_main"], fg=THEME["success"]
        ).pack(pady=(15, 2))

        tk.Label(
            frame, text="Press Ctrl+V to paste", 
            font=("Segoe UI", 9), bg=THEME["bg_main"], fg=THEME["instruction"]
        ).pack(pady=(0, 10))

        preview = text.replace('\n', ' ')
        if len(preview) > 150:
            preview = preview[:150] + "..."
            
        tk.Label(
            frame, text=preview, font=("Consolas", 9), 
            bg=THEME["bg_header"], fg=THEME["fg_text"], 
            wraplength=360, justify="left", padx=10, pady=10
        ).pack(fill=tk.X, padx=20)

        self.lbl_timer = tk.Label(
            frame, text=f"Auto-closing in {timeout}s", 
            font=("Segoe UI", 8), bg=THEME["bg_main"], fg="#666"
        )
        self.lbl_timer.pack(pady=(10, 5))

        tk.Button(
            frame, text="OK", command=self.destroy, 
            bg=THEME["accent"], fg="white", 
            activebackground=THEME["accent_hov"], activeforeground="white",
            bd=0, padx=25, pady=4, cursor="hand2"
        ).pack(pady=10)

        self._start_timer(timeout)

    def _start_timer(self, remaining: int):
        if remaining > 0:
            self.lbl_timer.config(text=f"Closing in {remaining}s")
            self.after(1000, lambda: self._start_timer(remaining - 1))
        else:
            self.destroy()

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.configure(bg=THEME["bg_main"])
        
        try:
            icon_path = get_resource_path("aa.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass 

        self.ocr = OCREngine()
        self.config = ConfigManager.load()
        
        self._setup_window_geometry()
        self._build_custom_titlebar()
        self._build_main_ui()
        
        self.root.after(10, lambda: force_taskbar_visibility(self.root))

    def _setup_window_geometry(self):
        w, h = 300, 150
        x = self.config.get("x", (self.root.winfo_screenwidth() - w) // 2)
        y = self.config.get("y", (self.root.winfo_screenheight() - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_custom_titlebar(self):
        self.title_bar = tk.Frame(self.root, bg=THEME["bg_header"], height=30)
        self.title_bar.pack(fill=tk.X, side=tk.TOP)
        self.title_bar.pack_propagate(False)

        self.title_bar.bind("<Button-1>", self._start_move)
        self.title_bar.bind("<B1-Motion>", self._do_move)

        tk.Label(
            self.title_bar, text=f"{APP_NAME}", 
            bg=THEME["bg_header"], fg=THEME["fg_text"], font=("Segoe UI", 9, "bold")
        ).pack(side=tk.LEFT, padx=10)

        self._create_titlebar_btn("✕", self._on_close, THEME["close_hov"])
        self._create_titlebar_btn("—", self._minimize, THEME["min_hov"])

    def _create_titlebar_btn(self, text, command, hover_color):
        btn = tk.Button(
            self.title_bar, text=text, command=command,
            bg=THEME["bg_header"], fg=THEME["fg_text"],
            bd=0, font=("Arial", 9, "bold"), width=4,
            activebackground=hover_color, activeforeground="white"
        )
        btn.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_main_ui(self):
        main_frame = tk.Frame(self.root, bg=THEME["bg_main"])
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.bind("<Button-1>", self._start_move)
        main_frame.bind("<B1-Motion>", self._do_move)

        btn = tk.Button(
            main_frame, text="CAPTURE ZONE", command=self._start_snip,
            font=("Segoe UI", 10, "bold"), bg=THEME["accent"], fg="white",
            activebackground=THEME["accent_hov"], activeforeground="white",
            bd=0, cursor="hand2", padx=20, pady=10
        )
        btn.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    def _start_move(self, event):
        self.last_x = event.x
        self.last_y = event.y

    def _do_move(self, event):
        deltax = event.x - self.last_x
        deltay = event.y - self.last_y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    def _minimize(self):
        self.root.overrideredirect(False) 
        self.root.iconify()
        self.root.bind("<Map>", self._restore_window)

    def _restore_window(self, event):
        if self.root.state() == 'normal':
            self.root.overrideredirect(True)
            self.root.unbind("<Map>")
            force_taskbar_visibility(self.root)

    def _on_close(self):
        self.config["x"] = self.root.winfo_x()
        self.config["y"] = self.root.winfo_y()
        ConfigManager.save(self.config)
        self.root.destroy()

    def _start_snip(self):
        self.root.withdraw()
        self.root.after(150, lambda: SnippingOverlay(self.root, self._process_snip))

    def _process_snip(self, img: Optional[Image.Image]):
        self.root.deiconify()
        self.root.overrideredirect(True)
        force_taskbar_visibility(self.root)

        if not img: return

        try:
            text = self.ocr.extract_text(img)
            if text and text.strip():
                pyperclip.copy(text)
                ResultPopup(self.root, text)
            else:
                messagebox.showwarning(APP_NAME, "No text found.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    enable_high_dpi_awareness()
    app = App()
    app.run()