import sys
import os
import json
import ctypes
import tkinter as tk
from tkinter import messagebox
import pytesseract
import pyperclip
from PIL import ImageGrab, Image, ImageOps

APP_NAME = "QuickOCR"
VERSION = "1.0.0"

# Win32 API Constants
GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080

THEME = {
    "bg_main":    "#1e1e1e",
    "bg_header":  "#252526",
    "fg_text":    "#cccccc",
    "accent":     "#007acc",
    "accent_hov": "#0098ff",
    "close_hov":  "#e81123",
    "min_hov":    "#3e3e42"
}

class ConfigManager:
    @staticmethod
    def _get_config_path():
        appdata = os.getenv('APPDATA')
        folder = os.path.join(appdata, APP_NAME)
        if not os.path.exists(folder):
            os.makedirs(folder)
        return os.path.join(folder, "config.json")

    @staticmethod
    def load():
        try:
            path = ConfigManager._get_config_path()
            if os.path.exists(path):
                with open(path, 'r') as f:
                    return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            pass
        return {}

    @staticmethod
    def save(data):
        try:
            path = ConfigManager._get_config_path()
            with open(path, 'w') as f:
                json.dump(data, f)
        except IOError:
            pass

class OCREngine:
    def __init__(self):
        self.base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        self.tesseract_cmd = os.path.join(self.base_path, 'Tesseract-OCR', 'tesseract.exe')
        
        if sys.platform.startswith('win'):
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd

    def extract_text(self, img):
        # Preprocessing pipeline
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

        # Virtual screen metrics for multi-monitor support
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
        self.rect = self.canvas.create_rectangle(*self.start_pos, *self.start_pos, outline='#007acc', width=2)

    def _on_drag(self, event):
        cur_x, cur_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_pos[0], self.start_pos[1], cur_x, cur_y)

    def _on_release(self, event):
        x1, y1 = self.start_pos
        x2, y2 = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.destroy()
        
        x_min, x_max = min(x1, x2), max(x1, x2)
        y_min, y_max = min(y1, y2), max(y1, y2)

        # Ignore accidental micro-clicks
        if (x_max - x_min) < 5 or (y_max - y_min) < 5:
            self.on_complete(None)
            return

        bbox = (x_min + self.v_x, y_min + self.v_y, x_max + self.v_x, y_max + self.v_y)
        img = ImageGrab.grab(bbox=bbox, all_screens=True)
        self.on_complete(img)

class ResultPopup(tk.Toplevel):
    def __init__(self, parent, text, timeout=10):
        super().__init__(parent)
        self.overrideredirect(True)
        self.configure(bg=THEME["bg_main"])
        self.attributes('-topmost', True)
        
        w, h = 400, 220
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w // 2) - (w // 2)
        y = (screen_h // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        
        frame = tk.Frame(self, bg=THEME["bg_main"], highlightbackground=THEME["accent"], highlightthickness=1)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="CAPTURED", font=("Segoe UI", 10, "bold"), 
                 bg=THEME["bg_main"], fg=THEME["accent"]).pack(pady=(15, 5))

        preview = text.replace('\n', ' ')[:150] + "..." if len(text) > 150 else text
        lbl = tk.Label(frame, text=preview, font=("Consolas", 9), 
                       bg=THEME["bg_header"], fg=THEME["fg_text"], 
                       wraplength=360, justify="left", padx=10, pady=10)
        lbl.pack(fill=tk.X, padx=20)

        self.lbl_timer = tk.Label(frame, text=f"Auto-closing in {timeout}s", 
                                  font=("Segoe UI", 8), bg=THEME["bg_main"], fg="#666")
        self.lbl_timer.pack(pady=(10, 5))

        tk.Button(frame, text="OK", command=self.destroy, bg=THEME["accent"], fg="white", 
                  activebackground=THEME["accent_hov"], activeforeground="white",
                  bd=0, padx=25, pady=4, cursor="hand2").pack(pady=10)

        self._start_timer(timeout)

    def _start_timer(self, remaining):
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
        self.ocr = OCREngine()
        
        self.config = ConfigManager.load()
        self._setup_window()
        self._build_custom_titlebar()
        self._build_ui()

        # Taskbar visibility fix for borderless window
        self.root.after(10, self._set_app_window)

    def _set_app_window(self):
        hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        style = style & ~WS_EX_TOOLWINDOW
        style = style | WS_EX_APPWINDOW
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
        
        self.root.withdraw()
        self.root.after(10, self.root.deiconify)

    def _setup_window(self):
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

        tk.Label(self.title_bar, text=f"{APP_NAME}", bg=THEME["bg_header"], fg=THEME["fg_text"], 
                 font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=10)

        btn_close = tk.Button(self.title_bar, text="✕", bg=THEME["bg_header"], fg=THEME["fg_text"],
                              bd=0, font=("Arial", 9), width=4, activebackground=THEME["close_hov"], activeforeground="white",
                              command=self._on_close)
        btn_close.pack(side=tk.RIGHT, fill=tk.Y)

        btn_min = tk.Button(self.title_bar, text="—", bg=THEME["bg_header"], fg=THEME["fg_text"],
                            bd=0, font=("Arial", 9, "bold"), width=4, activebackground=THEME["min_hov"], activeforeground="white",
                            command=self._minimize)
        btn_min.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_ui(self):
        main_frame = tk.Frame(self.root, bg=THEME["bg_main"])
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.bind("<Button-1>", self._start_move)
        main_frame.bind("<B1-Motion>", self._do_move)

        btn = tk.Button(main_frame, text="CAPTURE ZONE", command=self._start_snip,
                        font=("Segoe UI", 10, "bold"), bg=THEME["accent"], fg="white",
                        activebackground=THEME["accent_hov"], activeforeground="white",
                        bd=0, cursor="hand2", padx=20, pady=10)
        btn.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    def _start_move(self, event):
        self.x = event.x
        self.y = event.y

    def _do_move(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
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

    def _start_snip(self):
        self.root.withdraw()
        self.root.after(150, lambda: SnippingOverlay(self.root, self._process_snip))

    def _process_snip(self, img):
        self.root.deiconify()
        self.root.overrideredirect(True)
        
        # Re-apply taskbar styling after window restoration
        self.root.after(10, self._set_app_window)

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

    def _on_close(self):
        self.config["x"] = self.root.winfo_x()
        self.config["y"] = self.root.winfo_y()
        ConfigManager.save(self.config)
        self.root.destroy()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = App()
    app.run()