import tkinter as tk
from tkinter import messagebox
import threading
import time
import pytesseract
from PIL import ImageGrab, Image, ImageTk
import pyperclip
import cv2
import numpy as np
import os
import sys
import re
import ctypes
import shutil

# Enable DPI Awareness on Windows (适配 Windows 高 DPI 缩放)
if os.name == 'nt':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# Try to locate tesseract executable (尝试定位 tesseract 可执行文件)
def find_tesseract():
    # 1. Check if tesseract is in PATH (检查 PATH 环境变量)
    if shutil.which('tesseract'):
        return 'tesseract'
        
    # 2. Check common installation paths on Windows (检查 Windows 常见安装路径)
    possible_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        os.path.join(os.getenv('LOCALAPPDATA', ''), r'Tesseract-OCR\tesseract.exe'),
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None

# Initialize Tesseract (初始化 Tesseract)
tesseract_cmd = find_tesseract()
if tesseract_cmd:
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

class ScreenSelector:
    """
    Screen selection window (屏幕选区窗口)
    """
    def __init__(self, master, callback):
        self.master = master
        self.callback = callback
        self.start_x = None
        self.start_y = None
        self.rect = None
        
        # Create a fullscreen overlay (创建一个全屏覆盖层)
        self.top = tk.Toplevel(master)
        self.top.attributes('-fullscreen', True)
        self.top.attributes('-alpha', 0.3)
        self.top.configure(background='grey')
        self.top.attributes('-topmost', True)
        
        self.canvas = tk.Canvas(self.top, cursor="cross", bg='grey')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_move_press)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        self.canvas.bind("<Escape>", lambda e: self.top.destroy())

    def on_button_press(self, event):
        # Record starting position (记录起始位置)
        self.start_x = event.x
        self.start_y = event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

    def on_move_press(self, event):
        # Update rectangle as mouse moves (随鼠标移动更新矩形)
        cur_x, cur_y = (event.x, event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_button_release(self, event):
        # Calculate final coordinates (计算最终坐标)
        if self.start_x is None or self.start_y is None:
            return
            
        end_x, end_y = (event.x, event.y)
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)
        
        # Ensure the rectangle has some size (确保矩形有一定大小)
        if x2 - x1 > 5 and y2 - y1 > 5:
            self.callback((x1, y1, x2, y2))
        
        self.top.destroy()

class MonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("进大哥房间快人一百步 (Screen Monitor)")
        self.root.geometry("500x600")
        self.root.attributes('-topmost', True)  # Keep window on top (窗口置顶)
        
        self.selected_area = None
        self.is_monitoring = False
        self.monitor_thread = None
        self.stop_event = threading.Event()
        self.last_copied_text = ""
        
        # UI Elements (UI 元素)
        self.frame = tk.Frame(root, padx=10, pady=10)
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        self.lbl_status = tk.Label(self.frame, text="状态: 空闲 (Status: Idle)", fg="blue", font=("Arial", 10, "bold"))
        self.lbl_status.pack(pady=5)
        
        self.lbl_area = tk.Label(self.frame, text="未选择区域 (No area selected)", fg="gray")
        self.lbl_area.pack(pady=5)
        
        self.btn_select = tk.Button(self.frame, text="选择屏幕区域 (Select Screen Area)", command=self.select_area, height=2)
        self.btn_select.pack(fill=tk.X, pady=5)
        
        self.btn_start = tk.Button(self.frame, text="开始监控 (Start Monitoring)", command=self.start_monitoring, state=tk.DISABLED, bg="#ddffdd", height=2)
        self.btn_start.pack(fill=tk.X, pady=5)
        
        self.btn_stop = tk.Button(self.frame, text="停止监控 (Stop Monitoring)", command=self.stop_monitoring, state=tk.DISABLED, bg="#ffdddd", height=2)
        self.btn_stop.pack(fill=tk.X, pady=5)
        
        self.lbl_image_preview = tk.Label(self.frame, text="[监控区域预览 / Preview]", bg="lightgray", height=5)
        self.lbl_image_preview.pack(fill=tk.X, pady=5, padx=5)

        self.lbl_preview = tk.Label(self.frame, text="最新识别文本 (Latest Text):")
        self.lbl_preview.pack(pady=(10, 0))
        
        self.txt_output = tk.Text(self.frame, height=5, width=40)
        self.txt_output.pack(pady=5)
        
        # Check for Tesseract (检查 Tesseract)
        if not tesseract_cmd:
            messagebox.showwarning("未找到 Tesseract (Tesseract Not Found)", 
                                   "在您的系统中未找到 Tesseract-OCR。\n"
                                   "Tesseract-OCR is not found on your system.\n\n"
                                   "请安装它并确保它在默认路径中或添加到 PATH 环境变量。\n"
                                   "Please install it and ensure it's in default path or added to PATH.\n")
            self.lbl_status.config(text="错误: 未找到 Tesseract (Error: Tesseract not found)", fg="red")

    def select_area(self):
        # Minimize window to see screen better (最小化窗口以便观察屏幕)
        self.root.iconify()
        # Delay slightly to allow minimize animation (稍作延迟等待最小化动画)
        self.root.after(200, lambda: ScreenSelector(self.root, self.on_area_selected))

    def on_area_selected(self, area):
        self.selected_area = area
        self.lbl_area.config(text=f"已选择 (Selected): {area}")
        self.btn_start.config(state=tk.NORMAL)
        self.root.deiconify()  # Restore window (恢复窗口)

    def start_monitoring(self):
        if not self.selected_area:
            return
            
        self.is_monitoring = True
        self.stop_event.clear()
        self.btn_start.config(state=tk.DISABLED)
        self.btn_select.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.lbl_status.config(text="状态: 监控中... (Monitoring...)", fg="green")
        
        self.monitor_thread = threading.Thread(target=self.monitor_loop)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()

    def stop_monitoring(self):
        self.is_monitoring = False
        self.stop_event.set()
        self.btn_start.config(state=tk.NORMAL)
        self.btn_select.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.lbl_status.config(text="状态: 已停止 (Stopped)", fg="red")

    def extract_digits(self, text):
        # Extract only digits from the text (仅提取数字)
        return "".join(re.findall(r'\d+', text))

    def monitor_loop(self):
        while not self.stop_event.is_set():
            try:
                # Capture screen (屏幕截图)
                img = ImageGrab.grab(bbox=self.selected_area)
                
                # Preprocessing for better OCR (图像预处理)
                # Convert to numpy array
                img_np = np.array(img)
                # Convert to grayscale (转灰度)
                gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
                # Thresholding to get black text on white background (二值化)
                # OTSU thresholding is usually good
                _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                
                # Convert back to PIL Image
                processed_img = Image.fromarray(thresh)
                
                # Run OCR (执行 OCR)
                # config='--psm 6' assumes a single uniform block of text
                text = pytesseract.image_to_string(processed_img, config='--psm 6').strip()
                
                # Update UI (must be done in main thread) (更新 UI，必须在主线程)
                self.root.after(0, self.update_preview, text, img)
                
                # Extract digits (提取数字)
                digits = self.extract_digits(text)
                
                if digits:
                    if digits != self.last_copied_text:
                        pyperclip.copy(digits)
                        self.last_copied_text = digits
                        self.root.after(0, lambda t=digits: self.lbl_status.config(text=f"已复制 (Copied): {t}", fg="green"))
                
                # Wait a bit (等待)
                time.sleep(1.0)
                
            except Exception as e:
                print(f"Error: {e}")
                self.root.after(0, lambda err=str(e): self.lbl_status.config(text=f"错误 (Error): {err}", fg="red"))
                time.sleep(2.0)

    def update_preview(self, text, image=None):
        self.txt_output.delete(1.0, tk.END)
        self.txt_output.insert(tk.END, text)
        
        if image:
            # Resize image to fit label height (approx 80px) while maintaining aspect ratio
            # 调整图片大小以适应标签高度
            target_h = 80
            aspect = image.width / image.height
            target_w = int(target_h * aspect)
            
            # Limit width if too wide (限制宽度)
            if target_w > 300:
                target_w = 300
                target_h = int(target_w / aspect)
                
            resized_img = image.resize((target_w, target_h), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(resized_img)
            self.lbl_image_preview.config(image=photo, text="", height=0, width=0) # Reset text/height/width to let image dictate size
            self.lbl_image_preview.image = photo  # Keep reference

if __name__ == "__main__":
    root = tk.Tk()
    app = MonitorApp(root)
    root.mainloop()
