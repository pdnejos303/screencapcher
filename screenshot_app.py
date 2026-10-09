import datetime
import json
import os
import sys
import threading
import queue
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import tkinter as tk
import customtkinter as ctk
try:
    import pywinstyles
except ImportError:
    pywinstyles = None
import shutil
import keyboard
import win32api
import win32com.client
from mss import mss
from mss.tools import to_png
from PIL import Image, ImageTk, ImageDraw

class DummyStream:
    def write(self, *args, **kwargs): pass
    def flush(self, *args, **kwargs): pass
    def reconfigure(self, *args, **kwargs): pass

if sys.stdout is None:
    sys.stdout = DummyStream()
    sys.stderr = DummyStream()
else:
    try:
        sys.stdout.reconfigure(line_buffering=True)
    except:
        pass

# ตั้งค่า Theme ของ UI ใหม่ให้สวยงาม
ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

# ════════════════════════════════════════════════════════════════
# 1. ค้นหาหน้าจอ
# ════════════════════════════════════════════════════════════════
sct = mss()
mss_monitors = sct.monitors

win_monitors_raw = win32api.EnumDisplayMonitors(None, None)
win_monitor_list = []

for hMonitor, hdcMonitor, pyRect in win_monitors_raw:
    info = win32api.GetMonitorInfo(hMonitor)
    device_name = info["Device"]
    win_num = int(device_name.replace("\\\\.\\DISPLAY", ""))
    rect = info["Monitor"]
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]

    mss_idx = None
    for idx in range(1, len(mss_monitors)):
        m = mss_monitors[idx]
        if m["left"] == rect[0] and m["top"] == rect[1]:
            mss_idx = idx
            break

    win_monitor_list.append({
        "win_num": win_num,
        "device": device_name,
        "left": rect[0],
        "top": rect[1],
        "width": width,
        "height": height,
        "mss_idx": mss_idx,
        "is_primary": bool(info["Flags"] & 1),
    })

win_monitor_list.sort(key=lambda x: x["win_num"])

# Global Variables
config_file = os.path.join(os.path.abspath(os.path.dirname(__file__)), "config.json")
settings = {
    "monitor": None,
    "save_folder": os.path.abspath(os.path.dirname(__file__)),
    "save_filename": f"TestReport_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
    "use_existing_word": False,
    "existing_word_path": "",
    "export_format": "Word",
    "backup_images": True,
    "hk_capture": "shift+s",
    "hk_capture_region": "shift+d",
    "hk_annotate": "shift+a",
    "hk_major_next": "shift+right",
    "hk_major_prev": "shift+left",
    "hk_minor_next": "ctrl+right",
    "hk_minor_prev": "ctrl+left"
}
saved_monitor_num = None

if os.path.exists(config_file):
    try:
        with open(config_file, "r", encoding="utf-8") as f:
            loaded = json.load(f)
            for k in ["save_folder", "hk_capture", "hk_capture_region", "hk_annotate", "hk_major_next", "hk_major_prev", "hk_minor_next", "hk_minor_prev", "use_existing_word", "existing_word_path", "export_format", "backup_images"]:
                if k in loaded: settings[k] = loaded[k]
            if "save_filename" in loaded and loaded["save_filename"]:
                settings["save_filename"] = loaded["save_filename"]
            if "monitor_num" in loaded:
                saved_monitor_num = loaded["monitor_num"]
    except: pass

setup_done = False


# ════════════════════════════════════════════════════════════════
# 2. Setup UI (CustomTkinter)
# ════════════════════════════════════════════════════════════════
def run_setup_ui():
    global setup_done
    
    app = ctk.CTk()
    app.title("Screenshot Overlay - Setup")
    app.geometry("1000x550")
    app.minsize(900, 500)
    
    # --- Modern Colors & Fonts ---
    COLOR_BG = "#0B0F19"
    COLOR_CARD = "#1A2235"
    COLOR_ACCENT = "#3B82F6"
    COLOR_ACCENT_HOVER = "#2563EB"
    COLOR_TEXT = "#FFFFFF"
    COLOR_TEXT_MUTED = "#94A3B8"
    FONT_FAMILY = "Segoe UI"
    
    app.configure(fg_color=COLOR_BG)

    app.update_idletasks()
    x = (app.winfo_screenwidth() // 2) - (1000 // 2)
    y = (app.winfo_screenheight() // 2) - (700 // 2)
    app.geometry(f"+{x}+{y}")

    # Header
    header_frame = ctk.CTkFrame(app, fg_color="transparent")
    header_frame.pack(fill="x", padx=40, pady=(30, 20))
    
    title_lbl = ctk.CTkLabel(header_frame, text="Screenshot Automation", font=(FONT_FAMILY, 32, "bold"), text_color=COLOR_TEXT)
    title_lbl.pack(side="left")
    
    subtitle_lbl = ctk.CTkLabel(header_frame, text="ตั้งค่าหน้าจอและปุ่มลัดก่อนเริ่มการทำงาน", font=(FONT_FAMILY, 14), text_color=COLOR_TEXT_MUTED)
    subtitle_lbl.pack(side="left", padx=15, pady=6)

    # Footer (fixed at bottom so it never gets cut off)
    bottom_frame = ctk.CTkFrame(app, fg_color="transparent")
    bottom_frame.pack(side="bottom", fill="x", padx=40, pady=(10, 25))
    
    # Main Grid Container (now a Master Scrollable Area)
    main_container = ctk.CTkScrollableFrame(app, fg_color="transparent")
    main_container.pack(side="top", fill="both", expand=True, padx=20, pady=(0, 10))
    
    inner_frame = ctk.CTkFrame(main_container, fg_color="transparent")
    inner_frame.pack(fill="both", expand=True)

    left_col = ctk.CTkFrame(inner_frame, fg_color="transparent")
    left_col.pack(side="left", fill="both", expand=True, padx=(10, 15))

    right_col = ctk.CTkFrame(inner_frame, fg_color="transparent")
    right_col.pack(side="right", fill="both", expand=True, padx=(15, 10))

    # --- Section 1: หน้าจอ ---
    frame_mon = ctk.CTkFrame(left_col, corner_radius=16, fg_color=COLOR_CARD, border_width=1, border_color="#2E3C56")
    frame_mon.pack(fill="both", expand=True)
    
    mon_header = ctk.CTkFrame(frame_mon, fg_color="transparent")
    mon_header.pack(fill="x", padx=25, pady=(25, 10))
    ctk.CTkLabel(mon_header, text="🖥️ เลือกหน้าจอ", font=(FONT_FAMILY, 18, "bold"), text_color=COLOR_TEXT).pack(anchor="w")
    ctk.CTkLabel(mon_header, text="ระบุหน้าจอที่ต้องการแคปเจอร์", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT_MUTED).pack(anchor="w")
    
    monitor_scroll = ctk.CTkFrame(frame_mon, fg_color="transparent")
    monitor_scroll.pack(fill="both", expand=True, padx=15, pady=10)

    selected_monitor_num = ctk.IntVar(value=win_monitor_list[0]["win_num"])
    monitor_buttons = {}

    def on_monitor_select(num):
        selected_monitor_num.set(num)
        for n, btn in monitor_buttons.items():
            if n == num:
                btn.configure(fg_color=COLOR_ACCENT, border_color=COLOR_ACCENT, text_color="white")
            else:
                btn.configure(fg_color="#222D44", border_color="#2E3C56", text_color=COLOR_TEXT_MUTED)

    for m in win_monitor_list:
        tag = "⭐ จอหลัก (Primary)" if m["is_primary"] else f"จอที่ {m['win_num']}"
        res_text = f"{m['width']} x {m['height']}"
        pos_text = f"X: {m['left']}, Y: {m['top']}"
        
        btn_container = ctk.CTkFrame(monitor_scroll, fg_color="transparent")
        btn_container.pack(fill="x", pady=8, padx=10)
        
        btn = ctk.CTkButton(
            btn_container, text=f"{tag}\n{res_text}   |   {pos_text}", 
            corner_radius=12, border_width=2,
            font=(FONT_FAMILY, 14, "bold"), height=75,
            command=lambda num=m['win_num']: on_monitor_select(num),
            hover_color=COLOR_ACCENT_HOVER
        )
        btn.pack(fill="x", expand=True)
        monitor_buttons[m['win_num']] = btn
        
    default_mon_num = saved_monitor_num if saved_monitor_num else next((m for m in win_monitor_list if m["is_primary"]), win_monitor_list[0])["win_num"]
    if default_mon_num not in [m["win_num"] for m in win_monitor_list]:
        default_mon_num = win_monitor_list[0]["win_num"]
    on_monitor_select(default_mon_num)

    # --- Section 2: ปลายทางไฟล์ (New/Existing) ---
    frame_file = ctk.CTkFrame(left_col, corner_radius=16, fg_color=COLOR_CARD, border_width=1, border_color="#2E3C56")
    frame_file.pack(fill="x", pady=(15, 0))
    
    file_header = ctk.CTkFrame(frame_file, fg_color="transparent")
    file_header.pack(fill="x", padx=25, pady=(20, 5))
    ctk.CTkLabel(file_header, text="📁 ปลายทางเอกสาร (Word)", font=(FONT_FAMILY, 18, "bold"), text_color=COLOR_TEXT).pack(anchor="w")
    ctk.CTkLabel(file_header, text="เลือกสร้างไฟล์ใหม่ หรือใช้ Template เดิม", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT_MUTED).pack(anchor="w")
    
    

    

    # New File Frame
    new_file_frame = ctk.CTkFrame(frame_file, fg_color="transparent")
    new_file_frame.pack(fill="x", padx=20, pady=5)
    folder_frame = ctk.CTkFrame(new_file_frame, fg_color="transparent")
    folder_frame.pack(fill="x", pady=5)
    entry_folder = ctk.CTkEntry(folder_frame, height=40, font=(FONT_FAMILY, 13), corner_radius=8, border_color="#2E3C56", fg_color="#0B0F19")
    entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 10))
    entry_folder.insert(0, settings["save_folder"])
    def browse_folder():
        folder = filedialog.askdirectory(initialdir=settings["save_folder"])
        if folder:
            entry_folder.delete(0, "end")
            entry_folder.insert(0, os.path.abspath(folder))
    ctk.CTkButton(folder_frame, text="Browse", width=100, height=40, corner_radius=8, fg_color="#2E3C56", hover_color="#3B4B68", font=(FONT_FAMILY, 13, "bold"), command=browse_folder).pack(side="left")
    
    file_frame = ctk.CTkFrame(new_file_frame, fg_color="transparent")
    file_frame.pack(fill="x", pady=(5, 10))
    ctk.CTkLabel(file_frame, text="ชื่อไฟล์:", text_color=COLOR_TEXT_MUTED, font=(FONT_FAMILY, 13)).pack(side="left", padx=(0, 10))
    entry_filename = ctk.CTkEntry(file_frame, height=40, font=(FONT_FAMILY, 13), corner_radius=8, border_color="#2E3C56", fg_color="#0B0F19")
    entry_filename.pack(side="left", fill="x", expand=True)
    entry_filename.insert(0, settings["save_filename"])

    global export_format_var
    export_format_var = ctk.StringVar(value=settings.get("export_format", "Word"))
    format_frame = ctk.CTkFrame(new_file_frame, fg_color="transparent")
    format_frame.pack(fill="x", pady=(5, 10))
    ctk.CTkLabel(format_frame, text="รูปแบบเอกสาร:", text_color=COLOR_TEXT_MUTED, font=(FONT_FAMILY, 13)).pack(side="left", padx=(0, 10))
    seg_btn = ctk.CTkSegmentedButton(format_frame, values=["Word", "PDF"], variable=export_format_var, font=(FONT_FAMILY, 13, "bold"), selected_color=COLOR_ACCENT)
    seg_btn.pack(side="left")

    # Existing File Frame
    
    
    # Backup Checkbox
    backup_var = ctk.BooleanVar(value=settings.get("backup_images", True))
    chk = ctk.CTkCheckBox(frame_file, text="✅ สำรองไฟล์รูปภาพ (.png) ต้นฉบับใน Backup_Images", variable=backup_var, font=(FONT_FAMILY, 13), fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER)
    chk.pack(anchor="w", padx=25, pady=(5, 20))

    # --- Section 3: Hotkeys ---
    frame_hk = ctk.CTkFrame(right_col, corner_radius=16, fg_color=COLOR_CARD, border_width=1, border_color="#2E3C56")
    frame_hk.pack(fill="both", expand=True)
    
    hk_header = ctk.CTkFrame(frame_hk, fg_color="transparent")
    hk_header.pack(fill="x", padx=25, pady=(25, 5))
    ctk.CTkLabel(hk_header, text="⌨️ ตั้งค่าปุ่มลัด (Hotkeys)", font=(FONT_FAMILY, 18, "bold"), text_color=COLOR_TEXT).pack(anchor="w")
    ctk.CTkLabel(hk_header, text="คลิกที่ปุ่มแล้วกดปุ่มบนคีย์บอร์ดที่ต้องการ", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT_MUTED).pack(anchor="w")
    
    hk_scroll = ctk.CTkFrame(frame_hk, fg_color="transparent")
    hk_scroll.pack(fill="both", expand=True, padx=15, pady=5)

    def create_hk_row(parent, label_text, key_name):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(row, text=label_text, width=170, anchor="w", font=(FONT_FAMILY, 14)).pack(side="left")
        
        btn_hk = ctk.CTkButton(row, text=str(settings[key_name]).upper(), width=160, height=35, corner_radius=8, 
                               fg_color="#222D44", border_width=1, border_color="#2E3C56", hover_color="#3B4B68", font=(FONT_FAMILY, 13, "bold"))
        btn_hk.pack(side="right")
        
        def listen_hotkey():
            btn_hk.configure(text="Listening...", fg_color="#EAB308", text_color="#000000", border_color="#EAB308")
            app.update()
            
            def hotkey_thread():
                try:
                    hk = keyboard.read_hotkey(suppress=False)
                    app.after(0, lambda: finish_listen(hk))
                except Exception as e:
                    app.after(0, lambda: finish_listen(settings[key_name]))

            def finish_listen(hk):
                settings[key_name] = hk
                btn_hk.configure(text=str(hk).upper(), fg_color="#222D44", text_color=COLOR_TEXT, border_color="#2E3C56")

            threading.Thread(target=hotkey_thread, daemon=True).start()

        btn_hk.configure(command=listen_hotkey)
        return btn_hk

    btn_hk_cap = create_hk_row(hk_scroll, "📸 แคปหน้าจอ (เต็มจอ)", "hk_capture")
    btn_hk_cap_region = create_hk_row(hk_scroll, "✂️ แคป (เลือกพื้นที่)", "hk_capture_region")
    btn_hk_annotate = create_hk_row(hk_scroll, "✏️ แคป (วาด/วง)", "hk_annotate")
    btn_hk_maj_next = create_hk_row(hk_scroll, "⬆️ ข้อหลักถัดไป (+1.0)", "hk_major_next")
    btn_hk_maj_prev = create_hk_row(hk_scroll, "⬇️ ข้อหลักก่อนหน้า (-1.0)", "hk_major_prev")
    btn_hk_min_next = create_hk_row(hk_scroll, "➡️ ข้อย่อยถัดไป (+0.1)", "hk_minor_next")
    btn_hk_min_prev = create_hk_row(hk_scroll, "⬅️ ข้อย่อยก่อนหน้า (-0.1)", "hk_minor_prev")

    # --- Start Button ---
    def on_start():
        global setup_done
        mon_num = selected_monitor_num.get()
        settings["monitor"] = next(m for m in win_monitor_list if m["win_num"] == mon_num)
        
        settings["use_existing_word"] = False
        settings["existing_word_path"] = ""
        settings["backup_images"] = backup_var.get()
        settings["save_folder"] = entry_folder.get()
        
        fmt = export_format_var.get()
        settings["export_format"] = fmt
        
        filename = entry_filename.get()
        if filename.endswith(".docx") or filename.endswith(".pdf"):
            filename = filename[:-5] if filename.endswith(".docx") else filename[:-4]
            
        if fmt == "Word":
            filename += ".docx"
        else:
            filename += ".pdf"
            settings["use_existing_word"] = False
            
        settings["save_filename"] = filename
        
        setup_done = True
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                save_data = {
                    "monitor_num": mon_num,
                    "save_folder": settings["save_folder"],
                    "save_filename": settings["save_filename"],
                    "use_existing_word": settings["use_existing_word"],
                    "existing_word_path": settings["existing_word_path"],
                    "export_format": settings["export_format"],
                    "backup_images": settings["backup_images"],
                    "hk_capture": settings["hk_capture"],
                    "hk_capture_region": settings["hk_capture_region"],
                    "hk_annotate": settings.get("hk_annotate", "shift+a"),
                    "hk_major_next": settings["hk_major_next"],
                    "hk_major_prev": settings["hk_major_prev"],
                    "hk_minor_next": settings["hk_minor_next"],
                    "hk_minor_prev": settings["hk_minor_prev"]
                }
                json.dump(save_data, f, ensure_ascii=False, indent=4)
        except: pass
        app.destroy()

    btn_start = ctk.CTkButton(bottom_frame, text="🚀 เปิด Overlay เริ่มทำงาน", font=(FONT_FAMILY, 22, "bold"), height=65, corner_radius=15, fg_color="#10B981", hover_color="#059669", text_color="#FFFFFF", command=on_start)
    btn_start.pack(fill="x", expand=True)

    app.mainloop()

run_setup_ui()

if not setup_done:
    sys.exit(0)

monitor_to_capture = mss_monitors[settings["monitor"]["mss_idx"]]
full_save_path = os.path.join(settings["save_folder"], settings["save_filename"])

# ════════════════════════════════════════════════════════════════
# 3. เปิด Word
# ════════════════════════════════════════════════════════════════
capture_history_pdf = []
if settings.get("export_format", "Word") == "Word":
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        
        if settings.get("use_existing_word") and os.path.exists(settings.get("existing_word_path", "")):
            doc = word.Documents.Open(settings["existing_word_path"])
            sel = word.Selection
            sel.EndKey(Unit=6) # wdStory
            sel.TypeParagraph()
            full_save_path = settings["existing_word_path"]
        else:
            doc = word.Documents.Add()
            sel = word.Selection
            sel.Font.Size = 20
            sel.Font.Bold = True
            sel.TypeText("Automation Screenshot Report\n")
            sel.Font.Size = 11
            sel.Font.Bold = False
            sel.TypeParagraph()
            full_save_path = os.path.join(settings["save_folder"], settings["save_filename"])
            doc.SaveAs2(full_save_path)
    except Exception as e:
        messagebox.showerror("Error", f"ไม่สามารถเปิด Microsoft Word ได้:\n{e}")
        sys.exit(1)
else:
    word = None
    doc = None
    full_save_path = os.path.join(settings["save_folder"], settings["save_filename"])


# ════════════════════════════════════════════════════════════════
# 4. ตัวแปรสถานะ
# ════════════════════════════════════════════════════════════════
current_dir = os.path.abspath(os.path.dirname(__file__))
capture_count = 0

step_queue = queue.Queue()
capture_event = threading.Event()
capture_region_event = threading.Event()
capture_annotate_event = threading.Event()
snip_active = False
annotate_active = False
custom_annotate_path = None
snip_region = None
stop_event = threading.Event()
state_lock = threading.Lock()

# ════════════════════════════════════════════════════════════════
# 5. Hotkey callbacks
# ════════════════════════════════════════════════════════════════
def on_capture(): capture_event.set()
def on_stop():
    stop_event.set()
    capture_event.set()

def next_minor(): step_queue.put(('minor', 1))
def prev_minor(): step_queue.put(('minor', -1))
def next_major(): step_queue.put(('major', 1))
def prev_major(): step_queue.put(('major', -1))

try:
    keyboard.add_hotkey(settings["hk_capture"], on_capture)
    keyboard.add_hotkey(settings.get("hk_capture_region", "shift+d"), lambda: capture_region_event.set())
    keyboard.add_hotkey(settings.get("hk_annotate", "shift+a"), lambda: capture_annotate_event.set())
    keyboard.add_hotkey("esc", on_stop)
    keyboard.add_hotkey(settings["hk_minor_next"], next_minor)
    keyboard.add_hotkey(settings["hk_minor_prev"], prev_minor)
    keyboard.add_hotkey(settings["hk_major_next"], next_major)
    keyboard.add_hotkey(settings["hk_major_prev"], prev_major)
except Exception as e:
    messagebox.showerror("Hotkey Error", f"ตั้งค่าปุ่มลัดผิดพลาด:\n{e}")
    sys.exit(1)


# ════════════════════════════════════════════════════════════════
# 6. Floating Overlay Window (CustomTkinter)
# ════════════════════════════════════════════════════════════════
overlay = None
overlay_entry_step = None
entry_detail = None
label_count = None

def create_overlay():
    global overlay, overlay_entry_step, entry_detail, label_count

    COLOR_BG = "#0B0F19"
    COLOR_CARD = "#1A2235"
    COLOR_ACCENT = "#3B82F6"
    COLOR_ACCENT_HOVER = "#2563EB"
    COLOR_TEXT = "#FFFFFF"
    COLOR_TEXT_MUTED = "#94A3B8"
    FONT_FAMILY = "Segoe UI"

    overlay = ctk.CTk()
    overlay.title("Screenshot Overlay")
    overlay.geometry("340x350")
    
    # หน้าต่างลอย ไร้ขอบ
    overlay.attributes("-topmost", True)
    overlay.overrideredirect(True)
    overlay.attributes("-alpha", 1.0)
    # Remove pywinstyles to ensure no weird transparent artifacts
    overlay.configure(fg_color=COLOR_BG)
    
    # Outer Frame
    main_frame = ctk.CTkFrame(overlay, corner_radius=16, border_width=1, border_color="#2E3C56", fg_color=COLOR_CARD)
    main_frame.pack(fill="both", expand=True, padx=2, pady=2)
    
    # ═══ Drag Bar ═══
    drag_bar = ctk.CTkFrame(main_frame, height=36, corner_radius=14, fg_color="transparent")
    drag_bar.pack(fill="x", padx=8, pady=8)
    
    # Gripper icon / indicator
    gripper = ctk.CTkLabel(drag_bar, text="•••", font=(FONT_FAMILY, 14, "bold"), text_color=COLOR_TEXT_MUTED)
    gripper.pack(side="left", padx=15)
    
    lbl_title = ctk.CTkLabel(drag_bar, text="Toolbox", font=(FONT_FAMILY, 13, "bold"), text_color=COLOR_TEXT)
    lbl_title.pack(side="left")
    
    def on_drag_start(event):
        overlay._drag_start_x = event.x_root
        overlay._drag_start_y = event.y_root
        overlay._window_start_x = overlay.winfo_x()
        overlay._window_start_y = overlay.winfo_y()
        
    def on_drag_motion(event):
        dx = event.x_root - overlay._drag_start_x
        dy = event.y_root - overlay._drag_start_y
        x = overlay._window_start_x + dx
        y = overlay._window_start_y + dy
        overlay.geometry(f"+{x}+{y}")
        
    drag_bar.bind("<Button-1>", on_drag_start)
    drag_bar.bind("<B1-Motion>", on_drag_motion)
    gripper.bind("<Button-1>", on_drag_start)
    gripper.bind("<B1-Motion>", on_drag_motion)
    lbl_title.bind("<Button-1>", on_drag_start)
    lbl_title.bind("<B1-Motion>", on_drag_motion)
    
    def ui_close():
        import os
        os._exit(0)
        
    btn_close = ctk.CTkButton(drag_bar, text="✕", width=30, height=30, corner_radius=8, fg_color="transparent", hover_color="#EF4444", text_color=COLOR_TEXT_MUTED, font=(FONT_FAMILY, 14, "bold"), command=ui_close)
    btn_close.pack(side="right", padx=5)
    
    is_minimized = False
    def toggle_minimize():
        nonlocal is_minimized
        if not is_minimized:
            step_frame.pack_forget()
            entry_detail.pack_forget()
            action_frame.pack_forget()
            status_frame.pack_forget()
            overlay.geometry("340x55")
            btn_minimize.configure(text="⬜")
            is_minimized = True
        else:
            step_frame.pack(fill="x", padx=20, pady=(5, 15))
            entry_detail.pack(fill="x", padx=20, pady=(0, 15))
            action_frame.pack(fill="x", padx=20)
            status_frame.pack(fill="x", padx=20, pady=(15, 10))
            overlay.geometry("340x350")
            btn_minimize.configure(text="—")
            is_minimized = False
            
    btn_minimize = ctk.CTkButton(drag_bar, text="—", width=30, height=30, corner_radius=8, fg_color="transparent", hover_color="#3B4B68", text_color=COLOR_TEXT_MUTED, font=(FONT_FAMILY, 14, "bold"), command=toggle_minimize)
    btn_minimize.pack(side="right", padx=5)
    
    # ═══ Step Display & Controls ═══
    step_frame = ctk.CTkFrame(main_frame, fg_color="#0B0F19", corner_radius=12)
    step_frame.pack(fill="x", padx=20, pady=(5, 15))
    
    # Control Row
    ctrl_row = ctk.CTkFrame(step_frame, fg_color="transparent")
    ctrl_row.pack(fill="x", padx=15, pady=15)
    
    # Major Controls (Left)
    maj_frame = ctk.CTkFrame(ctrl_row, fg_color="transparent")
    maj_frame.pack(side="left")
    ctk.CTkLabel(maj_frame, text="Major", font=(FONT_FAMILY, 11), text_color=COLOR_TEXT_MUTED).pack(pady=(0, 2))
    maj_btn_frame = ctk.CTkFrame(maj_frame, fg_color="transparent")
    maj_btn_frame.pack()
    ctk.CTkButton(maj_btn_frame, text="−", width=30, height=30, corner_radius=6, command=prev_major, fg_color="#222D44", hover_color="#2E3C56", text_color=COLOR_TEXT).pack(side="left", padx=2)
    ctk.CTkButton(maj_btn_frame, text="+", width=30, height=30, corner_radius=6, command=next_major, fg_color="#222D44", hover_color="#2E3C56", text_color=COLOR_TEXT).pack(side="left", padx=2)
    
    # Step Number (Center)
    overlay_entry_step = ctk.CTkEntry(
        ctrl_row, 
        font=(FONT_FAMILY, 42, "bold"), 
        text_color=COLOR_ACCENT,
        fg_color="transparent", 
        border_width=0, 
        justify="center"
    )
    overlay_entry_step.insert(0, "1.1")
    overlay_entry_step.pack(side="left", expand=True, fill="x")
    
    # Minor Controls (Right)
    min_frame = ctk.CTkFrame(ctrl_row, fg_color="transparent")
    min_frame.pack(side="right")
    ctk.CTkLabel(min_frame, text="Minor", font=(FONT_FAMILY, 11), text_color=COLOR_TEXT_MUTED).pack(pady=(0, 2))
    min_btn_frame = ctk.CTkFrame(min_frame, fg_color="transparent")
    min_btn_frame.pack()
    ctk.CTkButton(min_btn_frame, text="−", width=30, height=30, corner_radius=6, command=prev_minor, fg_color="#222D44", hover_color="#2E3C56", text_color=COLOR_TEXT).pack(side="left", padx=2)
    ctk.CTkButton(min_btn_frame, text="+", width=30, height=30, corner_radius=6, command=next_minor, fg_color="#222D44", hover_color="#2E3C56", text_color=COLOR_TEXT).pack(side="left", padx=2)
    
    # ═══ Detail Input ═══
    entry_detail = ctk.CTkEntry(main_frame, placeholder_text="พิมพ์รายละเอียดเพิ่มเติม (Optional)...", height=40, font=(FONT_FAMILY, 13), corner_radius=10, fg_color="#0B0F19", border_color="#2E3C56", text_color=COLOR_TEXT)
    entry_detail.pack(fill="x", padx=20, pady=(0, 15))
    
    # ═══ Action Buttons ═══
    action_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    action_frame.pack(fill="x", padx=20)
    
    def on_btn_capture(): on_capture()
    def on_btn_snip(): capture_region_event.set()
    
    btn_cap = ctk.CTkButton(action_frame, text="📸 แคปเต็มจอ", height=45, corner_radius=12, font=(FONT_FAMILY, 15, "bold"), fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, command=on_btn_capture)
    btn_cap.pack(side="left", fill="x", expand=True, padx=(0, 5))
    
    btn_snip = ctk.CTkButton(action_frame, text="✂️ เลือกพื้นที่", height=45, corner_radius=12, font=(FONT_FAMILY, 15, "bold"), fg_color="#0078D7", hover_color="#005A9E", command=on_btn_snip)
    btn_snip.pack(side="right", fill="x", expand=True, padx=(5, 0))

    # ═══ Status Info ═══
    status_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    status_frame.pack(fill="x", padx=20, pady=(15, 10))
    
    info_text = f"Shortcut: {str(settings['hk_capture']).upper()}"
    ctk.CTkLabel(status_frame, text=info_text, font=(FONT_FAMILY, 11), text_color=COLOR_TEXT_MUTED).pack(side="left")
    
    label_count = ctk.CTkLabel(status_frame, text="แคปแล้ว: 0", font=(FONT_FAMILY, 12, "bold"), text_color=COLOR_ACCENT)
    label_count.pack(side="right")

    # Position at bottom-right corner
    overlay.update_idletasks()
    screen_w = overlay.winfo_screenwidth()
    screen_h = overlay.winfo_screenheight()
    overlay.geometry(f"+{screen_w - 380}+{screen_h - 450}")

    overlay.after(100, check_events)
    overlay.mainloop()

def update_count(count):
    if label_count:
        label_count.configure(text=f"แคปแล้ว: {count}")

# ════════════════════════════════════════════════════════════════
# 7. Capture Logic (Main Thread)
# ════════════════════════════════════════════════════════════════

def build_pdf():
    try:
        from fpdf import FPDF
        pdf = FPDF()
        font_path = r"C:\Windows\Fonts\tahoma.ttf"
        has_thai_font = os.path.exists(font_path)
        if has_thai_font:
            pdf.add_font("Tahoma", "", font_path)
            
        for cap in capture_history_pdf:
            pdf.add_page()
            if has_thai_font: pdf.set_font("Tahoma", "", 14)
            else: pdf.set_font("Arial", "", 14)
            
            pdf.set_text_color(153, 51, 0)
            pdf.cell(200, 10, text=str(cap['step']), new_x="LMARGIN", new_y="NEXT")
            
            if cap['detail']:
                pdf.set_text_color(102, 102, 102)
                if has_thai_font: pdf.set_font("Tahoma", "", 11)
                else: pdf.set_font("Arial", "", 11)
                pdf.cell(200, 8, text=f"  {cap['detail']}", new_x="LMARGIN", new_y="NEXT")
                
            pdf.set_text_color(153, 153, 153)
            if has_thai_font: pdf.set_font("Tahoma", "", 9)
            else: pdf.set_font("Arial", "", 9)
            pdf.cell(200, 6, text=f"  Captured at: {cap['timestamp']}", new_x="LMARGIN", new_y="NEXT")
            
            pdf.ln(5)
            
            with Image.open(cap["image"]) as img:
                w, h = img.size
            max_w = 190
            max_h = 240
            ratio = min(max_w / w, max_h / h)
            pdf.image(cap['image'], w=w*ratio, h=h*ratio)
            
        pdf.output(full_save_path)
    except Exception as e:
        print("PDF Error:", e)

def process_capture(region=None, custom_image_path=None):
    global capture_count

    capture_count += 1
    step = overlay_entry_step.get().strip() if overlay_entry_step else "1.1"
    detail = entry_detail.get().strip() if entry_detail else ""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    filename = os.path.join(current_dir, f"_temp_cap_{capture_count}.png")
    if custom_image_path:
        import shutil
        shutil.copy(custom_image_path, filename)
    else:
        target_region = region if region else monitor_to_capture
        sct_img = sct.grab(target_region)
        to_png(sct_img.rgb, sct_img.size, output=filename)
    
    if settings.get("backup_images", False):
        try:
            backup_dir = os.path.join(settings["save_folder"], "Backup_Images")
            os.makedirs(backup_dir, exist_ok=True)
            b_filename = os.path.join(backup_dir, f"Step_{step}_{capture_count}.png")
            shutil.copy(filename, b_filename)
        except: pass

    is_pdf = settings.get("export_format", "Word") == "PDF"
    
    # Visual Feedback
    def show_flash():
        if overlay:
            flash = ctk.CTkToplevel(overlay)
            flash.overrideredirect(True)
            flash.attributes("-topmost", True)
            flash.attributes("-transparentcolor", "black")
            mon = settings["monitor"]
            flash.geometry(f"{mon['width']}x{mon['height']}+{mon['left']}+{mon['top']}")
            flash.configure(fg_color="#4ADE80")
            flash.attributes("-alpha", 0.3)
            lbl = ctk.CTkLabel(flash, text="📸 แคปเจอร์สำเร็จ!", font=("Segoe UI", 48, "bold"), text_color="white", fg_color="transparent")
            lbl.place(relx=0.5, rely=0.5, anchor="center")
            def fade_out(alpha):
                if alpha > 0:
                    flash.attributes("-alpha", alpha)
                    flash.after(30, fade_out, alpha - 0.05)
                else:
                    flash.destroy()
            flash.after(50, fade_out, 0.3)

    if is_pdf:
        # Save for PDF
        img_path = filename
        if settings.get("backup_images", False):
            img_path = os.path.join(settings["save_folder"], "Backup_Images", f"Step_{step}_{capture_count}.png")
        else:
            pdf_temp_dir = os.path.join(current_dir, "_pdf_temp_images")
            os.makedirs(pdf_temp_dir, exist_ok=True)
            img_path = os.path.join(pdf_temp_dir, f"temp_pdf_{capture_count}.png")
            import shutil
            shutil.copy(filename, img_path)
            
        capture_history_pdf.append({
            "step": step,
            "detail": detail,
            "timestamp": timestamp,
            "image": img_path
        })
        build_pdf()
        update_count(capture_count)
        show_flash()
        if entry_detail:
            entry_detail.delete(0, "end")
        try: os.remove(filename)
        except: pass
    else:
        # Word Logic
        try:
            sel = word.Selection
            sel.EndKey(Unit=6)
            sel.TypeParagraph()
            sel.Font.Size = 14
            sel.Font.Bold = True
            sel.Font.Color = 0x993300
            sel.TypeText(f"{step}")
            sel.Font.Bold = False
            sel.Font.Color = 0
            sel.TypeParagraph()
            if detail:
                sel.Font.Size = 11
                sel.Font.Italic = True
                sel.Font.Color = 0x666666
                sel.TypeText(f"  {detail}")
                sel.Font.Italic = False
                sel.Font.Color = 0
                sel.TypeParagraph()
            sel.Font.Size = 9
            sel.Font.Color = 0x999999
            sel.TypeText(f"  Captured at: {timestamp}")
            sel.Font.Color = 0
            sel.TypeParagraph()
            shape = sel.InlineShapes.AddPicture(FileName=filename, LinkToFile=False, SaveWithDocument=True)
            if shape.Width > 450:
                ratio = 450 / shape.Width
                shape.Width = 450
                shape.Height = int(shape.Height * ratio)
            sel.TypeParagraph()
            sel.Font.Size = 8
            sel.Font.Color = 0xCCCCCC
            sel.TypeText("─" * 25)
            sel.Font.Color = 0
            sel.TypeParagraph()
            doc.Save()
            update_count(capture_count)
            show_flash()
            if entry_detail:
                entry_detail.delete(0, "end")
        except Exception as e:
            pass
        try: os.remove(filename)
        except: pass


def start_snip():
    global snip_active
    if snip_active: return
    snip_active = True
    
    mon = settings["monitor"]
    
    snip_win = tk.Toplevel(overlay)
    snip_win.attributes("-topmost", True)
    snip_win.overrideredirect(True)
    snip_win.geometry(f"{mon['width']}x{mon['height']}+{mon['left']}+{mon['top']}")
    snip_win.attributes("-alpha", 0.3)
    snip_win.config(cursor="crosshair")
    
    canvas = tk.Canvas(snip_win, bg="black", highlightthickness=0)
    canvas.pack(fill="both", expand=True)
    
    rect_id = None
    start_x = 0
    start_y = 0

    def on_press(e):
        nonlocal start_x, start_y, rect_id
        start_x, start_y = e.x, e.y
        rect_id = canvas.create_rectangle(start_x, start_y, 1, 1, outline='white', width=2, fill="gray")
        
    def on_drag(e):
        if rect_id:
            canvas.coords(rect_id, start_x, start_y, e.x, e.y)
            
    def on_release(e):
        global snip_active, snip_region
        end_x, end_y = e.x, e.y
        snip_win.destroy()
        snip_active = False
        
        x1, x2 = min(start_x, end_x), max(start_x, end_x)
        y1, y2 = min(start_y, end_y), max(start_y, end_y)
        
        if x2 - x1 > 10 and y2 - y1 > 10:
            snip_region = {
                "left": mon["left"] + x1,
                "top": mon["top"] + y1,
                "width": x2 - x1,
                "height": y2 - y1
            }
            capture_event.set()

    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)
    snip_win.bind("<Escape>", lambda e: [snip_win.destroy(), globals().update(snip_active=False)])
    
    snip_win.focus_force()


def start_annotate():
    global annotate_active, custom_annotate_path
    if annotate_active: return
    annotate_active = True
    
    mon = settings["monitor"]
    filename = os.path.join(current_dir, f"_temp_annotate.png")
    sct_img = sct.grab(mon)
    to_png(sct_img.rgb, sct_img.size, output=filename)
    
    anno_win = tk.Toplevel(overlay)
    anno_win.attributes("-topmost", True)
    anno_win.overrideredirect(True)
    anno_win.geometry(f"{mon['width']}x{mon['height']}+{mon['left']}+{mon['top']}")
    
    bg_img = Image.open(filename)
    bg_photo = ImageTk.PhotoImage(bg_img)
    draw_img = bg_img.copy()
    draw_obj = ImageDraw.Draw(draw_img)
    
    canvas = tk.Canvas(anno_win, bg="black", highlightthickness=0, cursor="pencil")
    canvas.pack(fill="both", expand=True)
    canvas.create_image(0, 0, image=bg_photo, anchor="nw")
    canvas.image = bg_photo
    
    draw_data = []
    
    def on_press(e):
        draw_data.clear()
        draw_data.append((e.x, e.y))
        
    def on_drag(e):
        if draw_data:
            x1, y1 = draw_data[-1]
            x2, y2 = e.x, e.y
            canvas.create_line(x1, y1, x2, y2, fill="red", width=4, capstyle=tk.ROUND, smooth=True)
            draw_obj.line([x1, y1, x2, y2], fill="red", width=4)
            draw_data.append((x2, y2))
            
    def on_done(e):
        global annotate_active, custom_annotate_path
        draw_img.save(filename)
        anno_win.destroy()
        annotate_active = False
        custom_annotate_path = filename
        capture_event.set()
        
    def on_cancel(e):
        global annotate_active
        anno_win.destroy()
        annotate_active = False
        try: os.remove(filename)
        except: pass
        
    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    anno_win.bind("<Return>", on_done)
    anno_win.bind("<Escape>", on_cancel)
    
    lbl = tk.Label(anno_win, text="✏️ โหมดวาดเขียน: วาดบนหน้าจอ กด [Enter] เพื่อเซฟ หรือ กด [Esc] เพื่อยกเลิก", bg="#FBBF24", fg="black", font=("Segoe UI", 12, "bold"), padx=10, pady=5)
    lbl.place(x=20, y=20)
    
    anno_win.focus_force()

def update_step_value(major_diff=0, minor_diff=0):
    if overlay_entry_step is None: return
    current = overlay_entry_step.get().strip()
    if re.match(r'^\d+\.\d+$', current):
        parts = current.split('.')
        maj = int(parts[0])
        min_val = int(parts[1])
        if major_diff != 0:
            maj = max(1, maj + major_diff)
            min_val = 1
        if minor_diff != 0:
            min_val = max(1, min_val + minor_diff)
        new_val = f"{maj}.{min_val}"
    else:
        match = re.search(r'^(.*?)(\d+)$', current)
        if match:
            prefix = match.group(1)
            num_str = match.group(2)
            num = int(num_str)
            diff = major_diff if major_diff != 0 else minor_diff
            new_num = max(1, num + diff)
            new_val = f"{prefix}{new_num:0{len(num_str)}d}"
        else:
            new_val = current
            
    overlay_entry_step.delete(0, "end")
    overlay_entry_step.insert(0, new_val)

def check_events():
    while not step_queue.empty():
        action, diff = step_queue.get()
        if action == 'major': update_step_value(major_diff=diff)
        elif action == 'minor': update_step_value(minor_diff=diff)

    if stop_event.is_set():
        try:
            overlay.quit()
            overlay.destroy()
        except: pass
        os._exit(0)
        return

    if 'capture_region_event' in globals() and capture_region_event.is_set():
        capture_region_event.clear()
        start_snip()

    if 'capture_annotate_event' in globals() and capture_annotate_event.is_set():
        capture_annotate_event.clear()
        start_annotate()

    if capture_event.is_set():
        capture_event.clear()
        region = globals().get('snip_region', None)
        c_path = globals().get('custom_annotate_path', None)
        process_capture(region=region, custom_image_path=c_path)
        if 'snip_region' in globals(): globals()['snip_region'] = None
        if 'custom_annotate_path' in globals(): globals()['custom_annotate_path'] = None
        if 'snip_region' in globals():
            globals()['snip_region'] = None

    overlay.after(100, check_events)

create_overlay()

