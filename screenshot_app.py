import datetime
import json
import os
import sys
import threading
from tkinter import filedialog, messagebox

import customtkinter as ctk
import keyboard
import win32api
import win32com.client
from mss import mss
from mss.tools import to_png

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
    "hk_capture": "shift+s",
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
            for k in ["save_folder", "hk_capture", "hk_major_next", "hk_major_prev", "hk_minor_next", "hk_minor_prev"]:
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
    app.geometry("1000x650")
    app.minsize(900, 600)
    
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
    y = (app.winfo_screenheight() // 2) - (650 // 2)
    app.geometry(f"+{x}+{y}")

    # Header
    header_frame = ctk.CTkFrame(app, fg_color="transparent")
    header_frame.pack(fill="x", padx=40, pady=(30, 20))
    
    title_lbl = ctk.CTkLabel(header_frame, text="Screenshot Automation", font=(FONT_FAMILY, 32, "bold"), text_color=COLOR_TEXT)
    title_lbl.pack(side="left")
    
    subtitle_lbl = ctk.CTkLabel(header_frame, text="ตั้งค่าหน้าจอและปุ่มลัดก่อนเริ่มการทำงาน", font=(FONT_FAMILY, 14), text_color=COLOR_TEXT_MUTED)
    subtitle_lbl.pack(side="left", padx=15, pady=6)

    # Main Grid Container
    main_container = ctk.CTkFrame(app, fg_color="transparent")
    main_container.pack(fill="both", expand=True, padx=40, pady=10)

    left_col = ctk.CTkFrame(main_container, fg_color="transparent")
    left_col.pack(side="left", fill="both", expand=True, padx=(0, 15))

    right_col = ctk.CTkFrame(main_container, fg_color="transparent")
    right_col.pack(side="right", fill="both", expand=True, padx=(15, 0))

    # --- Section 1: หน้าจอ ---
    frame_mon = ctk.CTkFrame(left_col, corner_radius=16, fg_color=COLOR_CARD, border_width=1, border_color="#2E3C56")
    frame_mon.pack(fill="both", expand=True)
    
    mon_header = ctk.CTkFrame(frame_mon, fg_color="transparent")
    mon_header.pack(fill="x", padx=25, pady=(25, 10))
    ctk.CTkLabel(mon_header, text="🖥️ เลือกหน้าจอ", font=(FONT_FAMILY, 18, "bold"), text_color=COLOR_TEXT).pack(anchor="w")
    ctk.CTkLabel(mon_header, text="ระบุหน้าจอที่ต้องการแคปเจอร์", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT_MUTED).pack(anchor="w")
    
    monitor_scroll = ctk.CTkScrollableFrame(frame_mon, fg_color="transparent")
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

    # --- Section 2: ปลายทางไฟล์ ---
    frame_file = ctk.CTkFrame(right_col, corner_radius=16, fg_color=COLOR_CARD, border_width=1, border_color="#2E3C56")
    frame_file.pack(fill="x", pady=(0, 20))
    
    file_header = ctk.CTkFrame(frame_file, fg_color="transparent")
    file_header.pack(fill="x", padx=25, pady=(25, 10))
    ctk.CTkLabel(file_header, text="📁 พื้นที่จัดเก็บไฟล์", font=(FONT_FAMILY, 18, "bold"), text_color=COLOR_TEXT).pack(anchor="w")
    ctk.CTkLabel(file_header, text="ที่อยู่สำหรับบันทึกไฟล์รายงาน Word", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT_MUTED).pack(anchor="w")
    
    folder_frame = ctk.CTkFrame(frame_file, fg_color="transparent")
    folder_frame.pack(fill="x", padx=25, pady=(10, 5))
    
    entry_folder = ctk.CTkEntry(folder_frame, height=40, font=(FONT_FAMILY, 13), corner_radius=8, border_color="#2E3C56", fg_color="#0B0F19")
    entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 10))
    entry_folder.insert(0, settings["save_folder"])
    
    def browse_folder():
        folder = filedialog.askdirectory(initialdir=settings["save_folder"])
        if folder:
            report_folder = os.path.join(folder, "TEST_REPORT")
            if not os.path.exists(report_folder):
                try: os.makedirs(report_folder)
                except: pass
            entry_folder.delete(0, "end")
            entry_folder.insert(0, os.path.abspath(report_folder))
            
    btn_browse = ctk.CTkButton(folder_frame, text="Browse", width=100, height=40, corner_radius=8, fg_color="#2E3C56", hover_color="#3B4B68", font=(FONT_FAMILY, 13, "bold"), command=browse_folder)
    btn_browse.pack(side="left")

    file_frame = ctk.CTkFrame(frame_file, fg_color="transparent")
    file_frame.pack(fill="x", padx=25, pady=(10, 25))
    ctk.CTkLabel(file_frame, text="ชื่อไฟล์:", font=(FONT_FAMILY, 13, "bold")).pack(side="left", padx=(0, 15))
    entry_filename = ctk.CTkEntry(file_frame, height=40, font=(FONT_FAMILY, 13), corner_radius=8, border_color="#2E3C56", fg_color="#0B0F19")
    entry_filename.pack(side="left", fill="x", expand=True)
    entry_filename.insert(0, settings["save_filename"])

    # --- Section 3: Hotkeys ---
    frame_hk = ctk.CTkFrame(right_col, corner_radius=16, fg_color=COLOR_CARD, border_width=1, border_color="#2E3C56")
    frame_hk.pack(fill="both", expand=True)
    
    hk_header = ctk.CTkFrame(frame_hk, fg_color="transparent")
    hk_header.pack(fill="x", padx=25, pady=(25, 5))
    ctk.CTkLabel(hk_header, text="⌨️ ตั้งค่าปุ่มลัด (Hotkeys)", font=(FONT_FAMILY, 18, "bold"), text_color=COLOR_TEXT).pack(anchor="w")
    ctk.CTkLabel(hk_header, text="คลิกที่ปุ่มแล้วกดปุ่มบนคีย์บอร์ดที่ต้องการ", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT_MUTED).pack(anchor="w")
    
    hk_scroll = ctk.CTkScrollableFrame(frame_hk, fg_color="transparent")
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

    btn_hk_cap = create_hk_row(hk_scroll, "📸 แคปหน้าจอ", "hk_capture")
    btn_hk_maj_next = create_hk_row(hk_scroll, "⬆️ ข้อหลักถัดไป (+1.0)", "hk_major_next")
    btn_hk_maj_prev = create_hk_row(hk_scroll, "⬇️ ข้อหลักก่อนหน้า (-1.0)", "hk_major_prev")
    btn_hk_min_next = create_hk_row(hk_scroll, "➡️ ข้อย่อยถัดไป (+0.1)", "hk_minor_next")
    btn_hk_min_prev = create_hk_row(hk_scroll, "⬅️ ข้อย่อยก่อนหน้า (-0.1)", "hk_minor_prev")

    # --- Start Button ---
    bottom_frame = ctk.CTkFrame(app, fg_color="transparent")
    bottom_frame.pack(fill="x", padx=40, pady=(15, 30))
    
    def on_start():
        global setup_done
        mon_num = selected_monitor_num.get()
        settings["monitor"] = next(m for m in win_monitor_list if m["win_num"] == mon_num)
        
        settings["save_folder"] = entry_folder.get()
        filename = entry_filename.get()
        if not filename.endswith(".docx"):
            filename += ".docx"
        settings["save_filename"] = filename
        
        setup_done = True
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                save_data = {
                    "monitor_num": mon_num,
                    "save_folder": settings["save_folder"],
                    "save_filename": settings["save_filename"],
                    "hk_capture": settings["hk_capture"],
                    "hk_major_next": settings["hk_major_next"],
                    "hk_major_prev": settings["hk_major_prev"],
                    "hk_minor_next": settings["hk_minor_next"],
                    "hk_minor_prev": settings["hk_minor_prev"]
                }
                json.dump(save_data, f, ensure_ascii=False, indent=4)
        except: pass
        app.destroy()

    btn_start = ctk.CTkButton(bottom_frame, text="🚀 เปิด Overlay เริ่มทำงาน", font=(FONT_FAMILY, 16, "bold"), height=55, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, command=on_start)
    btn_start.pack(fill="x")

    app.mainloop()

run_setup_ui()

if not setup_done:
    sys.exit(0)

monitor_to_capture = mss_monitors[settings["monitor"]["mss_idx"]]
full_save_path = os.path.join(settings["save_folder"], settings["save_filename"])

# ════════════════════════════════════════════════════════════════
# 3. เปิด Word
# ════════════════════════════════════════════════════════════════
try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()

    sel = word.Selection
    sel.Font.Size = 20
    sel.Font.Bold = True
    sel.TypeText("Automation Screenshot Report\n")
    sel.Font.Size = 11
    sel.Font.Bold = False
    sel.TypeParagraph()

    doc.SaveAs2(full_save_path)
except Exception as e:
    messagebox.showerror("Error", f"ไม่สามารถเปิด Microsoft Word ได้:\n{e}")
    sys.exit(1)


# ════════════════════════════════════════════════════════════════
# 4. ตัวแปรสถานะ
# ════════════════════════════════════════════════════════════════
current_dir = os.path.abspath(os.path.dirname(__file__))
capture_count = 0

major = 1
minor = 1

capture_event = threading.Event()
stop_event = threading.Event()
state_lock = threading.Lock()

def get_step_label():
    return f"{major}.{minor}"

# ════════════════════════════════════════════════════════════════
# 5. Hotkey callbacks
# ════════════════════════════════════════════════════════════════
def on_capture(): capture_event.set()
def on_stop():
    stop_event.set()
    capture_event.set()

step_changed_event = threading.Event()

def next_minor():
    global minor
    with state_lock: minor += 1
    step_changed_event.set()

def prev_minor():
    global minor
    with state_lock:
        if minor > 1: minor -= 1
    step_changed_event.set()

def next_major():
    global major, minor
    with state_lock:
        major += 1
        minor = 1
    step_changed_event.set()

def prev_major():
    global major, minor
    with state_lock:
        if major > 1:
            major -= 1
            minor = 1
    step_changed_event.set()

try:
    keyboard.add_hotkey(settings["hk_capture"], on_capture)
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
overlay_label_step = None
entry_detail = None
label_count = None

def create_overlay():
    global overlay, overlay_label_step, entry_detail, label_count

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
    overlay.attributes("-alpha", 0.95)
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
    
    def ui_close(): stop_event.set()
        
    btn_close = ctk.CTkButton(drag_bar, text="✕", width=30, height=30, corner_radius=8, fg_color="transparent", hover_color="#EF4444", text_color=COLOR_TEXT_MUTED, font=(FONT_FAMILY, 14, "bold"), command=ui_close)
    btn_close.pack(side="right", padx=5)
    
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
    overlay_label_step = ctk.CTkLabel(ctrl_row, text=f"{get_step_label()}", font=(FONT_FAMILY, 42, "bold"), text_color=COLOR_ACCENT)
    overlay_label_step.pack(side="left", expand=True)
    
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
    
    btn_cap = ctk.CTkButton(action_frame, text="📸 แคปรูปเดี๋ยวนี้", height=45, corner_radius=12, font=(FONT_FAMILY, 15, "bold"), fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, command=on_btn_capture)
    btn_cap.pack(fill="x")

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
def process_capture():
    global capture_count

    capture_count += 1
    step = get_step_label()
    detail = entry_detail.get().strip() if entry_detail else ""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    filename = os.path.join(current_dir, f"_temp_cap_{capture_count}.png")
    sct_img = sct.grab(monitor_to_capture)
    to_png(sct_img.rgb, sct_img.size, output=filename)

    try:
        sel = word.Selection
        sel.EndKey(Unit=6)
        sel.TypeParagraph()

        # หมายเลขข้อ
        sel.Font.Size = 14
        sel.Font.Bold = True
        sel.Font.Color = 0x993300
        sel.TypeText(f"ข้อ {step}")
        sel.Font.Bold = False
        sel.Font.Color = 0
        sel.TypeParagraph()

        # รายละเอียด
        if detail:
            sel.Font.Size = 11
            sel.Font.Italic = True
            sel.Font.Color = 0x666666
            sel.TypeText(f"  {detail}")
            sel.Font.Italic = False
            sel.Font.Color = 0
            sel.TypeParagraph()

        # เวลา
        sel.Font.Size = 9
        sel.Font.Color = 0x999999
        sel.TypeText(f"  Captured at: {timestamp}")
        sel.Font.Color = 0
        sel.TypeParagraph()

        # รูปภาพ
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

        # ล้างช่องรายละเอียดหลังแคปเสร็จ (Optional: เพื่อให้พร้อมพิมพ์ข้อต่อไป)
        if entry_detail:
            entry_detail.delete(0, "end")

    except Exception as e:
        pass

    try: os.remove(filename)
    except: pass

def check_events():
    # อัปเดตข้อความบน UI หากมีการเปลี่ยนข้อ
    if step_changed_event.is_set():
        step_changed_event.clear()
        if overlay_label_step:
            overlay_label_step.configure(text=f"{get_step_label()}")

    if stop_event.is_set():
        try:
            overlay.quit()
            overlay.destroy()
        except: pass
        os._exit(0)
        return

    if capture_event.is_set():
        capture_event.clear()
        process_capture()

    overlay.after(100, check_events)

create_overlay()
