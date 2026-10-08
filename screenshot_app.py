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
    "hk_major_next": "shift+up",
    "hk_major_prev": "shift+down",
    "hk_minor_next": "alt+right",
    "hk_minor_prev": "alt+left"
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
    app.geometry("900x600")
    app.minsize(800, 500)
    app.resizable(True, True)

    app.update_idletasks()
    x = (app.winfo_screenwidth() // 2) - (900 // 2)
    y = (app.winfo_screenheight() // 2) - (600 // 2)
    app.geometry(f"+{x}+{y}")

    # Header
    header_frame = ctk.CTkFrame(app, fg_color="transparent")
    header_frame.pack(fill="x", padx=20, pady=(20, 10))
    title = ctk.CTkLabel(header_frame, text="✨ Screenshot Tool Setup", font=ctk.CTkFont(size=28, weight="bold"))
    title.pack(side="left")

    main_container = ctk.CTkFrame(app, fg_color="transparent")
    main_container.pack(fill="both", expand=True, padx=20, pady=5)

    left_col = ctk.CTkFrame(main_container, fg_color="transparent")
    left_col.pack(side="left", fill="both", expand=True, padx=(0, 10))

    right_col = ctk.CTkFrame(main_container, fg_color="transparent")
    right_col.pack(side="right", fill="both", expand=True, padx=(10, 0))

    # --- Section 1: หน้าจอ ---
    frame_mon = ctk.CTkFrame(left_col, corner_radius=15, fg_color="#2B2B2B")
    frame_mon.pack(fill="both", expand=True)
    
    ctk.CTkLabel(frame_mon, text="🖥️ เลือกหน้าจอ", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20, pady=(15, 5))
    
    # Scrollable frame for monitor list
    monitor_scroll = ctk.CTkScrollableFrame(frame_mon, fg_color="transparent")
    monitor_scroll.pack(fill="both", expand=True, padx=10, pady=10)

    selected_monitor_num = ctk.IntVar(value=win_monitor_list[0]["win_num"])
    monitor_buttons = {}

    def on_monitor_select(num):
        selected_monitor_num.set(num)
        for n, btn in monitor_buttons.items():
            if n == num:
                btn.configure(fg_color="#2FA572", border_color="#1F7A52", border_width=2, text_color="white")
            else:
                btn.configure(fg_color="#3A3A3A", border_color="#222222", border_width=1, text_color="#CCCCCC")

    for m in win_monitor_list:
        tag = "⭐ จอหลัก (Primary)" if m["is_primary"] else f"จอ {m['win_num']}"
        resolution = f"{m['width']} x {m['height']}"
        pos = f"X: {m['left']}, Y: {m['top']}"
        btn_text = f"{tag}\nResolution: {resolution}\nPosition: {pos}"
        
        btn = ctk.CTkButton(
            monitor_scroll, text=btn_text, corner_radius=10,
            font=ctk.CTkFont(size=14), height=80,
            command=lambda num=m['win_num']: on_monitor_select(num)
        )
        btn.pack(fill="x", pady=5, padx=10)
        monitor_buttons[m['win_num']] = btn
        
    default_mon_num = saved_monitor_num if saved_monitor_num else next((m for m in win_monitor_list if m["is_primary"]), win_monitor_list[0])["win_num"]
    if default_mon_num not in [m["win_num"] for m in win_monitor_list]:
        default_mon_num = win_monitor_list[0]["win_num"]
    on_monitor_select(default_mon_num)

    # --- Section 2: ปลายทางไฟล์ ---
    frame_file = ctk.CTkFrame(right_col, corner_radius=15, fg_color="#2B2B2B")
    frame_file.pack(fill="x", pady=(0, 10))
    
    ctk.CTkLabel(frame_file, text="📁 พื้นที่จัดเก็บไฟล์", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20, pady=(15, 5))
    
    folder_frame = ctk.CTkFrame(frame_file, fg_color="transparent")
    folder_frame.pack(fill="x", padx=20, pady=(5, 5))
    
    entry_folder = ctk.CTkEntry(folder_frame, height=35)
    entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 10))
    entry_folder.insert(0, settings["save_folder"])
    
    def browse_folder():
        folder = filedialog.askdirectory(initialdir=settings["save_folder"])
        if folder:
            report_folder = os.path.join(folder, "TEST REPORT")
            if not os.path.exists(report_folder):
                try: os.makedirs(report_folder)
                except: pass
            entry_folder.delete(0, "end")
            entry_folder.insert(0, os.path.abspath(report_folder))
            
    btn_browse = ctk.CTkButton(folder_frame, text="เลือกโฟลเดอร์", width=100, height=35, command=browse_folder)
    btn_browse.pack(side="left")

    file_frame = ctk.CTkFrame(frame_file, fg_color="transparent")
    file_frame.pack(fill="x", padx=20, pady=(5, 15))
    ctk.CTkLabel(file_frame, text="ชื่อไฟล์:").pack(side="left", padx=(0, 10))
    entry_filename = ctk.CTkEntry(file_frame, height=35)
    entry_filename.pack(side="left", fill="x", expand=True)
    entry_filename.insert(0, settings["save_filename"])

    # --- Section 3: Hotkeys ---
    frame_hk = ctk.CTkFrame(right_col, corner_radius=15, fg_color="#2B2B2B")
    frame_hk.pack(fill="both", expand=True)
    
    ctk.CTkLabel(frame_hk, text="⌨️ ตั้งค่าปุ่มลัด (Hotkeys)", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20, pady=(15, 5))
    
    hk_scroll = ctk.CTkScrollableFrame(frame_hk, fg_color="transparent")
    hk_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def create_hk_row(parent, label_text, key_name):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(row, text=label_text, width=160, anchor="w", font=ctk.CTkFont(size=13)).pack(side="left")
        
        btn_hk = ctk.CTkButton(row, text=settings[key_name], width=150, height=35, fg_color="#444444", hover_color="#555555")
        btn_hk.pack(side="left", fill="x", expand=True)
        
        def listen_hotkey():
            btn_hk.configure(text="กดปุ่มที่ต้องการ...", fg_color="#D32F2F", hover_color="#B71C1C")
            app.update()
            
            def hotkey_thread():
                try:
                    hk = keyboard.read_hotkey(suppress=False)
                    app.after(0, lambda: finish_listen(hk))
                except Exception as e:
                    app.after(0, lambda: finish_listen(settings[key_name]))

            def finish_listen(hk):
                settings[key_name] = hk
                btn_hk.configure(text=hk, fg_color="#444444", hover_color="#555555")

            threading.Thread(target=hotkey_thread, daemon=True).start()

        btn_hk.configure(command=listen_hotkey)
        return btn_hk

    btn_hk_cap = create_hk_row(hk_scroll, "📸 แคปหน้าจอ:", "hk_capture")
    btn_hk_maj_next = create_hk_row(hk_scroll, "⬆️ ข้อหลัก ถัดไป (2.1):", "hk_major_next")
    btn_hk_maj_prev = create_hk_row(hk_scroll, "⬇️ ข้อหลัก ก่อนหน้า:", "hk_major_prev")
    btn_hk_min_next = create_hk_row(hk_scroll, "➡️ ข้อย่อย ถัดไป (1.2):", "hk_minor_next")
    btn_hk_min_prev = create_hk_row(hk_scroll, "⬅️ ข้อย่อย ก่อนหน้า:", "hk_minor_prev")
    
    ctk.CTkLabel(hk_scroll, text="*คลิกที่ปุ่มแล้วกดปุ่มลัดที่ต้องการบนคีย์บอร์ดเพื่อเปลี่ยน", font=ctk.CTkFont(size=11, slant="italic"), text_color="gray").pack(anchor="w", padx=10, pady=(10, 10))

    # --- Start Button ---
    bottom_frame = ctk.CTkFrame(app, fg_color="transparent")
    bottom_frame.pack(fill="x", padx=20, pady=15)
    
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

    btn_start = ctk.CTkButton(bottom_frame, text="🚀 เริ่มทำงาน (Start)", font=ctk.CTkFont(size=18, weight="bold"), height=50, corner_radius=10, fg_color="#2FA572", hover_color="#1F7A52", command=on_start)
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

    overlay = ctk.CTk()
    overlay.title("Screenshot Overlay")
    # Increase height to accommodate buttons
    overlay.geometry("320x350")
    
    # ทำให้หน้าต่างลอยอยู่เสมอและไม่มีแถบข้างบน (Modern look)
    overlay.attributes("-topmost", True)
    overlay.overrideredirect(True)
    overlay.attributes("-alpha", 0.95)
    
    # สร้างกรอบนอกสุดเพื่อทำขอบสวยๆ
    main_frame = ctk.CTkFrame(overlay, corner_radius=15, border_width=2, border_color="#2FA572")
    main_frame.pack(fill="both", expand=True, padx=2, pady=2)
    
    # ═══ Drag Bar (แถมให้ลากหน้าต่างได้) ═══
    drag_bar = ctk.CTkFrame(main_frame, height=35, corner_radius=10, fg_color="#1E1E1E")
    drag_bar.pack(fill="x", padx=5, pady=5)
    
    lbl_title = ctk.CTkLabel(drag_bar, text="📸 Screenshot Tool", font=ctk.CTkFont(size=14, weight="bold"), text_color="#2FA572")
    lbl_title.pack(side="left", padx=10)
    
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
        
    # Bind to drag_bar and its children
    drag_bar.bind("<Button-1>", on_drag_start)
    drag_bar.bind("<B1-Motion>", on_drag_motion)
    lbl_title.bind("<Button-1>", on_drag_start)
    lbl_title.bind("<B1-Motion>", on_drag_motion)
    
    def ui_close():
        stop_event.set()
        
    btn_close = ctk.CTkButton(drag_bar, text="✖", width=25, height=25, corner_radius=5, fg_color="#D32F2F", hover_color="#B71C1C", text_color="white", command=ui_close)
    btn_close.pack(side="right", padx=5)
    
    # ═══ Step Display with Buttons ═══
    step_frame = ctk.CTkFrame(main_frame, fg_color="#2B2B2B", corner_radius=10)
    step_frame.pack(fill="x", padx=15, pady=(5, 5))
    
    # Major Controls
    row_major = ctk.CTkFrame(step_frame, fg_color="transparent")
    row_major.pack(fill="x", padx=10, pady=(10, 5))
    
    def on_major_prev(): prev_major()
    def on_major_next(): next_major()
    
    ctk.CTkButton(row_major, text="◀ ข้อหลัก", width=70, height=30, command=on_major_prev, fg_color="#444444", hover_color="#555555", font=ctk.CTkFont(weight="bold")).pack(side="left")
    ctk.CTkButton(row_major, text="ข้อหลัก ▶", width=70, height=30, command=on_major_next, fg_color="#444444", hover_color="#555555", font=ctk.CTkFont(weight="bold")).pack(side="right")
    
    overlay_label_step = ctk.CTkLabel(step_frame, text=f"ข้อ {get_step_label()}", font=ctk.CTkFont(size=36, weight="bold"), text_color="#FFFFFF")
    overlay_label_step.pack(pady=5)
    
    # Minor Controls
    row_minor = ctk.CTkFrame(step_frame, fg_color="transparent")
    row_minor.pack(fill="x", padx=10, pady=(0, 10))
    
    def on_minor_prev(): prev_minor()
    def on_minor_next(): next_minor()
    
    ctk.CTkButton(row_minor, text="◁ ข้อย่อย", width=70, height=30, command=on_minor_prev, fg_color="#555555", hover_color="#666666").pack(side="left")
    ctk.CTkButton(row_minor, text="ข้อย่อย ▷", width=70, height=30, command=on_minor_next, fg_color="#555555", hover_color="#666666").pack(side="right")

    # ═══ Detail Input ═══
    entry_detail = ctk.CTkEntry(main_frame, placeholder_text="พิมพ์รายละเอียดข้อเทสที่นี่...", width=280, height=35, font=ctk.CTkFont(size=13))
    entry_detail.pack(padx=15, pady=5)
    
    # ═══ Action Buttons ═══
    action_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    action_frame.pack(fill="x", padx=15, pady=(5, 0))
    
    def on_btn_capture(): on_capture()
    ctk.CTkButton(action_frame, text="📸 แคปรูปเดี๋ยวนี้", height=35, font=ctk.CTkFont(weight="bold"), fg_color="#2FA572", hover_color="#1F7A52", command=on_btn_capture).pack(fill="x")

    # ═══ Info ═══
    info_text = f"ตั้งค่าแคป: {settings['hk_capture']} | กดปุ่มในแอปได้เลย"
    ctk.CTkLabel(main_frame, text=info_text, font=ctk.CTkFont(size=11), text_color="#AAAAAA").pack(pady=(5, 0))
    
    label_count = ctk.CTkLabel(main_frame, text="แคปแล้ว: 0 ภาพ", font=ctk.CTkFont(size=11, weight="bold"), text_color="#2FA572")
    label_count.pack()

    # วางไว้มุมขวาบน
    overlay.update_idletasks()
    screen_w = overlay.winfo_screenwidth()
    overlay.geometry(f"+{screen_w - 340}+20")

    # เริ่ม Event polling
    overlay.after(100, check_events)
    overlay.mainloop()

def update_count(count):
    if label_count:
        label_count.configure(text=f"แคปแล้ว: {count} ภาพ")

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
            overlay_label_step.configure(text=f"ข้อ {get_step_label()}")

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
