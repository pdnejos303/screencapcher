import datetime
import os
import sys
import threading
import keyboard
import win32com.client
import win32api
from mss import mss
from mss.tools import to_png

sys.stdout.reconfigure(line_buffering=True)

print("=" * 60)
print("   Screenshot to Word - เครื่องมือแคปหน้าจอลง Word")
print("=" * 60)

# ค้นหาหน้าจอทั้งหมดจาก Windows API (ให้หมายเลขตรงกับ Windows Display Settings)
print("\nกำลังค้นหาหน้าจอจาก Windows...")

sct = mss()
mss_monitors = sct.monitors  # mss_monitors[0] = รวมทุกจอ, [1]+ = แต่ละจอ

# ดึงข้อมูลจอจาก Windows API
win_monitors_raw = win32api.EnumDisplayMonitors(None, None)
win_monitor_list = []

for hMonitor, hdcMonitor, pyRect in win_monitors_raw:
    info = win32api.GetMonitorInfo(hMonitor)
    device_name = info["Device"]  # เช่น \\.\DISPLAY1, \\.\DISPLAY2
    # ดึงหมายเลขจอจาก Windows (ตัวเลขท้ายชื่อ DISPLAY)
    win_num = int(device_name.replace("\\\\.\\DISPLAY", ""))
    rect = info["Monitor"]  # (left, top, right, bottom)
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]
    
    # หา mss monitor index ที่ตรงกับพิกัดนี้
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

# เรียงตามหมายเลข Windows
win_monitor_list.sort(key=lambda x: x["win_num"])

print(f"\nพบทั้งหมด {len(win_monitor_list)} หน้าจอ (หมายเลขตรงกับ Windows Display Settings):")
for m in win_monitor_list:
    primary_tag = " ⭐ (จอหลัก)" if m["is_primary"] else ""
    print(f"  [{m['win_num']}] หน้าจอ {m['win_num']} - {m['width']}x{m['height']}{primary_tag}")

# ถามผู้ใช้ว่าจะแคปจอไหน
valid_nums = [m["win_num"] for m in win_monitor_list]
if len(win_monitor_list) == 1:
    chosen = win_monitor_list[0]
    print(f"\nมีจอเดียว เลือกหน้าจอ {chosen['win_num']} อัตโนมัติ")
else:
    while True:
        try:
            num = int(input(f"\nเลือกหมายเลขหน้าจอ ({', '.join(map(str, valid_nums))}): "))
            if num in valid_nums:
                chosen = next(m for m in win_monitor_list if m["win_num"] == num)
                break
            print(f"หมายเลขไม่ถูกต้อง กรุณาเลือกจาก {valid_nums}")
        except ValueError:
            print("กรุณาพิมพ์ตัวเลข")

monitor_to_capture = mss_monitors[chosen["mss_idx"]]
print(f"✅ เลือกหน้าจอ {chosen['win_num']} ({chosen['width']}x{chosen['height']}) เรียบร้อย!")

# เปิด Word
print("\nกำลังเปิด Microsoft Word...")
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

    print("✅ เปิด Word สำเร็จ! (ตรวจสอบที่ Taskbar ด้านล่าง)")
except Exception as e:
    print(f"❌ เปิด Word ไม่ได้: {e}")
    input("กด Enter เพื่อปิด...")
    sys.exit(1)

current_dir = os.path.abspath(os.path.dirname(__file__))
capture_count = 0

# ใช้ Event เพื่อส่งสัญญาณจาก hotkey thread กลับมาที่ main thread
capture_event = threading.Event()
stop_event = threading.Event()


def on_hotkey():
    capture_event.set()


def on_esc():
    stop_event.set()
    capture_event.set()


keyboard.add_hotkey("shift+s", on_hotkey)
keyboard.add_hotkey("esc", on_esc)

print("\n" + "=" * 60)
print("🎯 พร้อมใช้งาน!")
print(f"   กด  Shift + S  ➔ แคปหน้าจอ {chosen['win_num']} แล้วรูปเด้งเข้า Word ทันที")
print("   กด  ESC        ➔ ปิดโปรแกรม")
print("")
print("   💡 Word เปิดค้างไว้ได้เลย! ไม่ต้องปิด!")
print("      กด Shift+S กี่ทีก็ได้ รูปจะต่อท้ายกันไปเรื่อยๆ")
print("      เมื่อพอใจแล้วค่อยกด Ctrl+S ใน Word เพื่อบันทึกไฟล์")
print("=" * 60)

# Main loop - ทำงานบน main thread เพื่อไม่ให้เกิด COM error
while not stop_event.is_set():
    capture_event.wait()
    capture_event.clear()

    if stop_event.is_set():
        break

    capture_count += 1
    print(f"\n📸 [{capture_count}] กำลังแคปหน้าจอ {chosen['win_num']}...")

    filename = os.path.join(current_dir, f"_temp_screenshot_{capture_count}.png")
    sct_img = sct.grab(monitor_to_capture)
    to_png(sct_img.rgb, sct_img.size, output=filename)

    try:
        sel = word.Selection
        sel.EndKey(Unit=6)
        sel.TypeParagraph()

        sel.Font.Size = 10
        sel.Font.Color = 8421504
        sel.TypeText(
            f"Captured #{capture_count} at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        sel.TypeParagraph()
        sel.Font.Color = 0

        shape = sel.InlineShapes.AddPicture(
            FileName=filename, LinkToFile=False, SaveWithDocument=True
        )

        if shape.Width > 450:
            ratio = 450 / shape.Width
            shape.Width = 450
            shape.Height = int(shape.Height * ratio)

        sel.TypeParagraph()

        print(
            f"✅ [{capture_count}] สำเร็จ! รูปเข้าไปอยู่ใน Word แล้ว (ดูได้เลยในหน้าต่าง Word)"
        )

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาด: {e}")

    try:
        os.remove(filename)
    except:
        pass

print("\nปิดโปรแกรมเรียบร้อยแล้ว")
