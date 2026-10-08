import datetime
import os
import keyboard
import pyautogui
import win32com.client

print("กำลังเชื่อมต่อกับ Microsoft Word...")
try:
    # สั่งเปิดโปรแกรม Word ขึ้นมา
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # ทำให้หน้าต่าง Word โชว์ขึ้นมาบนหน้าจอ
except Exception as e:
    print(f"เกิดข้อผิดพลาดในการเปิด Word: {e}")
    exit()

# สร้างเอกสารใหม่
doc = word.Documents.Add()
selection = word.Selection
selection.TypeText("Automation Screenshot Report\n")
selection.Style = word.ActiveDocument.Styles("Heading 1")
selection.TypeParagraph()

def capture_and_save():
    print("📸 กำลังแคปหน้าจอ...")
    
    # สร้าง Path สำหรับเซฟรูปชั่วคราว
    current_dir = os.path.abspath(os.path.dirname(__file__))
    filename = os.path.join(current_dir, f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
    
    # แคปหน้าจอและเซฟเป็นไฟล์ชั่วคราว
    screenshot = pyautogui.screenshot()
    screenshot.save(filename)
    
    # เลื่อนเคอร์เซอร์ใน Word ไปที่บรรทัดล่างสุด (wdStory = 6)
    selection.EndKey(Unit=6)
    selection.TypeParagraph()
    selection.TypeText(f"Captured at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # แทรกรูปภาพเข้าไปใน Word สดๆ
    shape = selection.InlineShapes.AddPicture(FileName=filename, LinkToFile=False, SaveWithDocument=True)
    
    # ปรับขนาดรูปไม่ให้ใหญ่เกินไป (ความกว้างมาตรฐานกระดาษประมาณ 450)
    if shape.Width > 450:
        ratio = 450 / shape.Width
        shape.Width = 450
        shape.Height = shape.Height * ratio
        
    # ลบไฟล์ภาพชั่วคราวทิ้ง เพื่อไม่ให้รกโฟลเดอร์
    try:
        os.remove(filename)
    except:
        pass
    
    print("\n[สำเร็จ] นำรูปเข้าไปใส่ใน Word ให้ดูสดๆ แล้ว!")

# ตั้งค่าปุ่มลัด
keyboard.add_hotkey("shift+s", capture_and_save)

print("="*60)
print("เปิดหน้าต่าง Word ขึ้นมาให้แล้ว! (สามารถย่อ/ขยายจอ Word ไว้ดูข้างๆ ได้)")
print(" - กดปุ่ม 'Shift + S' เพื่อแคปหน้าจอ แล้วรูปจะเด้งเข้าไปใน Word ทันที")
print(" - เมื่อใช้งานเสร็จแล้ว ให้กดปุ่ม 'ESC' เพื่อปิดสคริปต์นี้")
print(" - *อย่าลืมกด Save (Ctrl+S) ในหน้าต่าง Word ด้วยตัวเองเพื่อบันทึกไฟล์นะครับ*")
print("="*60)

keyboard.wait("esc")
print("ปิดโปรแกรมเรียบร้อยแล้ว")
