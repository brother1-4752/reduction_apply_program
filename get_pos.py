import pyautogui
import keyboard
import json
from datetime import datetime

positions = {}
count = 1

print("좌표 수집 시작")
print("Ctrl + Alt + S : 좌표 저장")
print("ESC : 종료")

def save_position():
    global count
    x, y = pyautogui.position()
    name = f"pos_{count}"
    positions[name] = {"x": x, "y": y}
    print(f"저장됨 ▶ {name}: ({x}, {y})")
    count += 1

keyboard.add_hotkey('ctrl+alt+s', save_position)
keyboard.wait('esc')

filename = f"positions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
with open(filename, "w", encoding="utf-8") as f:
    json.dump(positions, f, ensure_ascii=False, indent=2)

print(f"파일 저장 완료: {filename}")
