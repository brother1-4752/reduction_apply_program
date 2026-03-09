import pyautogui
import keyboard
import time
import sys

# ▶ 안전 설정
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0


# =========================
# ▶ PC별 좌표 정의
# =========================
POSITIONS = {
    "PC1": {
        # 회원관리시스템
        "회원번호_입력칸": (151, 214),
        "회원_선택": (275, 288),
        # 분기 시트
        "분기_회원번호": (1142, 207),
        "분기_시트": (1248, 999),
        "분기_데이터_시작": (1531, 206),
        # 데이터 가공 시트
        "데이터가공_시트": (1324, 999),
        "회원정보_기본_입력시작점": (1094, 209),
        "회원정보_기본_입력종료점": (1862, 211),
        "회원정보_상세_입력시작점": (1093, 262),
        "회원정보_상세_입력종료점": (1598, 260),
        # 공통
        "아래로": (1910, 983),
    },
    "PC2": {
        "회원번호_입력칸": (141, 216),
        "회원_선택": (281, 286),
        "분기_회원번호": (1135, 199),
        "분기_시트": (1242, 1008),
        "분기_데이터_시작": (1524, 198),
        "데이터가공_시트": (1318, 1007),
        "회원정보_기본_입력시작점": (1087, 203),
        # "회원정보_기본_입력종료점": (1862, 211),
        "회원정보_상세_입력시작점": (1087, 254),
        "회원정보_상세_입력종료점": (1591, 256),
        "아래로": (1911, 993),
    },
    "PC3": {
        "회원번호_입력칸": (145, 215),
        "회원_선택": (279, 287),
        # "분기_회원번호": (1135, 199),
        # "분기_시트": (1242, 1008),
        # "분기_데이터_시작": (1524, 198),
        # "데이터가공_시트": (1318, 1007),
        # "회원정보_기본_입력시작점": (1087, 203),
        # "회원정보_기본_입력종료점": (1862, 211),
        # "회원정보_상세_입력시작점": (1087, 254),
        # "회원정보_상세_입력종료점": (1591, 256),
        # "아래로": (1911, 993),
    },
}


# =========================
# ▶ PC 환경 선택
# =========================
print("🖥 실행할 PC 환경을 선택하세요")
print("1 : PC1")
print("2 : PC2")

pc_input = input("▶ 선택 (1 / 2): ").strip()

if pc_input == "1":
    pos = POSITIONS["PC1"]
    print("✅ PC1 세팅으로 실행합니다.")
elif pc_input == "2":
    pos = POSITIONS["PC2"]
    print("✅ PC2 세팅으로 실행합니다.")
else:
    print("❌ 잘못된 입력입니다. 프로그램 종료")
    sys.exit(1)


# =========================
# ▶ 공통 동작 함수
# =========================
def step_wait():
    time.sleep(1)


def below_click(name):
    print(f"  → 클릭: {name}")
    x, y = pos[name]
    pyautogui.click(x, y, duration=0.5)
    time.sleep(0.3)


def click(name):
    print(f"  → 클릭: {name}")
    x, y = pos[name]
    pyautogui.click(x, y, duration=0.5)
    time.sleep(1.05)
    pyautogui.click()


def shift_click(start, end):
    print(f"  → Shift 범위 선택: {start} ~ {end}")

    x1, y1 = pos[start]
    pyautogui.click(x1, y1, duration=1)
    time.sleep(0.3)

    x2, y2 = pos[end]
    pyautogui.moveTo(x2, y2, duration=1)
    time.sleep(0.3)

    pyautogui.keyDown("shift")
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.1)
    pyautogui.keyUp("shift")
    time.sleep(0.5)


def ctrl(key):
    time.sleep(0.1)
    pyautogui.hotkey("ctrl", key)
    time.sleep(0.2)


def format_duration(seconds):
    seconds = int(seconds)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60

    result = []
    if h > 0:
        result.append(f"{h:02d}시")
    if m > 0:
        result.append(f"{m:02d}분")
    result.append(f"{s:02d}초")

    return " ".join(result) + " 소요"


def paste_values_to_other_sheet():
    # 1) 일반 붙여넣기 (Ctrl+V)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)

    # 2) 값 붙여넣기 전환: Ctrl → V 시퀀스
    # Ctrl 한 번 "톡" 눌렀다가 떼면 Paste 옵션 아이콘이 활성화됨
    pyautogui.keyDown("ctrl")
    time.sleep(0.05)
    pyautogui.keyUp("ctrl")
    time.sleep(0.15)

    # 그 다음 V를 눌러서 "값만" 선택
    pyautogui.press("v")
    time.sleep(0.2)


# =========================
# ▶ 매크로 스텝 정의
# =========================
def get_steps():
    return [
        ("1-1. [EXCEL] 분기_시트 선택", lambda: click("분기_시트")),
        (
            "1-2. [EXCEL] 회원번호 선택 + 복사",
            lambda: (click("분기_회원번호"), step_wait(), ctrl("c")),
        ),
        (
            "2-1. [회원관리시스템] CMS 회원번호 입력",
            lambda: (
                click("회원번호_입력칸"),
                time.sleep(0.3),
                ctrl("v"),
                time.sleep(0.3),
                pyautogui.press("enter"),
            ),
        ),
        (
            "3-1. [회원관리시스템] 회원선택 복사",
            lambda: (click("회원_선택"), step_wait(), ctrl("c")),
        ),
        ("4-1. [EXCEL] 데이터 가공 시트 선택", lambda: click("데이터가공_시트")),
        (
            "4-2. [EXCEL] 이전 데이터 클리어링",
            lambda: (
                shift_click("회원정보_기본_입력시작점", "회원정보_기본_입력종료점"),
                step_wait(),
                pyautogui.keyDown("delete"),
                time.sleep(0.5),
                pyautogui.keyUp("delete", time.sleep(0.15)),
            ),
        ),
        (
            "4-2. [EXCEL] 회원정보 붙여넣기",
            lambda: (click("회원정보_기본_입력시작점"), step_wait(), ctrl("v")),
        ),
        (
            "5-1. [EXCEL] 최종 데이터 시작점 선택",
            lambda: click("회원정보_상세_입력시작점"),
        ),
        (
            "5-2. [EXCEL] 최종 데이터 범위 선택 + 복사",
            lambda: (
                shift_click("회원정보_상세_입력시작점", "회원정보_상세_입력종료점"),
                ctrl("c"),
            ),
        ),
        # (
        #     "5-3. [EXCEL] Enter",
        #     lambda: (
        #         pyautogui.press("enter"),
        #         time.sleep(0.5),
        #         pyautogui.press("enter"),
        #     ),
        # ),
        ("6-1. [EXCEL] 분기 시트 선택", lambda: (click("분기_시트"), time.sleep(0.5))),
        (
            "6-2. [EXCEL] 최종데이터 붙여넣기",
            lambda: (
                click("분기_데이터_시작"),
                time.sleep(0.5),
                paste_values_to_other_sheet(),
            ),
        ),
        # (
        #     "6-4. [EXCEL] Enter",
        #     lambda: (
        #         pyautogui.press("enter"),
        #         time.sleep(0.5),
        #         pyautogui.press("enter"),
        #     ),
        # ),
        ("6-4. [EXCEL] 아래 행으로 이동", lambda: below_click("아래로")),
        ("7-1. [EXCEL] 저장", lambda: ctrl("s")),
    ]


# =========================
# ▶ 디버그 / 자동 실행
# =========================
def debug_macro():
    steps = get_steps()
    step = 0

    print("\n🚀 디버그 매크로 시작")
    print("ENTER: 다음 | B: 이전 | R: 재실행 | Q: 종료")

    while 0 <= step < len(steps):
        desc, action = steps[step]
        print(f"\n[STEP {step + 1}/{len(steps)}] {desc}")

        cmd = input("▶ 명령: ").strip().lower()

        if cmd == "":
            action()
            step_wait()
            step += 1
        elif cmd == "b":
            step = max(step - 1, 0)
        elif cmd == "r":
            action()
        elif cmd == "q":
            break


def auto_macro():
    steps = get_steps()

    try:
        repeat_count = int(input("\n🔁 반복 횟수 입력: "))
    except ValueError:
        print("❌ 숫자만 입력하세요")
        return

    start = time.time()

    for r in range(1, repeat_count + 1):
        print(f"\n🔄 반복 {r}/{repeat_count}")
        for desc, action in steps:
            print(desc)
            action()
            step_wait()

    elapsed = time.time() - start
    print("\n✅ 자동 실행 완료")
    print(f"⏱ 총 소요 시간: {format_duration(elapsed)}")


# =========================
# ▶ 실행 대기
# =========================
print("\n🛠 매크로 대기중")
print("F8 : 디버그 실행")
print("F9 : 자동 실행")
print("ESC : 종료")

keyboard.add_hotkey("F8", debug_macro)
keyboard.add_hotkey("F9", auto_macro)

keyboard.wait("esc")
