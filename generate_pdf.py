import os
import math
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

from pdf_positions import POSITIONS

# =========================
# 테스트 설정
# =========================
TEST_MODE = False
TEST_LIMIT_GROUP = 1
TEST_ALL_QUARTERS = True

# =========================
# 기본 설정
# =========================
INPUT_EXCEL = os.path.join("input", "generate_pdf", "results_pdf.xlsx")
OUTPUT_DIR = os.path.join("output", "generate_pdf")
BG_IMAGE = "assets/form_bg.png"

APPLICATION_QUARTER = "1분기"

MAX_COURSE_PER_PAGE = 5
MAX_TOTAL_COURSE = 15

FONT_NAME = "HYSMyeongJo-Medium"

os.makedirs(OUTPUT_DIR, exist_ok=True)
pdfmetrics.registerFont(UnicodeCIDFont(FONT_NAME))


# =========================
# 텍스트 출력
# =========================
def draw_text(c, text, x, y, size=10):
    if pd.isna(text) or str(text).strip() == "":
        return
    c.setFont(FONT_NAME, size)
    c.drawString(x, y, str(text))


# =========================
# 주소 처리 함수 (신규 추가)
# =========================
def draw_address(c, address):

    if pd.isna(address) or str(address).strip() == "":
        return

    address = str(address).strip()

    # 괄호 포함 여부 체크
    if "(" in address and ")" in address:

        # 괄호 기준 분리
        main_part = address.split("(")[0].strip()
        bracket_part = "(" + address.split("(")[1]

        draw_text(c, main_part, *POSITIONS["상세 주소_괄호_있음_1"], size=9)
        draw_text(c, bracket_part, *POSITIONS["상세 주소_괄호_있음_2"], size=9)

    else:
        draw_text(c, address, *POSITIONS["상세 주소_괄호_없음"], size=9)


# =========================
# 개인정보 검증
# =========================
def validate_primary_row(row):
    required_fields = [
        "주민번호",
        "전화번호",
        "은행명",
        "예금주",
        "계좌번호",
        "상세 주소",
    ]

    missing = []
    for field in required_fields:
        if pd.isna(row[field]) or str(row[field]).strip() == "":
            missing.append(field)

    return missing


# =========================
# PDF 생성
# =========================
def create_pdf(user_row, course_rows, page_index, quarter):

    filename = f"{quarter}_감면신청서_{user_row['이름']}_{user_row['회원번호']}_{page_index}.pdf"
    filepath = os.path.join(OUTPUT_DIR, filename)

    c = canvas.Canvas(filepath, pagesize=A4)

    # 배경 이미지
    c.drawImage(BG_IMAGE, 0, 0, width=A4[0], height=A4[1])

    # 분기 체크
    if quarter in POSITIONS:
        draw_text(c, "■", *POSITIONS[quarter], size=18)
    else:
        print(f"[ERROR] 분기 키 없음: {quarter}")

    # 기본 정보
    draw_text(c, user_row["이름"], *POSITIONS["이름"])
    draw_text(c, user_row["주민번호"], *POSITIONS["주민번호"])
    draw_text(c, user_row["회원번호"], *POSITIONS["회원번호"])
    draw_text(c, user_row["전화번호"], *POSITIONS["전화번호"])

    # 🔥 수정된 주소 처리
    draw_address(c, user_row["상세 주소"])

    # 강좌 영역
    for idx, row in enumerate(course_rows):
        draw_text(c, row["강좌"], *POSITIONS["강좌"][idx])
        draw_text(c, f"{int(row['금액']):,}", *POSITIONS["금액"][idx])
        draw_text(c, f"{int(row['감면액']):,}", *POSITIONS["감면액"][idx])

    # 계좌 정보
    draw_text(c, user_row["은행명"], *POSITIONS["은행명"])
    draw_text(c, user_row["예금주"], *POSITIONS["예금주"])
    draw_text(c, user_row["계좌번호"], *POSITIONS["계좌번호"])

    c.showPage()
    c.save()

    print(f"✅ 생성 완료: {filename}")


# =========================
# 메인
# =========================
def main():
    df = pd.read_excel(INPUT_EXCEL)

    grouped = list(df.groupby(["회원번호", "이름"], sort=False))

    if TEST_MODE:
        grouped = grouped[:TEST_LIMIT_GROUP]
        print(f"\n[TEST MODE] 상위 {TEST_LIMIT_GROUP}명 생성\n")

    for (member_no, name), group in grouped:

        total_courses = len(group)

        if total_courses > MAX_TOTAL_COURSE:
            print(f"[ERROR] {name}({member_no}) 강좌 {total_courses}개 → 15개 초과")
            continue

        primary_rows = group[group["등록건 수"] == 1]

        if len(primary_rows) != 1:
            print(f"[ERROR] {name}({member_no}) 등록건 수 1 행 {len(primary_rows)}개")
            continue

        primary_row = primary_rows.iloc[0]

        missing_fields = validate_primary_row(primary_row)

        if missing_fields:
            print(f"[ERROR] {name}({member_no}) 개인정보 누락: {missing_fields}")
            continue

        courses = group.sort_values("등록건 수")
        total_pages = math.ceil(total_courses / MAX_COURSE_PER_PAGE)

        quarters = (
            ["1분기", "2분기", "3분기", "4분기"]
            if TEST_MODE and TEST_ALL_QUARTERS
            else [APPLICATION_QUARTER]
        )

        for quarter in quarters:
            for page in range(total_pages):
                start = page * MAX_COURSE_PER_PAGE
                end = start + MAX_COURSE_PER_PAGE

                page_courses = courses.iloc[start:end].to_dict("records")

                create_pdf(primary_row, page_courses, page + 1, quarter)


if __name__ == "__main__":
    main()
