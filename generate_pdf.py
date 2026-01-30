import os
import math
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

from pdf_positions import POSITIONS

# =========================
# 기본 설정
# =========================
INPUT_EXCEL = os.path.join("input", "generate_pdf", "results_pdf.xlsx")
BG_IMAGE = "assets/form_bg.png"
OUTPUT_DIR = os.path.join("output", "generate_pdf")

APPLICATION_QUARTER = "2026년 1분기"
MAX_COURSE_PER_PAGE = 5
MAX_TOTAL_COURSE = 15

FONT_NAME = "HYSMyeongJo-Medium"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================
# 한글 폰트 등록
# =========================
pdfmetrics.registerFont(UnicodeCIDFont(FONT_NAME))


# =========================
# 공통 텍스트 출력 함수
# =========================
def draw_text(c, text, x, y, size=10):
    if pd.isna(text):
        return
    c.setFont(FONT_NAME, size)
    c.drawString(x, y, str(text))


# =========================
# PDF 1장 생성
# =========================
def create_pdf(user, courses, page_index):
    filename = f"감면신청서_{user['이름']}_{user['회원번호']}_{page_index}.pdf"
    filepath = os.path.join(OUTPUT_DIR, filename)

    c = canvas.Canvas(filepath, pagesize=A4)

    # 배경 이미지
    c.drawImage(BG_IMAGE, 0, 0, width=A4[0], height=A4[1])

    # 상단 정보 (모든 페이지 동일)
    draw_text(c, APPLICATION_QUARTER, *POSITIONS["quarter"])
    draw_text(c, user["이름"], *POSITIONS["name"])
    draw_text(c, user["주민번호"], *POSITIONS["resident_id"])
    draw_text(c, user["회원번호"], *POSITIONS["member_no"])
    draw_text(c, user["전화번호"], *POSITIONS["phone"])
    draw_text(c, user["상세 주소"], *POSITIONS["address"], size=9)

    # 강좌 영역
    for idx, course in enumerate(courses):
        draw_text(c, course["강좌"], *POSITIONS["course"][idx])
        draw_text(c, f"{int(course['금액']):,}", *POSITIONS["price"][idx])

    # 계좌 정보
    draw_text(c, user["은행명"], *POSITIONS["bank"])
    draw_text(c, user["예금주"], *POSITIONS["account_holder"])
    draw_text(c, user["계좌번호"], *POSITIONS["account_number"])

    c.showPage()
    c.save()

    print(f"✅ 생성 완료: {filename}")


# =========================
# 메인 실행
# =========================
def main():
    df = pd.read_excel(INPUT_EXCEL)

    # PK 기준 groupby
    grouped = df.groupby(["회원번호", "이름"], sort=False)

    for (member_no, name), group in grouped:
        total_courses = len(group)

        # 🔴 검증
        if total_courses > MAX_TOTAL_COURSE:
            raise ValueError(
                f"[ERROR] {name}({member_no}) 강좌 수 {total_courses}개 → 최대 15개 초과"
            )

        # 공통 정보 (첫 행 기준)
        user = group.iloc[0]

        # 강좌 리스트
        courses = group[["강좌", "금액"]].to_dict("records")

        total_pages = math.ceil(total_courses / MAX_COURSE_PER_PAGE)

        for page in range(total_pages):
            start = page * MAX_COURSE_PER_PAGE
            end = start + MAX_COURSE_PER_PAGE
            page_courses = courses[start:end]

            create_pdf(user, page_courses, page + 1)


if __name__ == "__main__":
    main()
