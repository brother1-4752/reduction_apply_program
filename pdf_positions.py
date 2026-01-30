from reportlab.lib.pagesizes import A4

PAGE_WIDTH, PAGE_HEIGHT = A4

POSITIONS = {
    "quarter": (120, 760),
    "name": (120, 720),
    "resident_id": (300, 720),
    "member_no": (480, 720),
    "phone": (120, 690),
    "address": (120, 660),
    # 강좌 5개
    "course": [
        (120, 580),
        (120, 560),
        (120, 540),
        (120, 520),
        (120, 500),
    ],
    "price": [
        (380, 580),
        (380, 560),
        (380, 540),
        (380, 520),
        (380, 500),
    ],
    "bank": (120, 450),
    "account_holder": (260, 450),
    "account_number": (380, 450),
}
