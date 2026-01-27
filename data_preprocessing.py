from datetime import datetime, timedelta
from collections import defaultdict
from math import floor
import pandas as pd
import sys
import re
import os


# =====================
# csv 자동 탐색
# =====================
def find_csv_file():
    csv_files = [f for f in os.listdir(".") if f.lower().endswith(".csv")]

    if not csv_files:
        print("❌ 현재 디렉토리에 CSV 파일이 없습니다.")
        sys.exit(1)

    if len(csv_files) > 1:
        print("❌ CSV 파일이 여러 개 있습니다:")
        for f in csv_files:
            print(" -", f)
        print("CSV 파일을 하나만 남겨주세요.")
        sys.exit(1)

    return csv_files[0]


# =====================
# 설정
# =====================
INPUT_CSV = find_csv_file()
yesterday_str = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
OUTPUT_XLSX = f"{yesterday_str}_output.xlsx"

BASE_AMOUNTS = [60000, 50000, 70000, 40000, 80000, 90000, 110000]
RATIOS = [1.0, 0.9, 0.5, 0.45]

WHITELIST = {int(base * ratio) for base in BASE_AMOUNTS for ratio in RATIOS}

DETAIL_REGEX = re.compile(
    r"\(\s*(?P<payment>[^-]+)-(?P<status>[^\)]+)\)"
    r"\((?P<member>W?\d+)\)\s*(?P<name>.+)"
)

print(f"📄 입력 CSV: {INPUT_CSV}")
print(f"📄 출력 Excel: {OUTPUT_XLSX}")


# =====================
# 유틸
# =====================
def extract_course(text):
    m = re.search(r"\[(.*?)\]", str(text))
    return m.group(1).split()[0] if m else "오기"


def extract_amount(cash, card):
    cash = int(str(cash).replace(",", ""))
    card = int(str(card).replace(",", ""))
    return cash if cash != 0 else card


def refund_amount(raw):
    """
    24250 → 25000
    """
    return floor((raw * 100 / 97) / 10) * 10


# =====================
# 감면액 계산 유틸 (추가)
# =====================
DISCOUNT_40_BASE = [60000, 75000, 90000, 105000, 120000, 135000, 165000]
# 30% 그룹 → 원금(40_BASE) 매핑
DISCOUNT_30_TO_ORIGINAL = {
    67500: 75000,
    81000: 90000,
    94500: 105000,
}

# 분할 수강(1/3, 2/3, 1) 고려한 세트
DISCOUNT_40_SET = {
    int(base * r) for base in DISCOUNT_40_BASE for r in (1 / 3, 2 / 3, 1)
}
DISCOUNT_30_SET = {
    int(base * r) for base in DISCOUNT_30_TO_ORIGINAL.keys() for r in (1 / 3, 2 / 3, 1)
}


def calc_discount(amount: int, reg_count: int) -> int:
    """
    amount    : 실제 결제 금액
    reg_count : 등록건 수
    """

    # ---------------------
    # 감면율 결정
    # ---------------------
    if reg_count == 1:
        rate_40 = 0.4
        rate_30 = 0.3
    elif reg_count == 2:
        rate_40 = 0.45
        rate_30 = 0.35
    else:  # 3건 이상
        rate_40 = 0.5
        rate_30 = 0.4

    # ---------------------
    # 40% 베이스
    # ---------------------
    if amount in DISCOUNT_40_SET:
        return int(round(amount * rate_40))

    # ---------------------
    # 30% 베이스 (원금 기준 계산)
    # ---------------------
    if amount in DISCOUNT_30_SET:
        for discounted, original in DISCOUNT_30_TO_ORIGINAL.items():
            ratio = amount / discounted

            # 부동소수점 오차 안전 처리
            if abs(ratio - 1 / 3) < 0.001:
                original_amount = original / 3
            elif abs(ratio - 2 / 3) < 0.001:
                original_amount = original * 2 / 3
            elif abs(ratio - 1) < 0.001:
                original_amount = original
            else:
                continue

            return int(round(original_amount * rate_30))

    return 0


# =====================
# 데이터 로딩
# =====================
df = pd.read_csv(INPUT_CSV, sep=None, engine="python", encoding="cp949")

# =====================
# 데이터 구조
# =====================
records = defaultdict(list)  # (회원번호, 강좌) → [등록행...]
refund_rows = []  # (key, payment, amount)

# =====================
# 1차 패스
# records = {
#     ("10000000", "서예A"): [
#         {
#             "수강처": "상도",
#             "회원번호": "10000000",
#             "이름": "김철수",
#             "강좌": "서예A",
#             "금액": 75000
#         }
#     ]
# }
# # =====================
for _, row in df.iterrows():
    course = extract_course(row["일반/운영여부"])
    if not course:
        continue

    m = DETAIL_REGEX.search(str(row["세부내역"]))
    if not m:
        continue

    payment = m.group("payment").strip()
    status = m.group("status").strip()
    member = m.group("member")
    name = m.group("name").strip()

    if payment == "무입금":
        continue

    amount = extract_amount(row["현금"], row["카드"])
    key = (member, course)

    if status == "환불":
        refund_rows.append((key, payment, amount))
        continue

    if status in ("신규", "재등록"):
        records[key].append(
            {
                "수강처": row["수강처"],
                "회원번호": member,
                "이름": name,
                "강좌": course,
                "금액": amount,
            }
        )

# =====================
# 2차 패스: 환불 반영
# refund_rows = [
#     (("10000000", "서예"), "현금", -24250)
# ]
# =====================
for key, payment, amount in refund_rows:
    if key not in records or not records[key]:
        continue

    last_record = records[key][-1]

    # ---------------------
    # 현금 부분환불
    # ---------------------
    if payment == "현금":
        calc = refund_amount(abs(amount))

        if calc in WHITELIST:
            last_record["금액"] -= calc
            continue

    # ---------------------
    # 전액환불
    # ---------------------
    records[key].pop()
    if not records[key]:
        del records[key]

# =====================
# 결과 생성
# =====================
final_rows = []
for rows in records.values():
    final_rows.extend(rows)

result_df = pd.DataFrame(final_rows)
result_df = result_df.sort_values(["이름", "회원번호", "금액"]).reset_index(drop=True)

# 🔧 핵심 수정 2: 회원번호 + 이름 기준 등록건 수 계산
result_df["등록건 수"] = result_df.groupby(["회원번호", "이름"]).cumcount().add(1)

# =====================
# 감면액 컬럼 추가
# =====================
result_df["감면액"] = result_df.apply(
    lambda row: calc_discount(row["금액"], row["등록건 수"]), axis=1
)


result_df = result_df[
    ["수강처", "회원번호", "이름", "등록건 수", "강좌", "금액", "감면액"]
]

result_df.to_excel(OUTPUT_XLSX, index=False)

print("✅ 처리 완료")
print(f"📊 최종 생성 행 수: {len(result_df)}")
