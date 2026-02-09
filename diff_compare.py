import pandas as pd
from collections import defaultdict
import os
import sys
import re


# =========================
# 파일 자동 선택 함수
# =========================
def find_compare_files():
    # data_preprocessing 실행 후 얻은 결과 파일 경로
    base_dir = os.path.join("output", "data_preprocessing")

    if not os.path.isdir(base_dir):
        print(f"❌ 디렉토리가 존재하지 않습니다: {base_dir}")
        sys.exit(1)

    pattern = re.compile(r"(\d{4}-\d{2}-\d{2})_output\.xlsx")

    files = []
    for f in os.listdir(base_dir):
        m = pattern.match(f)
        if m:
            date = pd.to_datetime(m.group(1))
            files.append((date, f))

    if len(files) < 2:
        print("❌ 비교 가능한 파일이 2개 이상 필요합니다.")
        sys.exit(1)

    # 날짜 기준 정렬: 최근 일자 정렬 후 2개 선택(prev: 최근 직전, curr: 가장 최근)
    files.sort(key=lambda x: x[0])

    prev_file = os.path.join(base_dir, files[-2][1])
    curr_file = os.path.join(base_dir, files[-1][1])

    return prev_file, curr_file


# =========================
# 설정
# =========================
PREV_FILE, CURR_FILE = find_compare_files()
OUTPUT_FILE = f"output/diff_compare/{os.path.basename(PREV_FILE)}↔{os.path.basename(CURR_FILE)}_비교결과.xlsx"

KEY_COLS = ["회원번호", "이름"]

# =========================
# 데이터 로드
# =========================
prev = pd.read_excel(PREV_FILE)
curr = pd.read_excel(CURR_FILE)

print("📄 비교 파일 선택 완료")
print(f" - 이전: {os.path.basename(PREV_FILE)}")
print(f" - 현재: {os.path.basename(CURR_FILE)}")
print(f" - 이전 데이터: {len(prev)}건")
print(f" - 현재 데이터: {len(curr)}건")

# =========================
# KEY 생성
# =========================
prev["KEY"] = prev[KEY_COLS].astype(str).agg("|".join, axis=1)
curr["KEY"] = curr[KEY_COLS].astype(str).agg("|".join, axis=1)

# =========================
# KEY 단위 그룹핑
# =========================
prev_grp = prev.groupby("KEY")
curr_grp = curr.groupby("KEY")

all_keys = set(prev_grp.groups.keys()) | set(curr_grp.groups.keys())

results = []

# =========================
# 비교 로직
# =========================
for key in all_keys:
    prev_rows = (
        prev_grp.get_group(key).copy() if key in prev_grp.groups else pd.DataFrame()
    )
    curr_rows = (
        curr_grp.get_group(key).copy() if key in curr_grp.groups else pd.DataFrame()
    )

    member_no, name = key.split("|")

    # ---- 이전만 있음 → 삭제
    if curr_rows.empty:
        for _, r in prev_rows.iterrows():
            results.append(
                {
                    "구분": "삭제",
                    "수강처": r["수강처"],
                    "회원번호": member_no,
                    "이름": name,
                    "등록건 수": r["등록건 수"],
                    "강좌": r["강좌"],
                    "금액": r["금액"],
                    "감면액": r["감면액"],
                    "강좌(이전)": r["강좌"],
                    "강좌(현재)": "",
                    "금액(이전)": r["금액"],
                    "금액(현재)": "",
                }
            )
        continue

    # ---- 현재만 있음 → 추가
    if prev_rows.empty:
        for _, r in curr_rows.iterrows():
            results.append(
                {
                    "구분": "추가",
                    "수강처": r["수강처"],
                    "회원번호": member_no,
                    "이름": name,
                    "등록건 수": r["등록건 수"],
                    "강좌": r["강좌"],
                    "금액": r["금액"],
                    "감면액": r["감면액"],
                    "강좌(이전)": "",
                    "강좌(현재)": r["강좌"],
                    "금액(이전)": "",
                    "금액(현재)": r["금액"],
                }
            )
        continue

    # ---- 매칭용 복사본
    prev_used = set()
    curr_used = set()

    # 1️⃣ 금액 동일 + 강좌 동일 → 패스
    for pi, pr in prev_rows.iterrows():
        for ci, cr in curr_rows.iterrows():
            if ci in curr_used or pi in prev_used:
                continue
            if pr["강좌"] == cr["강좌"] and pr["금액"] == cr["금액"]:
                prev_used.add(pi)
                curr_used.add(ci)
                break

    # 2️⃣ 금액 동일 + 강좌 다름 → 강좌 변경 추정
    for pi, pr in prev_rows.iterrows():
        if pi in prev_used:
            continue
        for ci, cr in curr_rows.iterrows():
            if ci in curr_used:
                continue
            if pr["금액"] == cr["금액"]:
                results.append(
                    {
                        "구분": "강좌변경(추정)",
                        "수강처": cr["수강처"],
                        "회원번호": member_no,
                        "이름": name,
                        "등록건 수": cr["등록건 수"],
                        "강좌": cr["강좌"],
                        "금액": cr["금액"],
                        "감면액": cr["감면액"],
                        "강좌(이전)": pr["강좌"],
                        "강좌(현재)": cr["강좌"],
                        "금액(이전)": pr["금액"],
                        "금액(현재)": cr["금액"],
                    }
                )
                prev_used.add(pi)
                curr_used.add(ci)
                break

    # 3️⃣ 금액 변경 (강좌 동일)
    for pi, pr in prev_rows.iterrows():
        if pi in prev_used:
            continue
        for ci, cr in curr_rows.iterrows():
            if ci in curr_used:
                continue
            if pr["강좌"] == cr["강좌"] and pr["금액"] != cr["금액"]:
                results.append(
                    {
                        "구분": "금액변경",
                        "수강처": cr["수강처"],
                        "회원번호": member_no,
                        "이름": name,
                        "등록건 수": cr["등록건 수"],
                        "강좌": cr["강좌"],
                        "금액": cr["금액"],
                        "감면액": cr["감면액"],
                        "강좌(이전)": pr["강좌"],
                        "강좌(현재)": cr["강좌"],
                        "금액(이전)": pr["금액"],
                        "금액(현재)": cr["금액"],
                    }
                )
                prev_used.add(pi)
                curr_used.add(ci)
                break

    # 4️⃣ 남은 이전 → 삭제
    for pi, pr in prev_rows.iterrows():
        if pi not in prev_used:
            results.append(
                {
                    "구분": "삭제",
                    "수강처": cr["수강처"],
                    "회원번호": member_no,
                    "이름": name,
                    "등록건 수": cr["등록건 수"],
                    "강좌": cr["강좌"],
                    "금액": cr["금액"],
                    "감면액": cr["감면액"],
                    "강좌(이전)": pr["강좌"],
                    "강좌(현재)": "",
                    "금액(이전)": pr["금액"],
                    "금액(현재)": "",
                }
            )

    # 5️⃣ 남은 현재 → 추가
    for ci, cr in curr_rows.iterrows():
        if ci not in curr_used:
            results.append(
                {
                    "구분": "추가",
                    "수강처": cr["수강처"],
                    "회원번호": member_no,
                    "이름": name,
                    "등록건 수": cr["등록건 수"],
                    "강좌": cr["강좌"],
                    "금액": cr["금액"],
                    "감면액": cr["감면액"],
                    "강좌(이전)": "",
                    "강좌(현재)": cr["강좌"],
                    "금액(이전)": "",
                    "금액(현재)": cr["금액"],
                }
            )

# =========================
# 결과 저장
# =========================
result_df = pd.DataFrame(results).sort_values(["이름", "구분"]).reset_index(drop=True)

if result_df.empty:
    print("✅ 변경 사항 없음")
else:
    result_df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ 비교 완료 → {OUTPUT_FILE} 생성")
