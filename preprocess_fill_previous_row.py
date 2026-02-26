from datetime import datetime
import pandas as pd
import os


def preprocess_fill_previous_row():
    # ============================
    # 경로 설정
    # ============================
    input_dir = "input/preprocess_fill_previous_row"
    output_dir = "output/preprocess_fill_previous_row"
    today_str = datetime.today().strftime("%Y-%m-%d")

    input_filename = "동작문화원 감면신청서_26년도 1분기_20260225.csv"
    output_filename = f"동작문화원 감면신청서_26년도 1분기_output_{today_str}.xlsx"

    input_path = os.path.join(input_dir, input_filename)
    output_path = os.path.join(output_dir, output_filename)

    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"❌ 입력 파일이 존재하지 않습니다: {input_path}")

    # ============================
    # CSV 로딩 (🔥 전부 문자열로)
    # ============================
    df = pd.read_csv(input_path, encoding="cp949", dtype=str)

    # 컬럼명 공백 제거
    df.columns = df.columns.str.strip()

    # ============================
    # 문자열 정리 (🔥 중요)
    # ============================

    df["회원번호"] = df["회원번호"].astype(str).str.strip()
    df["이름"] = df["이름"].astype(str).str.strip()
    df["등록건 수"] = df["등록건 수"].astype(str).str.strip()

    # Unnamed 컬럼 제거
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    # ============================
    # 필요한 컬럼 정의
    # ============================
    columns = [
        "수강처",
        "회원번호",
        "이름",
        "등록건 수",
        "강좌",
        "금액",
        "감면액",
        "주민번호",
        "전화번호",
        "은행명",
        "예금주",
        "계좌번호",
        "동작구민 여부",
        "상세 주소",
        "비고",
        "신청자-예금주 일치여부",
    ]

    # 🔥 컬럼 존재 확인
    missing_cols = [col for col in columns if col not in df.columns]
    if missing_cols:
        print("❌ 누락된 컬럼:", missing_cols)
        print("현재 컬럼 목록:", list(df.columns))
        raise ValueError("필수 컬럼이 누락되었습니다.")

    df = df[columns]

    fill_columns = columns[7:]  # H~P

    filled_count = 0

    # ============================
    # 그룹 처리
    # ============================
    grouped = df.groupby(["회원번호", "이름"], dropna=False)

    for (member_no, name), group in grouped:

        # 등록건 수 == "1" (문자열 대응)
        base_rows = group[group["등록건 수"].astype(str).str.strip() == "1"]

        if base_rows.empty:
            continue

        base_row = base_rows.iloc[0]

        for idx in group.index:

            if str(df.loc[idx, "등록건 수"]).strip() == "1":
                continue

            for col in fill_columns:

                current_value = df.loc[idx, col]

                if pd.isna(current_value) or str(current_value).strip() == "":
                    df.loc[idx, col] = base_row[col]
                    filled_count += 1

    # ============================
    # 저장
    # ============================
    os.makedirs(output_dir, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")

    print("\n✅ 전처리 완료")
    print(f"📄 저장 위치: {output_path}")
    print(f"📊 총 채워진 셀 개수: {filled_count}")


if __name__ == "__main__":
    preprocess_fill_previous_row()
