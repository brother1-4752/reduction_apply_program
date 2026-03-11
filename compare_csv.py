import pandas as pd


def compare_files(file_a_csv, file_b_xlsx):

    # CSV (엑셀 저장 CSV는 보통 cp949)
    df_a = pd.read_csv(file_a_csv, encoding="cp949")
    df_b = pd.read_excel(file_b_xlsx)

    # B,C열만 사용 (회원번호, 이름)
    df_a = df_a.iloc[:, 1:3]
    df_b = df_b.iloc[:, 1:3]

    # 2행부터 비교
    df_a = df_a.iloc[1:]
    df_b = df_b.iloc[1:]

    max_len = max(len(df_a), len(df_b))

    mismatches = []

    for i in range(max_len):

        row_num = i + 2

        try:
            a_row = df_a.iloc[i].tolist()
        except:
            a_row = ["", ""]

        try:
            b_row = df_b.iloc[i].tolist()
        except:
            b_row = ["", ""]

        if a_row != b_row:
            mismatches.append((row_num, a_row, b_row))

    print(f"총 비교 행: {max_len}")
    print(f"불일치 행: {len(mismatches)}")

    for row, a, b in mismatches:
        print(f"\n{row}행 불일치")
        print("A:", a)
        print("B:", b)


if __name__ == "__main__":

    compare_files(
        "input/compare_csv/A.csv",
        "output/data_preprocessing/2026-03-09_output.xlsx"
    )