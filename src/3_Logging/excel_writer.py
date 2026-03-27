import pandas as pd
from pathlib import Path
from inference import process_row   # hàm đã xây dựng ở __init__.py


# ── HÀM CHÍNH ──────────────────────────────────────────────────────────────────
def run_inference_and_save(df: pd.DataFrame, filename: str = "email_analyzed.xlsx") -> pd.DataFrame:
    """
    Lặp qua từng hàng của DataFrame đã clean, gọi model AI để lấy
    category / urgency / summary, gán 3 cột mới vào DataFrame rồi
    xuất ra file Excel.

    Parameters
    ----------
    df       : pd.DataFrame  — DataFrame đầu ra của cleandata() (có cột Subject, Body, ...)
    filename : str           — Tên file Excel đầu ra

    Returns
    -------
    pd.DataFrame  — DataFrame đã có thêm 3 cột AI_Category, AI_Urgency, AI_Summary
    """
    categories: list[str] = []
    urgencies:  list[str] = []
    summaries:  list[str] = []

    total = len(df)
    print(f"[excel_writer] Bắt đầu phân tích {total} email...")

    for idx, (_, row) in enumerate(df.iterrows(), start=1):
        category, urgency, summary = process_row(row)
        categories.append(category)
        urgencies.append(urgency)
        summaries.append(summary)

        # Log tiến trình mỗi 50 hàng
        if idx % 50 == 0 or idx == total:
            print(f"  [{idx}/{total}] đã xử lý")

    # Gán 3 cột mới vào DataFrame gốc
    df = df.copy()
    df["AI_Category"] = categories
    df["AI_Urgency"]  = urgencies
    df["AI_Summary"]  = summaries

    # ── Xuất Excel ────────────────────────────────────────────────────────────
    output_dir = Path(__file__).resolve().parents[2] / "output" / "excel_log"
    output_dir.mkdir(parents=True, exist_ok=True)
    excel_path = output_dir / filename

    df.to_excel(excel_path, engine="openpyxl", index=False)
    print(f"[excel_writer] Đã lưu Excel → {excel_path}")

    return df