"""
Pipeline chính: CSV → clean → AI inference → Excel + Word report
"""

from data_loader import CsvLoader
from preprocessing import cleandata          # __init__.py của bước preprocessing
from excel_writer  import run_inference_and_save
from stats_calc import generate_report


def main():
    # ── BƯỚC 1: Load dữ liệu thô ─────────────────────────────────────────────
    print("=" * 60)
    print("BƯỚC 1 — Load dữ liệu CSV")
    print("=" * 60)
    loader     = CsvLoader(data_size=700, random_state=42)
    raw_series = loader.load_messages()

    # ── BƯỚC 2: Làm sạch email ────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("BƯỚC 2 — Làm sạch dữ liệu email")
    print("=" * 60)
    df_clean = cleandata(raw_series)
    print(f"  → DataFrame sau clean: {len(df_clean)} hàng, {list(df_clean.columns)} cột")

    # ── BƯỚC 3: AI Inference + xuất Excel ────────────────────────────────────
    print("\n" + "=" * 60)
    print("BƯỚC 3 — AI Inference & xuất Excel")
    print("=" * 60)
    df_analyzed = run_inference_and_save(df_clean, filename="email_analyzed.xlsx")

    # ── BƯỚC 4: Sinh báo cáo Word ─────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("BƯỚC 4 — Sinh báo cáo Word")
    print("=" * 60)
    stats, exec_summary = generate_report(df_analyzed, filename="email_report.docx")

    # ── HOÀN THÀNH ────────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("PIPELINE HOÀN THÀNH")
    print("=" * 60)
    print(f"  ✓ Excel  → output/excel_log/email_analyzed.xlsx")
    print(f"  ✓ Report → {doc_path}")


if __name__ == "__main__":
    main()