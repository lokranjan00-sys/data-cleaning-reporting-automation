"""Automated Data Cleaning and Reporting Pipeline.

Reads raw messy sales data, performs automated data preprocessing (handling
missing values, duplicates, and inconsistent formatting), and generates:
1. A styled multi-sheet Excel report (cleaned data, category summary, audit log)
2. A visual summary dashboard image (PNG)
3. A standalone HTML summary report
"""

from pathlib import Path
import html
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


BASE_DIR = Path(__file__).resolve().parent
RAW_DATA_PATH = (
    BASE_DIR / "data" / "raw_sales_data.csv"
    if (BASE_DIR / "data" / "raw_sales_data.csv").exists()
    else BASE_DIR / "raw_sales_data.csv"
)
OUTPUT_DIR = BASE_DIR / "output"

CITY_TYPO_MAP = {
    "Banglore": "Bengaluru",
    "Bangalore": "Bengaluru",
    "Hyd": "Hyderabad",
}

PAYMENT_MAP = {
    "Upi": "UPI",
    "Netbanking": "NetBanking",
}


def load_data(file_path: Path) -> pd.DataFrame:
    """Load raw CSV or Excel data into a DataFrame."""
    if not file_path.exists():
        raise FileNotFoundError(f"Input file not found: {file_path}")

    if file_path.suffix.lower() == ".csv":
        return pd.read_csv(file_path, dtype=str, keep_default_na=True)
    if file_path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(file_path, dtype=str)
    raise ValueError(f"Unsupported file format: {file_path.suffix}")


def parse_dates_safely(series: pd.Series) -> pd.Series:
    """Parse mixed date formats consistently into datetime."""
    cleaned = series.astype(str).str.strip()
    try:
        return pd.to_datetime(cleaned, format="mixed", errors="coerce")
    except TypeError:
        return pd.to_datetime(cleaned, errors="coerce")


def clean_data(raw_df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Clean raw data and return (cleaned_df, audit_metrics)."""
    df = raw_df.copy()

    # Replace common placeholder strings with actual missing values
    df = df.replace(
        to_replace=[r"^\s*$", r"^(?i:n/a|na|null|none|-)$"],
        value=pd.NA,
        regex=True,
    )

    initial_rows = len(df)
    initial_missing = int(df.isna().sum().sum())

    # 1. Standardize text columns (trim whitespace & apply consistent casing)
    text_cols = ["Customer_Name", "Region", "Category", "Product", "Payment_Method"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda val: val.strip().title() if isinstance(val, str) else val
            )

    if "Order_ID" in df.columns:
        df["Order_ID"] = df["Order_ID"].astype(str).str.strip().str.upper()

    # Fix known city typos and payment casing
    if "Region" in df.columns:
        df["Region"] = df["Region"].replace(CITY_TYPO_MAP)
    if "Payment_Method" in df.columns:
        df["Payment_Method"] = df["Payment_Method"].replace(PAYMENT_MAP)

    # 2. Remove exact and Order_ID-based duplicates
    before_dedup = len(df)
    df = df.drop_duplicates()
    if "Order_ID" in df.columns:
        df = df.drop_duplicates(subset=["Order_ID"], keep="first")
    duplicates_removed = before_dedup - len(df)

    # 3. Standardize dates
    df["Order_Date"] = parse_dates_safely(df["Order_Date"])
    invalid_dates_fixed = int(df["Order_Date"].isna().sum())
    if invalid_dates_fixed > 0:
        fallback_date = df["Order_Date"].median()
        df["Order_Date"] = df["Order_Date"].fillna(fallback_date)

    # 4. Clean numeric columns and impute missing values
    for num_col in ["Quantity", "Unit_Price"]:
        cleaned_num = (
            df[num_col]
            .astype(str)
            .str.replace(r"[^\d.]", "", regex=True)
            .replace("", pd.NA)
        )
        df[num_col] = pd.to_numeric(cleaned_num, errors="coerce")

    missing_qty_fixed = int(df["Quantity"].isna().sum())
    missing_price_fixed = int(df["Unit_Price"].isna().sum())

    # Fill missing Quantity with median per Category (fallback overall median)
    df["Quantity"] = df["Quantity"].fillna(
        df.groupby("Category")["Quantity"].transform("median")
    )
    df["Quantity"] = df["Quantity"].fillna(df["Quantity"].median()).round().astype(int)

    # Fill missing Unit_Price with median per Category (fallback overall median)
    df["Unit_Price"] = df["Unit_Price"].fillna(
        df.groupby("Category")["Unit_Price"].transform("median")
    )
    df["Unit_Price"] = df["Unit_Price"].fillna(df["Unit_Price"].median()).round(2)

    # 5. Fill categorical missing values
    df["Customer_Name"] = df["Customer_Name"].fillna("Guest Customer")
    df["Region"] = df["Region"].fillna(df["Region"].mode().iloc[0])
    df["Payment_Method"] = df["Payment_Method"].fillna("Unknown")

    # 6. Derived metric and consistent sorting
    df["Total_Amount"] = (df["Quantity"] * df["Unit_Price"]).round(2)
    df = df.sort_values("Order_Date").reset_index(drop=True)
    df["Order_Date"] = df["Order_Date"].dt.strftime("%Y-%m-%d")

    final_missing = int(df.isna().sum().sum())
    audit = {
        "Initial Rows": initial_rows,
        "Final Clean Rows": len(df),
        "Duplicates Removed": duplicates_removed,
        "Initial Missing Cells": initial_missing,
        "Missing Quantity Imputed": missing_qty_fixed,
        "Missing Price Imputed": missing_price_fixed,
        "Remaining Missing Cells": final_missing,
        "Total Revenue": round(float(df["Total_Amount"].sum()), 2),
        "Average Order Value": round(float(df["Total_Amount"].mean()), 2),
    }
    return df, audit


def build_summaries(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Create category and region summary DataFrames."""
    category_summary = (
        df.groupby("Category", as_index=False)
        .agg(
            Orders=("Order_ID", "count"),
            Total_Units=("Quantity", "sum"),
            Total_Revenue=("Total_Amount", "sum"),
            Avg_Order_Value=("Total_Amount", "mean"),
        )
        .sort_values("Total_Revenue", ascending=False)
    )
    category_summary["Total_Revenue"] = category_summary["Total_Revenue"].round(2)
    category_summary["Avg_Order_Value"] = category_summary["Avg_Order_Value"].round(2)

    region_summary = (
        df.groupby("Region", as_index=False)
        .agg(Orders=("Order_ID", "count"), Total_Revenue=("Total_Amount", "sum"))
        .sort_values("Total_Revenue", ascending=False)
    )
    region_summary["Total_Revenue"] = region_summary["Total_Revenue"].round(2)
    return category_summary, region_summary


def style_excel_workbook(excel_path: Path) -> None:
    """Apply clean header styling and auto-fit column widths."""
    wb = load_workbook(excel_path)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for sheet in wb.worksheets:
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_cells in sheet.columns:
            max_len = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in col_cells
            )
            col_letter = col_cells[0].column_letter
            sheet.column_dimensions[col_letter].width = min(max_len + 4, 32)

    wb.save(excel_path)


def export_excel_report(
    clean_df: pd.DataFrame,
    category_df: pd.DataFrame,
    region_df: pd.DataFrame,
    audit: dict,
    excel_path: Path,
) -> None:
    """Write cleaned data, summaries, and audit log into a multi-sheet Excel file."""
    audit_df = pd.DataFrame(
        [{"Metric": key, "Value": val} for key, val in audit.items()]
    )

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
        category_df.to_excel(writer, sheet_name="Category_Summary", index=False)
        region_df.to_excel(writer, sheet_name="Region_Summary", index=False)
        audit_df.to_excel(writer, sheet_name="Data_Quality_Audit", index=False)

    style_excel_workbook(excel_path)


def export_visual_summary(
    clean_df: pd.DataFrame,
    category_df: pd.DataFrame,
    region_df: pd.DataFrame,
    audit: dict,
    chart_path: Path,
) -> None:
    """Generate a 2x2 visual summary dashboard PNG."""
    fig, axes = plt.subplots(2, 2, figsize=(13, 8))
    fig.suptitle(
        "Automated Data Cleaning & Sales Reporting Summary",
        fontsize=15,
        fontweight="bold",
        y=0.98,
    )

    # 1. Revenue by Category
    ax1 = axes[0, 0]
    ax1.bar(category_df["Category"], category_df["Total_Revenue"], color="#2563EB")
    ax1.set_title("Revenue by Category")
    ax1.set_ylabel("Total Revenue")
    ax1.grid(axis="y", linestyle="--", alpha=0.4)

    # 2. Revenue by Region
    ax2 = axes[0, 1]
    ax2.barh(region_df["Region"], region_df["Total_Revenue"], color="#0D9488")
    ax2.set_title("Revenue by Region")
    ax2.set_xlabel("Total Revenue")
    ax2.invert_yaxis()
    ax2.grid(axis="x", linestyle="--", alpha=0.4)

    # 3. Daily Revenue Trend
    ax3 = axes[1, 0]
    daily = clean_df.groupby("Order_Date", as_index=False)["Total_Amount"].sum()
    ax3.plot(daily["Order_Date"], daily["Total_Amount"], marker="o", color="#7C3AED")
    ax3.set_title("Daily Revenue Trend")
    ax3.set_ylabel("Revenue")
    ax3.tick_params(axis="x", rotation=45)
    ax3.grid(True, linestyle="--", alpha=0.4)

    # 4. Data Quality Before vs After
    ax4 = axes[1, 1]
    labels = ["Raw Rows", "Clean Rows", "Duplicates Removed", "Missing Fixed"]
    values = [
        audit["Initial Rows"],
        audit["Final Clean Rows"],
        audit["Duplicates Removed"],
        audit["Initial Missing Cells"],
    ]
    colors = ["#64748B", "#16A34A", "#DC2626", "#F59E0B"]
    ax4.bar(labels, values, color=colors)
    ax4.set_title("Data Cleaning Audit Summary")
    ax4.tick_params(axis="x", rotation=15)
    ax4.grid(axis="y", linestyle="--", alpha=0.4)

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig(chart_path, dpi=150)
    plt.close(fig)


def export_html_report(
    category_df: pd.DataFrame,
    audit: dict,
    chart_filename: str,
    html_path: Path,
) -> None:
    """Create a clean standalone HTML summary report."""
    audit_rows = "".join(
        f"<tr><td>{html.escape(str(k))}</td><td><strong>{html.escape(str(v))}</strong></td></tr>"
        for k, v in audit.items()
    )
    category_table = category_df.to_html(index=False, classes="table")

    content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>Automated Data Cleaning & Reporting Summary</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 32px; color: #1f2937; background: #f8fafc; }}
    h1, h2 {{ color: #0f172a; }}
    .card {{ background: #ffffff; padding: 20px; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 24px; }}
    table, .table {{ border-collapse: collapse; width: 100%; margin-top: 10px; }}
    th, td {{ border: 1px solid #e2e8f0; padding: 10px 12px; text-align: left; }}
    th {{ background: #1e3a8a; color: #ffffff; }}
    img {{ max-width: 100%; height: auto; border-radius: 8px; }}
  </style>
</head>
<body>
  <h1>Data Cleaning &amp; Reporting Automation</h1>
  <div class="card">
    <h2>1. Data Quality Audit</h2>
    <table>
      <tr><th>Metric</th><th>Value</th></tr>
      {audit_rows}
    </table>
  </div>
  <div class="card">
    <h2>2. Category Performance Summary</h2>
    {category_table}
  </div>
  <div class="card">
    <h2>3. Visual Summary Dashboard</h2>
    <img src="{html.escape(chart_filename)}" alt="Visual Summary Dashboard" />
  </div>
</body>
</html>
"""
    html_path.write_text(content, encoding="utf-8")


def main() -> None:
    """Run the full data cleaning and automated reporting workflow."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    raw_df = load_data(RAW_DATA_PATH)
    clean_df, audit = clean_data(raw_df)
    category_df, region_df = build_summaries(clean_df)

    clean_csv_path = OUTPUT_DIR / "cleaned_sales_data.csv"
    excel_path = OUTPUT_DIR / "automated_sales_report.xlsx"
    chart_path = OUTPUT_DIR / "visual_summary.png"
    html_path = OUTPUT_DIR / "automation_report.html"

    clean_df.to_csv(clean_csv_path, index=False)
    export_excel_report(clean_df, category_df, region_df, audit, excel_path)
    export_visual_summary(clean_df, category_df, region_df, audit, chart_path)
    export_html_report(category_df, audit, chart_path.name, html_path)

    print("Automation complete! Generated files:")
    for path in [clean_csv_path, excel_path, chart_path, html_path]:
        print(f" - {path}")


if __name__ == "__main__":
    main()
