from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill


FOLDER = Path(__file__).resolve().parent
SOURCE = FOLDER / "aqi.xlsx"
OUTPUT = FOLDER / "aqi_cleaned.xlsx"


def main():
    if not SOURCE.exists():
        raise FileNotFoundError(f"Source workbook not found: {SOURCE}")

    df = pd.read_excel(SOURCE, sheet_name=0)
    original_rows = len(df)

    text_columns = df.select_dtypes(include=["object", "string"]).columns
    trimmed_cells = 0
    for column in text_columns:
        values = df[column].astype("string")
        trimmed_cells += int((values != values.str.strip()).sum())
        df[column] = values.str.strip()

    empty_columns = df.columns[df.isna().all()].tolist()
    df = df.drop(columns=empty_columns)
    df = df.drop_duplicates().reset_index(drop=True)
    duplicate_rows_removed = original_rows - len(df)

    df["date"] = pd.to_datetime(df["date"], errors="raise")
    for column in ["number_of_monitoring_stations", "aqi_value"]:
        df[column] = pd.to_numeric(df[column], errors="raise")

    bins = [0, 50, 100, 200, 300, 400, float("inf")]
    labels = ["Good", "Satisfactory", "Moderate", "Poor", "Very Poor", "Severe"]
    expected_status = pd.cut(
        df["aqi_value"], bins=bins, labels=labels, right=True
    ).astype("string")
    mismatch = df["air_quality_status"] != expected_status
    corrected_rows = int(mismatch.sum())
    df.loc[mismatch, "air_quality_status"] = expected_status[mismatch]

    if df[["date", "aqi_value", "number_of_monitoring_stations"]].isna().any().any():
        raise ValueError("Missing values found in required columns.")
    if not df["aqi_value"].between(0, 500).all():
        raise ValueError("AQI values outside the expected 0-500 range.")
    if not (df["number_of_monitoring_stations"] > 0).all():
        raise ValueError("Non-positive monitoring station count found.")
    if (df["air_quality_status"] != expected_status).any():
        raise ValueError("AQI category labels remain inconsistent.")

    summary = pd.DataFrame(
        [
            ("Source file", SOURCE.name),
            ("Cleaned file", OUTPUT.name),
            ("Rows in source", original_rows),
            ("Rows in cleaned dataset", len(df)),
            ("Duplicate rows removed", duplicate_rows_removed),
            (
                "Entirely empty columns removed",
                ", ".join(empty_columns) if empty_columns else "None",
            ),
            ("Whitespace-only text normalization", f"{trimmed_cells} cells trimmed"),
            ("AQI status labels corrected", corrected_rows),
            ("Missing values remaining", int(df.isna().sum().sum())),
            (
                "AQI range check",
                f"Passed ({df['aqi_value'].min()}-{df['aqi_value'].max()})",
            ),
            ("Station count check", "Passed (all values > 0)"),
            ("AQI/status consistency check", "Passed"),
        ],
        columns=["Cleaning step / check", "Result"],
    )

    with pd.ExcelWriter(
        OUTPUT,
        engine="openpyxl",
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        summary.to_excel(writer, sheet_name="Cleaning Summary", index=False)

        data_ws = writer.book["Cleaned Data"]
        data_ws.freeze_panes = "A2"
        data_ws.auto_filter.ref = data_ws.dimensions
        data_ws.row_dimensions[1].height = 30
        for cell in data_ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        widths = {
            "A": 13,
            "B": 24,
            "C": 22,
            "D": 30,
            "E": 26,
            "F": 12,
            "G": 22,
            "H": 85,
        }
        for column, width in widths.items():
            if column in data_ws.column_dimensions:
                data_ws.column_dimensions[column].width = width

        summary_ws = writer.book["Cleaning Summary"]
        summary_ws.freeze_panes = "A2"
        summary_ws.column_dimensions["A"].width = 38
        summary_ws.column_dimensions["B"].width = 105
        for cell in summary_ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for row in summary_ws.iter_rows(min_row=2, min_col=2, max_col=2):
            row[0].alignment = Alignment(wrap_text=True, vertical="top")

    saved_data = pd.read_excel(OUTPUT, sheet_name="Cleaned Data")
    if saved_data.shape != df.shape:
        raise ValueError(
            f"Output verification failed: expected {df.shape}, got {saved_data.shape}."
        )
    print(f"Cleaned workbook created: {OUTPUT}")
    print(
        f"Rows: {len(saved_data):,}; duplicate rows removed: "
        f"{duplicate_rows_removed:,}; status labels corrected: {corrected_rows:,}"
    )


if __name__ == "__main__":
    main()
