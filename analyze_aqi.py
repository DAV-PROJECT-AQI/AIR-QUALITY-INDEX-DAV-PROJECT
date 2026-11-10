from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill


FOLDER = Path(__file__).resolve().parent
CLEANED_FILE = FOLDER / "aqi_cleaned.xlsx"
SOURCE_FILE = FOLDER / "aqi.xlsx"
OUTPUT_FILE = FOLDER / "aqi_analytics.xlsx"

AQI_BINS = [0, 50, 100, 200, 300, 400, float("inf")]
AQI_LABELS = ["Good", "Satisfactory", "Moderate", "Poor", "Very Poor", "Severe"]
SEASON_BY_MONTH = {
    12: "Winter",
    1: "Winter",
    2: "Winter",
    3: "Summer",
    4: "Summer",
    5: "Summer",
    6: "Monsoon",
    7: "Monsoon",
    8: "Monsoon",
    9: "Monsoon",
    10: "Post-monsoon",
    11: "Post-monsoon",
}
SEASON_ORDER = ["Winter", "Summer", "Monsoon", "Post-monsoon"]


def load_data():
    if CLEANED_FILE.exists():
        try:
            return pd.read_excel(CLEANED_FILE, sheet_name="Cleaned Data"), CLEANED_FILE
        except PermissionError:
            if not SOURCE_FILE.exists():
                raise
            print(
                f"Warning: {CLEANED_FILE.name} is locked or inaccessible; "
                f"using {SOURCE_FILE.name} instead."
            )
    if SOURCE_FILE.exists():
        return pd.read_excel(SOURCE_FILE, sheet_name=0), SOURCE_FILE
    raise FileNotFoundError(
        f"Could not find {CLEANED_FILE.name} or {SOURCE_FILE.name} in {FOLDER}"
    )


def make_analysis(df):
    required = {
        "date",
        "state",
        "area",
        "aqi_value",
        "air_quality_status",
        "prominent_pollutants",
    }
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")

    df = df.copy()
    source_rows = len(df)
    missing_required_values = int(df[list(required)].isna().sum().sum())
    empty_columns = df.columns[df.isna().all()].tolist()
    df["date"] = pd.to_datetime(df["date"], errors="raise")
    df["aqi_value"] = pd.to_numeric(df["aqi_value"], errors="raise")
    df = df.dropna(subset=["date", "state", "area", "aqi_value"])
    excluded_rows = source_rows - len(df)
    if df.empty:
        raise ValueError("No complete AQI observations are available for analysis.")
    if not df["aqi_value"].between(0, 500).all():
        raise ValueError("AQI values outside the expected 0-500 range.")

    expected_status = pd.cut(
        df["aqi_value"],
        bins=AQI_BINS,
        labels=AQI_LABELS,
        right=True,
        include_lowest=True,
    )
    if expected_status.isna().any():
        raise ValueError("Could not assign an AQI category to all observations.")
    status_mismatches = int(
        (df["air_quality_status"].astype("string") != expected_status.astype("string")).sum()
    )
    df["air_quality_status"] = expected_status.astype("string")
    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.month
    df["month_start"] = df["date"].dt.to_period("M").dt.to_timestamp()
    df["weekday"] = df["date"].dt.day_name()
    df["season"] = df["month"].map(SEASON_BY_MONTH)

    overview = pd.DataFrame(
        [
            ("Source workbook", ""),
            ("Complete AQI observations", len(df)),
            ("Date coverage", f"{df['date'].min():%Y-%m-%d} to {df['date'].max():%Y-%m-%d}"),
            ("States / territories", df["state"].nunique()),
            ("Cities / areas", df["area"].nunique()),
            ("Average AQI", round(df["aqi_value"].mean(), 2)),
            ("Median AQI", round(df["aqi_value"].median(), 2)),
            ("Minimum AQI", int(df["aqi_value"].min())),
            ("Maximum AQI", int(df["aqi_value"].max())),
            ("AQI standard deviation", round(df["aqi_value"].std(), 2)),
            ("AQI status labels aligned to AQI bands", status_mismatches),
            ("AQI category bands", "Good 0-50; Satisfactory 51-100; Moderate 101-200; Poor 201-300; Very Poor 301-400; Severe 401-500"),
            ("Season definitions", "Winter: Dec-Feb; Summer: Mar-May; Monsoon: Jun-Sep; Post-monsoon: Oct-Nov"),
            ("City ranking minimum sample", "30 observations per city"),
            ("Year coverage note", "First and last years may be partial."),
        ],
        columns=["Metric", "Value"],
    )

    monthly = (
        df.groupby("month_start")["aqi_value"]
        .agg(observations="size", average_aqi="mean", median_aqi="median", maximum_aqi="max")
        .reset_index()
        .rename(columns={"month_start": "month"})
    )
    for column in ["average_aqi", "median_aqi"]:
        monthly[column] = monthly[column].round(2)

    yearly = (
        df.groupby("year")["aqi_value"]
        .agg(observations="size", average_aqi="mean", median_aqi="median", maximum_aqi="max")
        .reset_index()
    )
    yearly[["average_aqi", "median_aqi"]] = yearly[
        ["average_aqi", "median_aqi"]
    ].round(2)

    state_summary = (
        df.groupby("state")
        .agg(
            observations=("aqi_value", "size"),
            average_aqi=("aqi_value", "mean"),
            median_aqi=("aqi_value", "median"),
            maximum_aqi=("aqi_value", "max"),
            cities=("area", "nunique"),
        )
        .reset_index()
        .sort_values("average_aqi", ascending=False)
    )
    state_summary[["average_aqi", "median_aqi"]] = state_summary[
        ["average_aqi", "median_aqi"]
    ].round(2)

    city_summary = (
        df.groupby(["state", "area"])
        .agg(
            observations=("aqi_value", "size"),
            average_aqi=("aqi_value", "mean"),
            median_aqi=("aqi_value", "median"),
            maximum_aqi=("aqi_value", "max"),
        )
        .reset_index()
    )
    city_summary = city_summary[city_summary["observations"] >= 30].copy()
    city_summary[["average_aqi", "median_aqi"]] = city_summary[
        ["average_aqi", "median_aqi"]
    ].round(2)
    city_summary = city_summary.sort_values("average_aqi", ascending=False)

    category_summary = (
        df["air_quality_status"]
        .value_counts()
        .reindex(AQI_LABELS, fill_value=0)
        .rename_axis("category")
        .reset_index(name="observations")
    )
    category_summary["share_percent"] = (
        category_summary["observations"] / len(df) * 100
    ).round(2)

    pollutant_summary = (
        df["prominent_pollutants"]
        .dropna()
        .astype("string")
        .str.split(",")
        .explode()
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .value_counts()
        .rename_axis("pollutant")
        .reset_index(name="observations")
    )
    pollutant_summary["share_of_pollutant_mentions_percent"] = (
        pollutant_summary["observations"] / pollutant_summary["observations"].sum() * 100
    ).round(2)

    season_summary = (
        df.groupby("season", observed=False)["aqi_value"]
        .agg(observations="size", average_aqi="mean", median_aqi="median", maximum_aqi="max")
        .reindex(SEASON_ORDER)
        .reset_index()
    )
    season_summary[["average_aqi", "median_aqi"]] = season_summary[
        ["average_aqi", "median_aqi"]
    ].round(2)

    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    weekday_summary = (
        df.groupby("weekday")["aqi_value"]
        .agg(observations="size", average_aqi="mean", median_aqi="median")
        .reindex(weekday_order)
        .reset_index()
    )
    weekday_summary[["average_aqi", "median_aqi"]] = weekday_summary[
        ["average_aqi", "median_aqi"]
    ].round(2)

    state_categories = pd.crosstab(
        df["state"], df["air_quality_status"], normalize="index"
    ).reindex(columns=AQI_LABELS, fill_value=0)
    state_categories.columns = [f"{category} share" for category in state_categories.columns]
    state_categories = state_categories.reset_index()
    share_columns = [column for column in state_categories.columns if column.endswith(" share")]
    state_categories[share_columns] = (state_categories[share_columns] * 100).round(2)

    top_pollutants = ", ".join(
        f"{row.pollutant} ({row.observations:,})"
        for row in pollutant_summary.head(5).itertuples(index=False)
    )
    worst_state = state_summary.iloc[0]
    best_state = state_summary.iloc[-1]
    worst_month = monthly.loc[monthly["average_aqi"].idxmax()]
    best_month = monthly.loc[monthly["average_aqi"].idxmin()]
    most_common_category = category_summary.loc[
        category_summary["observations"].idxmax()
    ]
    insights = pd.DataFrame(
        [
            ("Highest-average-AQI state", f"{worst_state['state']}: {worst_state['average_aqi']:.2f} across {int(worst_state['observations']):,} observations"),
            ("Lowest-average-AQI state", f"{best_state['state']}: {best_state['average_aqi']:.2f} across {int(best_state['observations']):,} observations"),
            ("Highest-average-AQI month", f"{worst_month['month']:%B %Y}: {worst_month['average_aqi']:.2f}"),
            ("Lowest-average-AQI month", f"{best_month['month']:%B %Y}: {best_month['average_aqi']:.2f}"),
            ("Most common AQI category", f"{most_common_category['category']}: {int(most_common_category['observations']):,} observations ({most_common_category['share_percent']:.2f}%)"),
            ("Most-mentioned pollutants", top_pollutants),
            ("Highest city average (30+ observations)", f"{city_summary.iloc[0]['area']}, {city_summary.iloc[0]['state']}: {city_summary.iloc[0]['average_aqi']:.2f}" if not city_summary.empty else "No city has 30 or more observations."),
            ("Interpretation note", "These are descriptive summaries of the supplied monitoring records; they do not establish causes or population-wide exposure."),
        ],
        columns=["Finding", "Result"],
    )

    quality = pd.DataFrame(
        [
            ("Rows in source workbook", source_rows),
            ("Rows excluded for missing required fields", excluded_rows),
            ("Missing required-field cells in source", missing_required_values),
            ("Entirely empty source columns", ", ".join(empty_columns) if empty_columns else "None"),
            ("Missing cells in analyzed rows (including optional fields)", int(df.isna().sum().sum())),
            ("Exact duplicate rows in analyzed data", int(df.duplicated().sum())),
            ("AQI values outside 0-500", int((~df["aqi_value"].between(0, 500)).sum())),
            ("AQI/status mismatches corrected for analysis", status_mismatches),
            ("City averages with at least 30 records", len(city_summary)),
        ],
        columns=["Check", "Result"],
    )

    return {
        "Key Insights": insights,
        "Overview": overview,
        "Monthly Trends": monthly,
        "Yearly Summary": yearly,
        "State Analysis": state_summary,
        "City Analysis": city_summary,
        "AQI Categories": category_summary,
        "Pollutant Frequency": pollutant_summary,
        "Seasonal Analysis": season_summary,
        "Weekday Analysis": weekday_summary,
        "State Category Shares": state_categories,
        "Data Quality": quality,
    }, status_mismatches


def write_workbook(sheets):
    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        for sheet_name, data in sheets.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.row_dimensions[1].height = 30
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            for column_cells in ws.columns:
                letter = column_cells[0].column_letter
                max_length = max(
                    (len(str(cell.value)) for cell in column_cells if cell.value is not None),
                    default=10,
                )
                ws.column_dimensions[letter].width = min(max(max_length + 2, 12), 55)
            if sheet_name in {"Key Insights", "Overview"}:
                for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
                    row[0].alignment = Alignment(wrap_text=True, vertical="top")

def main():
    df, data_file = load_data()
    sheets, mismatches = make_analysis(df)
    sheets["Overview"].loc[
        sheets["Overview"]["Metric"] == "Source workbook", "Value"
    ] = data_file.name
    write_workbook(sheets)

    check = pd.ExcelFile(OUTPUT_FILE)
    expected_sheets = list(sheets)
    if check.sheet_names != expected_sheets:
        raise OSError(f"Analytics workbook sheets did not verify: {check.sheet_names}")
    check_overview = pd.read_excel(OUTPUT_FILE, sheet_name="Overview")
    if int(check_overview.loc[check_overview["Metric"] == "Complete AQI observations", "Value"].iloc[0]) != len(df.dropna(subset=["date", "state", "area", "aqi_value"])):
        raise OSError("Analytics workbook row-count verification failed.")

    print(f"Used data: {data_file}")
    print(f"Created analytics workbook: {OUTPUT_FILE}")
    print(f"Analyzed observations: {len(df):,}; AQI/status mismatches aligned: {mismatches:,}")
    for row in sheets["Key Insights"].itertuples(index=False):
        print(f"{row.Finding}: {row.Result}")


if __name__ == "__main__":
    main()
