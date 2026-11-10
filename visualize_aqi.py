from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots


FOLDER = Path(__file__).resolve().parent
CLEANED_FILE = FOLDER / "aqi_cleaned.xlsx"
SOURCE_FILE = FOLDER / "aqi.xlsx"
OUTPUT_FOLDER = FOLDER / "aqi_graphs"

AQI_COLORS = {
    "Good": "#2E8B57",
    "Satisfactory": "#9ACD32",
    "Moderate": "#FFD700",
    "Poor": "#FF8C00",
    "Very Poor": "#DC143C",
    "Severe": "#8B0000",
}
STATUS_ORDER = ["Good", "Satisfactory", "Moderate", "Poor", "Very Poor", "Severe"]


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


def save_monthly_trend(path, monthly):
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(monthly.index, monthly.values, color="#156082", linewidth=2)
    ax.fill_between(monthly.index, monthly.values, color="#156082", alpha=0.12)
    ax.set(
        title="Monthly Average Air Quality Index",
        xlabel="Month",
        ylabel="Average AQI",
    )
    ax.grid(axis="y", alpha=0.25)
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_state_averages(path, state_averages):
    ordered = state_averages.sort_values("average_aqi")
    fig, ax = plt.subplots(figsize=(11, 8))
    ax.barh(
        ordered.index,
        ordered["average_aqi"],
        color="#4C78A8",
        edgecolor="white",
    )
    ax.set(
        title="15 States with Highest Average AQI",
        xlabel="Average AQI",
        ylabel="State",
    )
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_category_distribution(path, category_counts):
    counts = category_counts.reindex(STATUS_ORDER, fill_value=0)
    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.bar(
        counts.index,
        counts.values,
        color=[AQI_COLORS[status] for status in counts.index],
    )
    ax.bar_label(bars, labels=[f"{value:,}" for value in counts.values], padding=3)
    ax.set(
        title="AQI Observations by Air Quality Category",
        xlabel="Air quality category",
        ylabel="Number of observations",
    )
    ax.tick_params(axis="x", rotation=20)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_pollutants(path, pollutant_counts):
    ordered = pollutant_counts.sort_values()
    fig, ax = plt.subplots(figsize=(9, 6))
    ax.barh(ordered.index, ordered.values, color="#59A14F")
    ax.set(
        title="Most Frequently Reported Prominent Pollutants",
        xlabel="Number of observations",
        ylabel="Pollutant",
    )
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_aqi_histogram(path, df):
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.hist(df["aqi_value"], bins=35, color="#4C78A8", edgecolor="white")
    ax.set(
        title="Distribution of AQI Values",
        xlabel="AQI",
        ylabel="Number of observations",
    )
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_top_areas(path, df):
    locations = (
        df.assign(location=df["area"].astype("string") + ", " + df["state"].astype("string"))
        .groupby("location")["aqi_value"]
        .agg(average_aqi="mean", observations="size")
        .query("observations >= 30")
        .nlargest(15, "average_aqi")
        .sort_values("average_aqi")
    )
    if locations.empty:
        raise ValueError("Not enough city observations to create the city AQI chart.")
    fig, ax = plt.subplots(figsize=(11, 8))
    ax.barh(locations.index, locations["average_aqi"], color="#E15759")
    ax.set(
        title="15 Cities with Highest Average AQI (at least 30 observations)",
        xlabel="Average AQI",
        ylabel="City, state",
    )
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_category_pie(path, category_counts):
    counts = category_counts.reindex(STATUS_ORDER, fill_value=0)
    counts = counts[counts > 0]
    fig, ax = plt.subplots(figsize=(9, 7))
    ax.pie(
        counts.values,
        labels=counts.index,
        colors=[AQI_COLORS[status] for status in counts.index],
        autopct="%1.1f%%",
        startangle=90,
        wedgeprops={"edgecolor": "white", "linewidth": 1},
    )
    ax.set_title("Share of Observations by AQI Category")
    ax.axis("equal")
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_pollutant_pie(path, pollutant_counts):
    # Group less common pollutants together to keep the pie readable.
    counts = pollutant_counts.copy()
    other_count = int(counts.iloc[6:].sum())
    counts = counts.iloc[:6]
    if other_count:
        counts.loc["Other pollutants"] = other_count
    fig, ax = plt.subplots(figsize=(9, 7))
    ax.pie(
        counts.values,
        labels=counts.index,
        autopct="%1.1f%%",
        startangle=90,
        wedgeprops={"edgecolor": "white", "linewidth": 1},
    )
    ax.set_title("Share of Reported Prominent Pollutants")
    ax.axis("equal")
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_weekday_aqi(path, df):
    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    weekday_average = (
        df.assign(weekday=df["date"].dt.day_name())
        .groupby("weekday")["aqi_value"]
        .mean()
        .reindex(weekday_order)
    )
    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.bar(
        weekday_average.index,
        weekday_average.values,
        color="#F28E2B",
    )
    ax.bar_label(bars, fmt="%.1f", padding=3)
    ax.set(
        title="Average AQI by Day of Week",
        xlabel="Day of week",
        ylabel="Average AQI",
    )
    ax.tick_params(axis="x", rotation=20)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_year_month_heatmap(path, df):
    dated = df.assign(year=df["date"].dt.year, month=df["date"].dt.month)
    heatmap = dated.pivot_table(
        index="year",
        columns="month",
        values="aqi_value",
        aggfunc="mean",
        observed=False,
    ).reindex(columns=range(1, 13))
    fig, ax = plt.subplots(figsize=(12, 6))
    image = ax.imshow(heatmap, aspect="auto", cmap="YlOrRd", interpolation="nearest")
    ax.set_xticks(range(12), ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    ax.set_yticks(range(len(heatmap.index)), heatmap.index.astype(str))
    ax.set(
        title="Average AQI by Year and Month",
        xlabel="Month",
        ylabel="Year",
    )
    fig.colorbar(image, ax=ax, label="Average AQI")
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_station_group_aqi(path, df):
    if "number_of_monitoring_stations" not in df.columns:
        raise ValueError("The workbook is missing number_of_monitoring_stations.")
    station_bins = pd.cut(
        pd.to_numeric(df["number_of_monitoring_stations"], errors="raise"),
        bins=[0, 1, 2, 5, 10, float("inf")],
        labels=["1", "2", "3-5", "6-10", "11+"],
        include_lowest=True,
    )
    grouped = (
        df.assign(station_group=station_bins)
        .groupby("station_group", observed=False)["aqi_value"]
        .agg(["mean", "size"])
    )
    grouped = grouped[grouped["size"] > 0]
    fig, ax = plt.subplots(figsize=(9, 6))
    bars = ax.bar(grouped.index.astype(str), grouped["mean"], color="#76B7B2")
    ax.bar_label(
        bars,
        labels=[f"{value:.1f}\n(n={count:,})" for value, count in zip(grouped["mean"], grouped["size"])],
        padding=3,
    )
    ax.set(
        title="Average AQI by Number of Monitoring Stations",
        xlabel="Stations reporting for the observation",
        ylabel="Average AQI",
    )
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_state_category_mix(path, df):
    state_means = df.groupby("state")["aqi_value"].mean()
    top_states = state_means.nlargest(10).index
    mix = pd.crosstab(
        df.loc[df["state"].isin(top_states), "state"],
        df.loc[df["state"].isin(top_states), "air_quality_status"],
        normalize="index",
    ).reindex(columns=STATUS_ORDER, fill_value=0)
    mix = mix.loc[state_means.loc[top_states].sort_values(ascending=False).index]
    fig, ax = plt.subplots(figsize=(12, 7))
    bottom = pd.Series(0.0, index=mix.index)
    for status in STATUS_ORDER:
        values = mix[status]
        ax.bar(
            mix.index,
            values,
            bottom=bottom,
            label=status,
            color=AQI_COLORS[status],
        )
        bottom += values
    ax.set(
        title="AQI Category Mix in 10 Highest-AQI States",
        xlabel="State",
        ylabel="Share of observations",
        ylim=(0, 1),
    )
    ax.tick_params(axis="x", rotation=30)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda value, _: f"{value:.0%}"))
    ax.legend(title="AQI category", bbox_to_anchor=(1.02, 1), loc="upper left")
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_interactive_dashboard(path, monthly, state_averages, category_counts, pollutant_counts):
    fig = make_subplots(
        rows=2,
        cols=2,
        subplot_titles=(
            "Monthly Average AQI",
            "15 States with Highest Average AQI",
            "AQI Observations by Category",
            "Most Frequently Reported Pollutants",
        ),
        vertical_spacing=0.16,
        horizontal_spacing=0.12,
    )

    fig.add_trace(
        go.Scatter(
            x=monthly.index,
            y=monthly.values,
            mode="lines+markers",
            name="Monthly average",
            line={"color": "#156082"},
            hovertemplate="%{x|%b %Y}<br>Average AQI: %{y:.1f}<extra></extra>",
        ),
        row=1,
        col=1,
    )

    states = state_averages.sort_values("average_aqi", ascending=True)
    fig.add_trace(
        go.Bar(
            x=states["average_aqi"],
            y=states.index,
            orientation="h",
            name="State average",
            marker_color="#4C78A8",
            customdata=states["observations"],
            hovertemplate="%{y}<br>Average AQI: %{x:.1f}<br>Observations: %{customdata:,}<extra></extra>",
        ),
        row=1,
        col=2,
    )

    counts = category_counts.reindex(STATUS_ORDER, fill_value=0)
    fig.add_trace(
        go.Bar(
            x=counts.index,
            y=counts.values,
            name="AQI category",
            marker_color=[AQI_COLORS[status] for status in counts.index],
            hovertemplate="%{x}<br>Observations: %{y:,}<extra></extra>",
        ),
        row=2,
        col=1,
    )

    pollutants = pollutant_counts.sort_values(ascending=True)
    fig.add_trace(
        go.Bar(
            x=pollutants.values,
            y=pollutants.index,
            orientation="h",
            name="Pollutant frequency",
            marker_color="#59A14F",
            hovertemplate="%{y}<br>Observations: %{x:,}<extra></extra>",
        ),
        row=2,
        col=2,
    )

    fig.update_xaxes(title_text="Month", row=1, col=1)
    fig.update_yaxes(title_text="Average AQI", row=1, col=1)
    fig.update_xaxes(title_text="Average AQI", row=1, col=2)
    fig.update_xaxes(title_text="Category", row=2, col=1)
    fig.update_yaxes(title_text="Observations", row=2, col=1)
    fig.update_xaxes(title_text="Observations", row=2, col=2)
    fig.update_layout(
        title="Air Quality in India: AQI Dashboard",
        template="plotly_white",
        height=900,
        width=1450,
        showlegend=False,
        margin={"t": 100, "l": 80, "r": 40, "b": 70},
    )
    fig.write_html(path, include_plotlyjs=True, full_html=True)


def main():
    df, data_file = load_data()
    required = {
        "date",
        "state",
        "area",
        "aqi_value",
        "air_quality_status",
        "prominent_pollutants",
    }
    missing_columns = required.difference(df.columns)
    if missing_columns:
        raise ValueError(f"Missing required data columns: {sorted(missing_columns)}")

    df["date"] = pd.to_datetime(df["date"], errors="raise")
    df["aqi_value"] = pd.to_numeric(df["aqi_value"], errors="raise")
    df = df.dropna(subset=["date", "aqi_value", "state"])
    if df.empty:
        raise ValueError("No complete AQI observations are available to visualize.")
    if not df["aqi_value"].between(0, 500).all():
        raise ValueError("AQI values outside the expected 0-500 range.")
    bins = [0, 50, 100, 200, 300, 400, float("inf")]
    expected_status = pd.cut(
        df["aqi_value"],
        bins=bins,
        labels=STATUS_ORDER,
        right=True,
        include_lowest=True,
    )
    if expected_status.isna().any():
        raise ValueError("Could not assign an AQI category to every observation.")
    corrected_statuses = int(
        (df["air_quality_status"].astype("string") != expected_status.astype("string")).sum()
    )
    df["air_quality_status"] = expected_status.astype("string")

    monthly = df.set_index("date")["aqi_value"].resample("MS").mean().dropna()
    state_averages = (
        df.groupby("state")["aqi_value"]
        .agg(average_aqi="mean", observations="size")
        .nlargest(15, "average_aqi")
    )
    category_counts = df["air_quality_status"].value_counts()
    all_pollutant_counts = (
        df["prominent_pollutants"]
        .dropna()
        .astype("string")
        .str.split(",")
        .explode()
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .value_counts()
    )
    pollutant_counts = all_pollutant_counts.head(10)
    if monthly.empty or state_averages.empty or pollutant_counts.empty:
        raise ValueError("The workbook does not contain enough data to create all charts.")

    OUTPUT_FOLDER.mkdir(exist_ok=True)
    outputs = {
        "monthly": OUTPUT_FOLDER / "aqi_monthly_trend.png",
        "states": OUTPUT_FOLDER / "aqi_state_averages.png",
        "categories": OUTPUT_FOLDER / "aqi_category_distribution.png",
        "pollutants": OUTPUT_FOLDER / "aqi_pollutants.png",
        "aqi_histogram": OUTPUT_FOLDER / "aqi_value_distribution.png",
        "top_areas": OUTPUT_FOLDER / "aqi_city_averages.png",
        "category_pie": OUTPUT_FOLDER / "aqi_category_pie.png",
        "pollutant_pie": OUTPUT_FOLDER / "aqi_pollutant_pie.png",
        "weekdays": OUTPUT_FOLDER / "aqi_weekday_averages.png",
        "year_month": OUTPUT_FOLDER / "aqi_year_month_heatmap.png",
        "station_groups": OUTPUT_FOLDER / "aqi_by_station_count.png",
        "state_category_mix": OUTPUT_FOLDER / "aqi_state_category_mix.png",
        "dashboard": OUTPUT_FOLDER / "aqi_interactive_dashboard.html",
    }

    save_monthly_trend(outputs["monthly"], monthly)
    save_state_averages(outputs["states"], state_averages)
    save_category_distribution(outputs["categories"], category_counts)
    save_pollutants(outputs["pollutants"], pollutant_counts)
    save_aqi_histogram(outputs["aqi_histogram"], df)
    save_top_areas(outputs["top_areas"], df)
    save_category_pie(outputs["category_pie"], category_counts)
    save_pollutant_pie(outputs["pollutant_pie"], all_pollutant_counts)
    save_weekday_aqi(outputs["weekdays"], df)
    save_year_month_heatmap(outputs["year_month"], df)
    save_station_group_aqi(outputs["station_groups"], df)
    save_state_category_mix(outputs["state_category_mix"], df)
    save_interactive_dashboard(
        outputs["dashboard"],
        monthly,
        state_averages,
        category_counts,
        pollutant_counts,
    )

    for path in outputs.values():
        if not path.is_file() or path.stat().st_size == 0:
            raise OSError(f"Chart output was not created correctly: {path}")
    print(f"Used data: {data_file}")
    print(f"AQI category labels aligned with AQI values: {corrected_statuses:,}")
    print(f"Created graph files in: {OUTPUT_FOLDER}")
    for path in outputs.values():
        print(f"  {path.name} ({path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
