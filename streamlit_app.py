from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


EXCEL_FILE = Path(__file__).with_name("B2B_Transaction_Data_.xlsx")
CSV_FILE = Path(__file__).with_name("B2B_Transaction_Data.csv.gz")
COLUMNS = [
    "InvoiceNo", "CustomerID", "InvoiceDate", "StockCode", "Description",
    "Category", "City", "Quantity", "NetPrice", "UnitPrice",
]
TEAL = "#16736e"
ORANGE = "#a85b20"


@st.cache_data(show_spinner="Reading the transaction workbook…")
def load_data(path, modified_at):
    # The timestamp lets Streamlit reload the data when the workbook changes.
    identifiers = {"InvoiceNo": "string", "CustomerID": "string", "StockCode": "string"}
    if path.endswith(".xlsx"):
        source = pd.read_excel(path, sheet_name="Transaction", usecols=COLUMNS,
                               engine="openpyxl", dtype=identifiers)
    else:
        source = pd.read_csv(path, usecols=COLUMNS, dtype=identifiers, low_memory=False)
    df = source.copy()
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], errors="coerce").dt.normalize()
    for name in ["Quantity", "NetPrice", "UnitPrice"]:
        df[name] = pd.to_numeric(df[name], errors="coerce").replace(
            [np.inf, -np.inf], np.nan
        )

    valid = df["InvoiceDate"].notna() & df["Quantity"].notna() & df["NetPrice"].notna()
    df = df.loc[valid].copy()
    for name in ["InvoiceNo", "CustomerID", "StockCode", "Description", "Category", "City"]:
        df[name] = df[name].astype("string").str.strip().replace("", pd.NA)
    df["City"] = df["City"].str.upper()
    df["Category"] = df["Category"].fillna("Unclassified")
    df["City"] = df["City"].fillna("Unknown")
    df["Description"] = df["Description"].fillna("Unnamed product")

    df["Net sales"] = df["Quantity"] * df["NetPrice"]
    df["Units sold"] = df["Quantity"].clip(lower=0)
    df["Units returned"] = -df["Quantity"].clip(upper=0)
    quality = {
        "Source rows": len(source),
        "Rows used": len(df),
        "Excluded rows with missing or invalid dates, quantities or net prices": int((~valid).sum()),
        "Repeated rows in the source": int(source.duplicated().sum()),
        "Rows missing a customer ID": int(df["CustomerID"].isna().sum()),
        "Rows with zero quantity": int(df["Quantity"].eq(0).sum()),
        "Rows with negative net unit price": int(df["NetPrice"].lt(0).sum()),
    }
    return df, quality


def metric_change(current, previous):
    if previous is None or previous <= 0:
        return None
    return f"{(current / previous - 1):+.1%}"


def counts(df):
    purchased = df.loc[df["Quantity"] > 0]
    return {
        "Net sales": df["Net sales"].sum(),
        "Sales invoices": purchased["InvoiceNo"].nunique(),
        "Units sold": df["Units sold"].sum(),
        "Active customers": purchased["CustomerID"].nunique(),
    }


def show_plot(fig, height=335):
    fig.update_layout(
        height=height,
        template="simple_white",
        margin={"l": 8, "r": 18, "t": 12, "b": 8},
        font={"size": 13, "color": "#263845"},
        hoverlabel={"namelength": -1},
    )
    fig.update_xaxes(showline=False, zeroline=False)
    fig.update_yaxes(showline=False, zeroline=False)
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})


def rank_chart(data, label, value, limit=8, value_format=",.0f"):
    top = data.nlargest(limit, value).sort_values(value)
    fig = px.bar(top, x=value, y=label, orientation="h", color_discrete_sequence=[TEAL])
    fig.update_traces(hovertemplate=f"%{{y}}<br>{value} %{{x:{value_format}}}<extra></extra>")
    fig.update_yaxes(title=None, type="category", categoryorder="array",
                     categoryarray=top[label].tolist(), automargin=True)
    fig.update_xaxes(rangemode="tozero", tickformat="~s", showgrid=True,
                     gridcolor="#e8edf1")
    show_plot(fig)


def sales_view(df, start, end):
    chart_period = "W-SUN" if (end - start).days <= 92 else "M"
    chart_label = "Weekly" if chart_period == "W-SUN" else "Monthly"
    periods = pd.period_range(start, end, freq=chart_period)
    trend = df.groupby(df["InvoiceDate"].dt.to_period(chart_period))["Net sales"].sum()
    trend = trend.reindex(periods, fill_value=0)

    trend_date = trend.index.to_timestamp()
    fig = go.Figure(go.Scatter(
        x=trend_date, y=trend, mode="lines+markers",
        line={"color": TEAL, "width": 2.5}, marker={"size": 5},
        hovertemplate="%{x|%d %b %Y}<br>Net sales %{y:,.2f}<extra></extra>",
    ))
    fig.update_xaxes(title=f"{'Week' if chart_period == 'W-SUN' else 'Month'} starting")
    fig.update_yaxes(title="Net sales", tickformat="~s", showgrid=True, gridcolor="#e8edf1")

    trend_area, products_area = st.columns([3, 2])
    with trend_area:
        st.subheader(f"{chart_label} net sales")
        show_plot(fig)
        if start != periods[0].start_time or end != periods[-1].end_time.normalize():
            st.caption("The first or last period may be partial.")
    with products_area:
        st.subheader("Top categories by net sales")
        categories = df.groupby("Category", as_index=False)["Net sales"].sum()
        rank_chart(categories, "Category", "Net sales", value_format=",.2f")

    leading = categories.loc[categories["Net sales"].idxmax()]
    sale_units, return_units = df[["Units sold", "Units returned"]].sum()
    st.write(f'**{leading["Category"]}** had the highest net sales '
             f'({leading["Net sales"]:,.2f}) in the selected period.')
    if sale_units > 0 and return_units > 0:
        st.caption(f"Returns recorded in this period were {return_units:,.0f} units, "
                   f"equivalent to {return_units / sale_units:.1%} of units sold in the period.")


def marketing_view(df, end):
    sales = df.loc[(df["Quantity"] > 0) & df["CustomerID"].notna()]
    if sales.empty:
        st.info("There are no sales with a customer ID in this selection.")
        return

    customers = sales.groupby("CustomerID").agg(
        Invoices=("InvoiceNo", "nunique"),
        Last_purchase=("InvoiceDate", "max"),
    )
    customer_sales = df.groupby("CustomerID")["Net sales"].sum()
    customers["Net sales"] = customer_sales.reindex(customers.index)
    customers["Days since last purchase"] = (end - customers["Last_purchase"]).dt.days
    repeat = customers["Invoices"] >= 2
    inactive = repeat & (customers["Days since last purchase"] >= 60)
    city_count = sales.groupby("City")["CustomerID"].nunique().reset_index(name="Customers")

    cards = st.columns(3)
    cards[0].metric("Repeat customers", f"{repeat.sum():,}", border=True)
    cards[1].metric("Repeat customer share", f"{repeat.mean():.1%}", border=True)
    cards[2].metric("Repeat customers inactive 60+ days", f"{inactive.sum():,}", border=True)

    cities, followup = st.columns([2, 3])
    with cities:
        st.subheader("Active customers by city")
        rank_chart(city_count, "City", "Customers")
    with followup:
        st.subheader("Customers to follow up")
        if inactive.any():
            target = customers.loc[inactive].nlargest(10, "Net sales")
            st.dataframe(target[["Invoices", "Net sales", "Days since last purchase"]],
                         width="stretch", column_config={
                             "Net sales": st.column_config.NumberColumn(format="%.2f"),
                         })
        else:
            st.write("No repeat customer in this selection reached 60 days without a purchase.")
        st.caption("A repeat customer has at least two distinct sales invoices in the "
                   "selected period. Inactivity is measured at the end of that period.")


def completed_series(df, start, end, freq):
    first = start.to_period(freq)
    last = end.to_period(freq)
    if start > first.start_time:
        first += 1
    if end < last.end_time.normalize():
        last -= 1
    periods = pd.period_range(first, last, freq=freq)
    demand = df.groupby(df["InvoiceDate"].dt.to_period(freq))["Units sold"].sum()
    return demand.reindex(periods, fill_value=0)


def project(values, horizon, model, season):
    if model == "Last period":
        return np.repeat(values[-1], horizon)
    if model == "Mean of last 3 periods":
        return np.repeat(values[-3:].mean(), horizon)
    return np.resize(values[-season:], horizon)


def demand_forecast(demand, horizon, season):
    values = demand.to_numpy(dtype=float)
    # Evaluate each model on the same rolling windows using only earlier periods.
    last_origins = range(max(3, len(values) - horizon - 5), len(values) - horizon + 1)
    origins = list(last_origins)
    models = ["Last period", "Mean of last 3 periods"]
    if origins[0] >= season:
        models.append("Same period last year")

    mae = {}
    for model in models:
        errors = [
            np.abs(values[origin:origin + horizon] -
                   project(values[:origin], horizon, model, season))
            for origin in origins
        ]
        mae[model] = np.concatenate(errors).mean()

    best = min(mae, key=mae.get)
    future = pd.period_range(demand.index[-1] + 1, periods=horizon, freq=demand.index.freq)
    forecast = pd.Series(project(values, horizon, best, season), index=future)
    return forecast, best, mae, len(origins)


def forecast_view(df, start, end):
    demand = completed_series(df, start, end, "M")
    horizon, season, interval = 3, 12, "months"
    if len(demand) < 10:
        demand = completed_series(df, start, end, "W-SUN")
        horizon, season, interval = 4, 52, "weeks"

    if len(demand) < 10 or not demand.gt(0).any():
        st.info("Select at least 10 complete weeks with recorded sales to show a forecast.")
        return

    forecast, best, mae, origins = demand_forecast(demand, horizon, season)
    forecast = forecast.round().astype(int)
    history = demand.tail(24 if interval == "months" else 32)
    fig = go.Figure()
    fig.add_scatter(x=history.index.to_timestamp(), y=history, name="Recorded units",
                    mode="lines+markers", line={"color": TEAL, "width": 2.5},
                    hovertemplate="%{x|%d %b %Y}<br>%{y:,.0f} units")
    fig.add_scatter(x=forecast.index.to_timestamp(), y=forecast, name="Forecast",
                    mode="lines+markers", line={"color": ORANGE, "dash": "dash", "width": 2.5},
                    hovertemplate="%{x|%d %b %Y}<br>%{y:,.0f} units")
    fig.update_layout(legend={"orientation": "h", "y": 1.12, "x": 0})
    fig.update_yaxes(title="Units sold", rangemode="tozero", showgrid=True, gridcolor="#e8edf1")
    fig.update_xaxes(title=f"{interval[:-1].title()} starting")
    st.write(f"Forecast for the next **{horizon} {interval}** "
             f"**{forecast.sum():,.0f} units** in total.")
    show_plot(fig, height=355)

    with st.expander("Forecast values and backtest"):
        plan = pd.DataFrame({
            f"{interval[:-1].title()} starting": forecast.index.to_timestamp().date,
            "Forecast units": forecast.to_numpy(),
        })
        st.dataframe(plan, hide_index=True, width="stretch")
        tests = pd.DataFrame(list(mae.items()), columns=["Model", "Mean absolute error (units)"])
        st.dataframe(tests.sort_values("Mean absolute error (units)").round(1),
                     hide_index=True, width="stretch")
        st.caption(f"Selected {best} by forecast error across {origins} rolling cutoffs. "
                   "The history uses complete periods only and fills periods without recorded "
                   "sales with zero. Units sold measure observed demand; stock levels and "
                   "lead times are unavailable.")


def supply_view(df, start, end):
    st.subheader("Demand forecast")
    forecast_view(df, start, end)

    st.subheader("Product priorities by units sold")
    sold = df.loc[(df["Units sold"] > 0) & df["StockCode"].notna()]
    if sold.empty:
        st.info("There are no sales with a stock code in this selection.")
        return

    products = sold.groupby("StockCode").agg(
        Product=("Description", "last"),
        Units=("Units sold", "sum"),
    ).sort_values("Units", ascending=False)
    share = products["Units"] / products["Units"].sum()
    preceding_share = share.cumsum() - share
    products["ABC"] = np.select([preceding_share < 0.80, preceding_share < 0.95],
                               ["A", "B"], default="C")

    warehouse_table = products.head(12).reset_index().rename(columns={"StockCode": "Stock code"})
    left, right = st.columns([2, 3])
    with left:
        st.write("**ABC classes**")
        summary = products.groupby("ABC")["Units"].agg(["size", "sum"]).reset_index()
        summary.columns = ["Class", "Products", "Units"]
        summary["Share (%)"] = (100 * summary["Units"] / products["Units"].sum()).round(1)
        st.dataframe(summary, hide_index=True, width="stretch")
        st.caption("A covers approximately the first 80% of sold units, B the next 15%, "
                   "and C the remainder. Start stock reviews with A products.")
    with right:
        st.write("**Top products**")
        st.dataframe(warehouse_table, hide_index=True, width="stretch")


def main():
    st.set_page_config(page_title="B2B Book Distribution", layout="wide")
    st.title("B2B Book Distribution")
    st.caption("Sales, customer activity and demand planning")

    path = EXCEL_FILE if EXCEL_FILE.is_file() else CSV_FILE
    if not path.is_file():
        st.error(f"Place {EXCEL_FILE.name} in the same folder as streamlit_app.py.")
        st.stop()
    try:
        data, quality = load_data(str(path), path.stat().st_mtime_ns)
    except (ValueError, OSError) as error:
        st.error(f"The transaction workbook could not be read. {error}")
        st.stop()
    if data.empty:
        st.error("The workbook contains no usable transactions.")
        st.stop()

    st.sidebar.header("Filters")
    first, last = data["InvoiceDate"].min(), data["InvoiceDate"].max()
    dates = st.sidebar.date_input("Invoice date", (first.date(), last.date()),
                                  min_value=first.date(), max_value=last.date())
    if len(dates) != 2:
        st.info("Select a start date and an end date.")
        st.stop()
    start, end = map(pd.Timestamp, dates)
    categories = st.sidebar.multiselect("Category", sorted(data["Category"].unique()),
                                       placeholder="All categories")
    cities = st.sidebar.multiselect("City", sorted(data["City"].unique()),
                                   placeholder="All cities")
    matching = data
    if categories:
        matching = matching.loc[matching["Category"].isin(categories)]
    if cities:
        matching = matching.loc[matching["City"].isin(cities)]
    filtered = matching.loc[matching["InvoiceDate"].between(start, end)]
    if filtered.empty:
        st.info("No transactions match the selected filters.")
        st.stop()

    span = end - start + pd.Timedelta(days=1)
    previous_start = start - span
    previous_end = start - pd.Timedelta(days=1)
    previous = matching.loc[matching["InvoiceDate"].between(previous_start, previous_end)]
    comparison = counts(previous) if previous_start >= first else {}

    st.write(f"**{start:%d %b %Y} – {end:%d %b %Y}**")
    for card, (label, value) in zip(st.columns(4), counts(filtered).items()):
        digits = 2 if label == "Net sales" else 0
        card.metric(label, f"{value:,.{digits}f}",
                    delta=metric_change(value, comparison.get(label)), border=True)
    if comparison:
        st.caption("Change versus the preceding date range of equal length.")

    sales, marketing, supply = st.tabs(["Sales", "Marketing", "Supply chain"])
    with sales:
        sales_view(filtered, start, end)
    with marketing:
        marketing_view(filtered, end)
    with supply:
        supply_view(filtered, start, end)

    with st.expander("Data and calculation notes"):
        st.write("Net sales are calculated as quantity × net unit price. Negative quantities "
                 "reduce net sales and are counted separately as returned units. "
                 "The workbook has no currency field, so sales are shown in source units.")
        st.write("Repeated rows are retained because the workbook does not have a unique "
                 "invoice-line identifier. Missing categories and cities are shown separately.")
        st.dataframe(pd.DataFrame(quality.items(), columns=["Check", "Rows"]),
                     hide_index=True, width="stretch")


if __name__ == "__main__":
    main()
