#!/usr/bin/env python3

"""
Generate a lightweight Excel statistics dashboard using only XlsxWriter.

- No external heavy deps (no pandas/numpy)
- Generates a realistic sample dataset
- Creates Data sheet and Dashboard with KPIs and charts

Usage:
  python generate_excel_dashboard.py --output /workspace/statistics_dashboard.xlsx --rows 1000
"""

from __future__ import annotations

import argparse
import os
import random
from datetime import date, datetime, timedelta
from typing import List, Tuple, Dict

import xlsxwriter


def daterange(start_date: date, end_date: date) -> List[date]:
    days = (end_date - start_date).days
    return [start_date + timedelta(days=i) for i in range(days + 1)]


def create_sample_dataset(num_rows: int = 1000, seed: int = 1) -> List[Dict[str, object]]:
    random.seed(seed)

    end_month = date.today().replace(day=1)
    start_month = (end_month - timedelta(days=365)).replace(day=1)
    dates = daterange(start_month, end_month)

    categories = ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot"]
    subcategories = ["S1", "S2", "S3", "S4"]
    regions = ["North", "South", "East", "West"]

    rows: List[Dict[str, object]] = []
    for _ in range(num_rows):
        d = random.choice(dates)
        cat = random.choice(categories)
        sub = random.choice(subcategories)
        reg = random.choice(regions)
        units = random.randint(1, 40)
        price = round(random.uniform(5.0, 150.0), 2)
        value = round(units * price, 2)
        rows.append(
            {
                "Date": d,
                "Category": cat,
                "Subcategory": sub,
                "Region": reg,
                "Units": units,
                "Price": price,
                "Value": value,
            }
        )

    rows.sort(key=lambda r: r["Date"])  # sort by date
    return rows


def compute_kpis(rows: List[Dict[str, object]]) -> Dict[str, float]:
    total_value = sum(r["Value"] for r in rows)
    average_value = round(total_value / len(rows), 2) if rows else 0.0
    total_units = int(sum(r["Units"] for r in rows))
    num_records = len(rows)
    unique_categories = len({r["Category"] for r in rows})
    return {
        "Total Value": round(total_value, 2),
        "Average Value": average_value,
        "Total Units": total_units,
        "Records": num_records,
        "Unique Categories": unique_categories,
    }


def month_label(d: date) -> str:
    return f"{d.year:04d}-{d.month:02d}"


def compute_monthly_summary(rows: List[Dict[str, object]]) -> List[Tuple[str, float]]:
    agg: Dict[str, float] = {}
    for r in rows:
        key = month_label(r["Date"]) if isinstance(r["Date"], date) else str(r["Date"])
        agg[key] = agg.get(key, 0.0) + float(r["Value"])
    return sorted(agg.items(), key=lambda kv: kv[0])


def compute_category_summary(rows: List[Dict[str, object]], top_n: int = 10) -> List[Tuple[str, float]]:
    agg: Dict[str, float] = {}
    for r in rows:
        cat = str(r["Category"])
        agg[cat] = agg.get(cat, 0.0) + float(r["Value"])
    items = sorted(agg.items(), key=lambda kv: kv[1], reverse=True)
    return items[:top_n]


def compute_histogram(rows: List[Dict[str, object]], bins: int = 10) -> List[Tuple[str, int]]:
    values = [float(r["Value"]) for r in rows]
    if not values:
        edges = [i / bins for i in range(bins + 1)]
        counts = [0] * bins
        labels = [f"{edges[i]:.0f}–{edges[i+1]:.0f}" for i in range(bins)]
        return list(zip(labels, counts))

    vmin, vmax = min(values), max(values)
    if vmin == vmax:
        vmin = 0.0
    width = (vmax - vmin) / bins if bins > 0 else 1.0
    edges = [vmin + i * width for i in range(bins + 1)]
    counts = [0 for _ in range(bins)]
    for v in values:
        if v == vmax:
            idx = bins - 1
        else:
            idx = int((v - vmin) / width) if width > 0 else 0
            idx = max(0, min(bins - 1, idx))
        counts[idx] += 1
    labels = [f"{edges[i]:,.0f}–{edges[i+1]:,.0f}" for i in range(bins)]
    return list(zip(labels, counts))


def write_dashboard(rows: List[Dict[str, object]], output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    workbook = xlsxwriter.Workbook(output_path, {
        'in_memory': False,
    })

    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    money_fmt = workbook.add_format({"num_format": "$#,##0.00"})
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})

    data_ws = workbook.add_worksheet("Data")
    headers = ["Date", "Category", "Subcategory", "Region", "Units", "Price", "Value"]
    for c, h in enumerate(headers):
        data_ws.write(0, c, h, header_fmt)

    for r_idx, r in enumerate(rows, start=1):
        data_ws.write_datetime(r_idx, 0, datetime(r["Date"].year, r["Date"].month, r["Date"].day), date_fmt)
        data_ws.write_string(r_idx, 1, str(r["Category"]))
        data_ws.write_string(r_idx, 2, str(r["Subcategory"]))
        data_ws.write_string(r_idx, 3, str(r["Region"]))
        data_ws.write_number(r_idx, 4, int(r["Units"]))
        data_ws.write_number(r_idx, 5, float(r["Price"]), money_fmt)
        data_ws.write_number(r_idx, 6, float(r["Value"]), money_fmt)

    for col in range(len(headers)):
        data_ws.set_column(col, col, 14)
    data_ws.autofilter(0, 0, len(rows), len(headers) - 1)
    data_ws.freeze_panes(1, 0)

    kpis = compute_kpis(rows)
    monthly = compute_monthly_summary(rows)
    by_cat = compute_category_summary(rows)
    hist = compute_histogram(rows)

    dash_ws = workbook.add_worksheet("Dashboard")

    title_fmt = workbook.add_format({"bold": True, "font_size": 18})
    kpi_label_fmt = workbook.add_format({"bold": True, "font_size": 11})
    kpi_value_fmt = workbook.add_format({"bold": True, "font_size": 14})
    subhead_fmt = workbook.add_format({"bold": True, "font_size": 12})
    small_fmt = workbook.add_format({"font_size": 9, "italic": True, "font_color": "#666666"})

    dash_ws.write("A1", "Statistics Dashboard", title_fmt)
    dash_ws.write("A2", f"Refreshed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", small_fmt)

    kpi_items = list(kpis.items())
    for i, (label, value) in enumerate(kpi_items):
        row = 4 + i // 3 * 2
        col = (i % 3) * 4
        dash_ws.write(row, col, label, kpi_label_fmt)
        if isinstance(value, float):
            if "Value" in label:
                dash_ws.write_number(row, col + 1, value, money_fmt)
            else:
                dash_ws.write_number(row, col + 1, value, kpi_value_fmt)
        else:
            dash_ws.write_number(row, col + 1, value, kpi_value_fmt)

    monthly_start_row = 12
    monthly_start_col = 0
    dash_ws.write(monthly_start_row - 2, monthly_start_col, "Monthly Trend (Value)", subhead_fmt)
    dash_ws.write_row(monthly_start_row - 1, monthly_start_col, ["Month", "Value"])
    for i, (m, v) in enumerate(monthly):
        dash_ws.write_string(monthly_start_row + i, monthly_start_col, m)
        dash_ws.write_number(monthly_start_row + i, monthly_start_col + 1, v, money_fmt)

    bycat_start_row = 12
    bycat_start_col = 6
    dash_ws.write(bycat_start_row - 2, bycat_start_col, "Top Categories (Value)", subhead_fmt)
    dash_ws.write_row(bycat_start_row - 1, bycat_start_col, ["Category", "Value"])
    for i, (c, v) in enumerate(by_cat):
        dash_ws.write_string(bycat_start_row + i, bycat_start_col, c)
        dash_ws.write_number(bycat_start_row + i, bycat_start_col + 1, v, money_fmt)

    hist_start_row = 30
    hist_start_col = 0
    dash_ws.write(hist_start_row - 2, hist_start_col, "Value Distribution (Histogram)", subhead_fmt)
    dash_ws.write_row(hist_start_row - 1, hist_start_col, ["Bin", "Count"])
    for i, (label, cnt) in enumerate(hist):
        dash_ws.write_string(hist_start_row + i, hist_start_col, label)
        dash_ws.write_number(hist_start_row + i, hist_start_col + 1, cnt)

    line_chart = workbook.add_chart({"type": "line"})
    line_chart.add_series(
        {
            "name": "Value",
            "categories": [
                "Dashboard",
                monthly_start_row,
                monthly_start_col,
                monthly_start_row + len(monthly) - 1,
                monthly_start_col,
            ],
            "values": [
                "Dashboard",
                monthly_start_row,
                monthly_start_col + 1,
                monthly_start_row + len(monthly) - 1,
                monthly_start_col + 1,
            ],
        }
    )
    line_chart.set_title({"name": "Monthly Trend"})
    line_chart.set_x_axis({"name": "Month"})
    line_chart.set_y_axis({"name": "Value"})
    line_chart.set_legend({"none": True})

    col_chart = workbook.add_chart({"type": "column"})
    col_chart.add_series(
        {
            "name": "Value",
            "categories": [
                "Dashboard",
                bycat_start_row,
                bycat_start_col,
                bycat_start_row + len(by_cat) - 1,
                bycat_start_col,
            ],
            "values": [
                "Dashboard",
                bycat_start_row,
                bycat_start_col + 1,
                bycat_start_row + len(by_cat) - 1,
                bycat_start_col + 1,
            ],
        }
    )
    col_chart.set_title({"name": "Top Categories"})
    col_chart.set_x_axis({"name": "Category"})
    col_chart.set_y_axis({"name": "Value"})
    col_chart.set_legend({"none": True})

    hist_chart = workbook.add_chart({"type": "column"})
    hist_chart.add_series(
        {
            "name": "Count",
            "categories": [
                "Dashboard",
                hist_start_row,
                hist_start_col,
                hist_start_row + len(hist) - 1,
                hist_start_col,
            ],
            "values": [
                "Dashboard",
                hist_start_row,
                hist_start_col + 1,
                hist_start_row + len(hist) - 1,
                hist_start_col + 1,
            ],
        }
    )
    hist_chart.set_title({"name": "Value Distribution"})
    hist_chart.set_x_axis({"name": "Value Bins"})
    hist_chart.set_y_axis({"name": "Count"})
    hist_chart.set_legend({"none": True})

    dash_ws.insert_chart("A20", line_chart, {"x_scale": 1.7, "y_scale": 1.4})
    dash_ws.insert_chart("G20", col_chart, {"x_scale": 1.7, "y_scale": 1.4})
    dash_ws.insert_chart("A38", hist_chart, {"x_scale": 1.7, "y_scale": 1.4})

    data_ws.activate()

    workbook.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate Excel statistics dashboard (lightweight)")
    parser.add_argument(
        "--output",
        type=str,
        default="/workspace/statistics_dashboard.xlsx",
        help="Path to output Excel file.",
    )
    parser.add_argument("--rows", type=int, default=1000, help="Rows for sample data.")
    args = parser.parse_args()

    write_dashboard(create_sample_dataset(args.rows), args.output)
    print(f"Dashboard created: {args.output}")


if __name__ == "__main__":
    main()