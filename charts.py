"""
charts.py
=========
Interactive Visualization and Native Excel Chart Embedding Module.
Uses Plotly for rich browser dashboards and openpyxl for embedding real native charts in Excel.
"""

from __future__ import annotations
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
import numpy as np

try:
    import plotly.express as px
    import plotly.graph_objects as go
except ImportError:
    px = None
    go = None

try:
    import openpyxl
    from openpyxl.chart import BarChart, PieChart, LineChart, Reference
except ImportError:
    openpyxl = None
    BarChart = None
    PieChart = None
    LineChart = None
    Reference = None


def generate_plotly_chart(
    df: pd.DataFrame,
    chart_type: str = "bar",
    x_col: Optional[str] = None,
    y_col: Optional[str] = None,
    color_col: Optional[str] = None,
    title: Optional[str] = None
) -> Optional[Any]:
    """
    Creates an interactive Plotly chart figure.
    chart_type: 'bar', 'line', 'pie', 'scatter', 'histogram'
    """
    if px is None:
        return None

    c_type = chart_type.lower()
    cols = list(df.columns)
    if not cols:
        return None

    # Detect categorical and numerical columns if not provided
    num_cols = list(df.select_dtypes(include=[np.number]).columns)
    cat_cols = [c for c in cols if c not in num_cols]

    actual_x = x_col or (cat_cols[0] if cat_cols else cols[0])
    actual_y = y_col or (num_cols[0] if num_cols else (cols[1] if len(cols) > 1 else cols[0]))

    chart_title = title or f"{actual_y} by {actual_x}"

    try:
        if c_type == "bar":
            fig = px.bar(
                df, x=actual_x, y=actual_y, color=color_col,
                title=chart_title, template="plotly_white"
            )
        elif c_type == "line":
            fig = px.line(
                df, x=actual_x, y=actual_y, color=color_col,
                title=chart_title, template="plotly_white"
            )
        elif c_type == "pie":
            fig = px.pie(
                df, names=actual_x, values=actual_y,
                title=chart_title, template="plotly_white"
            )
        elif c_type == "scatter":
            fig = px.scatter(
                df, x=actual_x, y=actual_y, color=color_col,
                title=chart_title, template="plotly_white"
            )
        elif c_type == "histogram":
            fig = px.histogram(
                df, x=actual_y,
                title=chart_title, template="plotly_white"
            )
        else:
            fig = px.bar(df, x=actual_x, y=actual_y, title=chart_title)

        fig.update_layout(
            margin=dict(l=20, r=20, t=40, b=20),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Plus Jakarta Sans, sans-serif")
        )
        return fig
    except Exception as e:
        print("Chart generation error:", e)
        return None


def generate_quick_dashboard(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Automatically creates 3-4 useful charts and KPIs for any dataset.
    """
    kpis: Dict[str, Any] = {}
    charts: List[Any] = []

    total_rows = len(df)
    kpis["Total Records"] = f"{total_rows:,}"

    num_cols = list(df.select_dtypes(include=[np.number]).columns)
    cat_cols = [c for c in df.columns if c not in num_cols]

    # Calculate smart KPIs
    for num_c in num_cols[:2]:
        total_val = df[num_c].sum()
        avg_val = df[num_c].mean()
        if "id" not in num_c.lower() and "zip" not in num_c.lower():
            kpis[f"Total {num_c}"] = f"{total_val:,.2f}" if isinstance(total_val, float) else f"{total_val:,}"
            kpis[f"Avg {num_c}"] = f"{avg_val:,.2f}"

    if px:
        # Chart 1: Categorical Breakdown (Bar or Pie)
        if cat_cols and num_cols:
            first_cat = cat_cols[0]
            first_num = [c for c in num_cols if "id" not in c.lower()][0] if any("id" not in c.lower() for c in num_cols) else num_cols[0]
            grouped = df.groupby(first_cat, as_index=False)[first_num].sum().sort_values(by=first_num, ascending=False).head(10)
            
            fig1 = px.bar(
                grouped, x=first_cat, y=first_num,
                title=f"Top 10 {first_cat} by Total {first_num}",
                template="plotly_white",
                color_discrete_sequence=["#10B981"]
            )
            charts.append(fig1)

            # Chart 2: Distribution Pie Chart
            if len(cat_cols) > 1:
                second_cat = cat_cols[1]
                grp2 = df.groupby(second_cat, as_index=False)[first_num].sum().head(8)
                fig2 = px.pie(
                    grp2, names=second_cat, values=first_num,
                    title=f"Distribution by {second_cat}",
                    hole=0.4
                )
                charts.append(fig2)

        # Chart 3: Numeric Distribution / Histogram
        if num_cols:
            fig3 = px.histogram(
                df, x=num_cols[0],
                title=f"Distribution of {num_cols[0]}",
                template="plotly_white",
                color_discrete_sequence=["#3B82F6"]
            )
            charts.append(fig3)

    return {
        "kpis": kpis,
        "charts": charts
    }


def embed_native_excel_chart(
    worksheet: Any,
    chart_type: str = "bar",
    title: str = "Assistant Summary Chart",
    min_col: int = 1,
    min_row: int = 1,
    max_row: int = 10,
    max_col: int = 2,
    cell_placement: str = "E2"
) -> bool:
    """
    Inserts a real native openpyxl chart into the specified Excel worksheet.
    The chart will render dynamically inside Microsoft Excel or Google Sheets.
    """
    if openpyxl is None or BarChart is None:
        return False

    try:
        if chart_type == "pie":
            chart = PieChart()
            data = Reference(worksheet, min_col=max_col, min_row=min_row, max_row=max_row)
            labels = Reference(worksheet, min_col=min_col, min_row=min_row + 1, max_row=max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)
        elif chart_type == "line":
            chart = LineChart()
            data = Reference(worksheet, min_col=max_col, min_row=min_row, max_row=max_row)
            labels = Reference(worksheet, min_col=min_col, min_row=min_row + 1, max_row=max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)
        else:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            data = Reference(worksheet, min_col=max_col, min_row=min_row, max_row=max_row)
            labels = Reference(worksheet, min_col=min_col, min_row=min_row + 1, max_row=max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)

        chart.title = title
        chart.width = 16
        chart.height = 10
        worksheet.add_chart(chart, cell_placement)
        return True
    except Exception as e:
        print("Failed to embed native Excel chart:", e)
        return False
