# Automated-Data-Report-Generator-
# Automated Data Report Generator

Reads raw Excel sales data, computes KPIs automatically, generates AI-style plain-English business insights, and produces a formatted 4-sheet Excel report in under 30 seconds.

No API key needed.

---

## What It Does

- Reads any Excel sales file
- Computes KPIs: total revenue, units sold, return rate, growth, top products and regions
- AI Insight Engine generates 7 plain-English business observations automatically
- Builds a fully formatted Excel report with charts, color coding, and executive summary

---

## Report Sheets

| Sheet | Contents |
|---|---|
| Executive Summary | KPI boxes + 7 AI-generated insights |
| Monthly Trend | Month-by-month revenue line chart |
| Product Analysis | Product revenue bar chart + return rates |
| Raw Data | Full cleaned dataset |

---

## Setup and Run (Windows)

Step 1 - Install dependencies

    pip install -r requirements.txt

Step 2 - Run the script

    python report_generator.py

The script automatically generates sample sales data on first run.
Replace sales_data.xlsx with your own data and re-run for instant reports.

---

## Tech Stack

| Tool | Purpose |
|---|---|
| Python + Pandas | Data loading, cleaning, KPI computation |
| OpenPyXL | Excel report building with charts and formatting |
| AI Insight Engine | Rule-based plain-English insight generation |

---

## Sample AI Insights Generated

    Business processed 14,250 units generating Rs 82,34,500 total revenue.
    Monthly Revenue grew sharply by 23.4% compared to previous period.
    Laptop is the top performer, contributing 28.1% of total revenue.
    Keyboard is underperforming - 1,240 units below average. Consider promotional push.
    Recommended: Scale marketing in North region. Investigate South region low performance.

---

## Business Use Case

Analysts at MNCs spend hours every week manually building Excel performance reports.
This tool automates the entire process - point it at any sales Excel file and get
a boardroom-ready report in under 30 seconds.

---

Author: Yashwanth Eluri
LinkedIn: https://www.linkedin.com/in/yashwanth-eluri--analyst/
GitHub: https://github.com/yashwantheluri6751
