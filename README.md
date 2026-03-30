# 📊 India Operations Weekly KPI Dashboard

An end-to-end data engineering and business intelligence project 
simulating a real-world Finance operations reporting system 
inspired by Amazon's India e-commerce operations.

## 🚀 Project Overview

This project automates the weekly KPI reporting process for 
India e-commerce operations, tracking 8 financial metrics 
across 5 regions over 26 weeks using Python, SQL, Excel, 
and Power BI.

## 📁 Project Structure

├── generate_india_ops_data.py  # Generates synthetic dataset
├── kpi_queries.sql             # 10 analytical SQL queries
├── build_dashboard.py          # Builds Excel dashboard
├── ops_data.csv                # Weekly operations data
├── product_data.csv            # Product category data
├── region_targets.csv          # Regional budget targets
└── weekly_kpi_report.xlsx      # Final Excel dashboard

## 📊 KPIs Tracked

| KPI | Description |
|---|---|
| Total Revenue | Weekly revenue per region |
| GPM % | Gross Profit Margin |
| EBITDA Margin % | Earnings before interest & tax |
| Budget Variance % | Actual vs budget revenue |
| Cost per Order | Fulfillment cost efficiency |
| Return Rate % | Product return rate by category |
| Revenue WoW % | Week-over-week revenue change |

## 🗺️ Regions Covered
Mumbai · Delhi · Bangalore · Chennai · Hyderabad

## 🛠️ Tech Stack

- **Python** — Data generation, ETL, Excel automation
- **SQL** — CTEs, LAG window functions, rolling averages
- **SQLite** — Data warehouse schema
- **Power BI** — Interactive dashboard, DAX measures
- **Excel** — 5-sheet formatted KPI report

## 📈 Power BI Dashboard Pages

1. **Executive Summary** — 4 KPI cards, revenue vs budget chart
2. **Region Scorecard** — Matrix table with all KPIs by region
3. **Cost Analysis** — Cost per order trend by week
4. **Category Returns** — Return rate by product category

## 💡 Key SQL Techniques Used

- Common Table Expressions (CTEs)
- LAG window functions for WoW comparisons
- Rolling averages for trend analysis
- Subqueries for KPI benchmarking

## 📬 Contact

**Vaishnav S** — Data Engineer  
📧 vaishnavsudha2003@gmail.com  
🔗 [LinkedIn](https://linkedin.com/in/vaishnav-s-7ab082299)  
🌐 [Portfolio](https://my-portfolio-seven-sand-24.vercel.app/)
