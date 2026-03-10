

import pandas as pd
import numpy as np
import sqlite3
import os
from datetime import date, timedelta

# ─────────────────────────────────────────────
# CONFIG — tweak these to change dataset size
# ─────────────────────────────────────────────
WEEKS        = 26          # 6 months of weekly data
START_DATE   = date(2024, 1, 1)
RANDOM_SEED  = 42

REGIONS = {
    "Mumbai"    : {"base_rev": 5_200_000, "base_orders": 22_000, "cost_factor": 1.05},
    "Delhi"     : {"base_rev": 4_800_000, "base_orders": 20_500, "cost_factor": 1.02},
    "Bangalore" : {"base_rev": 4_400_000, "base_orders": 19_000, "cost_factor": 0.97},
    "Chennai"   : {"base_rev": 3_600_000, "base_orders": 15_500, "cost_factor": 0.94},
    "Hyderabad" : {"base_rev": 3_200_000, "base_orders": 13_800, "cost_factor": 0.98},
}

CATEGORIES = ["Electronics", "Fashion", "Grocery", "Home & Kitchen", "Books"]
CAT_WEIGHTS = [0.30, 0.25, 0.20, 0.15, 0.10]   # revenue share per category

np.random.seed(RANDOM_SEED)


# ─────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────

def seasonal_factor(week_num):
    """
    Simulate Indian e-commerce seasonality:
    - Dip in Jan-Feb (post-festive hangover)
    - Spike in wk 14-16 (April sale events)
    - Big spike in wk 40-44 (Diwali / Big Billion Days)
    """
    base = 1.0
    # April sale bump
    if 13 <= week_num <= 16:
        base += 0.12
    # Mid-year prime-style sale
    if 24 <= week_num <= 26:
        base += 0.08
    return base


def add_noise(value, pct=0.06):
    """Add Gaussian noise of ±pct to a value."""
    return value * np.random.normal(1.0, pct)


def returns_rate(category):
    """Returns rate varies by category (realistic e-commerce benchmarks)."""
    rates = {
        "Electronics"    : 0.06,
        "Fashion"        : 0.14,   # fashion has higher returns
        "Grocery"        : 0.02,
        "Home & Kitchen" : 0.05,
        "Books"          : 0.03,
    }
    return rates.get(category, 0.05)


# ─────────────────────────────────────────────
# 1. WEEKLY OPS TABLE  (main fact table)
# ─────────────────────────────────────────────

def generate_weekly_ops():
    rows = []

    for week_num in range(1, WEEKS + 1):
        week_start = START_DATE + timedelta(weeks=week_num - 1)
        growth     = 1 + (week_num - 1) * 0.007     # ~0.7% weekly compounding growth
        season     = seasonal_factor(week_num)

        for region, cfg in REGIONS.items():

            # ── Revenue & Orders ──────────────────────────
            revenue = round(add_noise(cfg["base_rev"] * growth * season), 0)
            orders  = int(add_noise(cfg["base_orders"] * growth * season, pct=0.05))

            # Budget is set at start of month (less volatile, slightly below actuals on avg)
            budget_revenue = round(cfg["base_rev"] * growth * 0.98, 0)
            budget_orders  = int(cfg["base_orders"] * growth * 0.98)

            # ── Cost Structure ────────────────────────────
            # COGS: 58–65% of revenue (varies by week / region)
            cogs_pct         = np.random.uniform(0.58, 0.65)
            cogs             = round(revenue * cogs_pct, 0)

            # OPEX: 18–24% of revenue
            opex_pct         = np.random.uniform(0.18, 0.24)
            opex             = round(revenue * opex_pct, 0)

            # Fulfillment cost per order: ₹40–₹60 depending on region
            base_cpo         = 48 * cfg["cost_factor"]
            fulfillment_cost = round(orders * add_noise(base_cpo, pct=0.08), 0)

            # Marketing spend: 4–8% of revenue
            marketing_spend  = round(revenue * np.random.uniform(0.04, 0.08), 0)

            # ── Workforce ─────────────────────────────────
            # Headcount grows slightly over time, with seasonal warehousing spikes
            base_hc   = {"Mumbai":145, "Delhi":130, "Bangalore":118, "Chennai":95, "Hyderabad":88}
            headcount = int(base_hc[region] + week_num * 0.4 + np.random.randint(-4, 5))
            if season > 1.08:
                headcount += np.random.randint(8, 18)   # temp workers during sale events

            # ── Derived Metrics ───────────────────────────
            gross_profit  = revenue - cogs
            ebitda        = revenue - cogs - opex
            net_revenue   = revenue - round(revenue * np.random.uniform(0.02, 0.04), 0)  # after discounts

            rows.append({
                "week_num"          : week_num,
                "week_start"        : week_start.isoformat(),
                "region"            : region,
                "revenue"           : revenue,
                "net_revenue"       : net_revenue,
                "budget_revenue"    : budget_revenue,
                "cogs"              : cogs,
                "gross_profit"      : gross_profit,
                "opex"              : opex,
                "marketing_spend"   : marketing_spend,
                "fulfillment_cost"  : fulfillment_cost,
                "ebitda"            : ebitda,
                "orders"            : orders,
                "budget_orders"     : budget_orders,
                "headcount"         : headcount,
                "avg_order_value"   : round(revenue / orders, 2),
                "cost_per_order"    : round(fulfillment_cost / orders, 2),
                "gpm_pct"           : round(gross_profit / revenue * 100, 2),
                "ebitda_margin_pct" : round(ebitda / revenue * 100, 2),
                "opex_pct"          : round(opex / revenue * 100, 2),
                "budget_var_pct"    : round((revenue - budget_revenue) / budget_revenue * 100, 2),
                "rev_per_headcount" : round(revenue / headcount, 0),
            })

    return pd.DataFrame(rows)


# ─────────────────────────────────────────────
# 2. PRODUCT CATEGORY TABLE  (dimension table)
# ─────────────────────────────────────────────

def generate_product_data(weekly_ops_df):
    rows = []
    for _, row in weekly_ops_df.iterrows():
        # Split weekly revenue across categories using weights + noise
        noisy_weights = [w * add_noise(1.0, 0.15) for w in CAT_WEIGHTS]
        total_w = sum(noisy_weights)
        splits  = [w / total_w for w in noisy_weights]

        for cat, share in zip(CATEGORIES, splits):
            cat_rev    = round(row["revenue"] * share, 0)
            cat_orders = int(row["orders"] * share * add_noise(1.0, 0.10))
            ret_rate   = returns_rate(cat)
            returned   = int(cat_orders * np.random.uniform(ret_rate * 0.7, ret_rate * 1.3))

            rows.append({
                "week_num"      : row["week_num"],
                "week_start"    : row["week_start"],
                "region"        : row["region"],
                "category"      : cat,
                "revenue"       : cat_rev,
                "orders"        : cat_orders,
                "returns"       : returned,
                "net_orders"    : cat_orders - returned,
                "return_rate_pct": round(returned / max(cat_orders, 1) * 100, 2),
                "avg_order_value": round(cat_rev / max(cat_orders, 1), 2),
            })

    return pd.DataFrame(rows)


# ─────────────────────────────────────────────
# 3. REGION TARGETS TABLE  (budget/targets)
# ─────────────────────────────────────────────

def generate_region_targets():
    rows = []
    for week_num in range(1, WEEKS + 1):
        week_start = START_DATE + timedelta(weeks=week_num - 1)
        growth     = 1 + (week_num - 1) * 0.007

        for region, cfg in REGIONS.items():
            rows.append({
                "week_num"          : week_num,
                "week_start"        : week_start.isoformat(),
                "region"            : region,
                "revenue_target"    : round(cfg["base_rev"] * growth * 0.98, 0),
                "orders_target"     : int(cfg["base_orders"] * growth * 0.98),
                "gpm_target_pct"    : 38.5,
                "cpo_target"        : round(48 * cfg["cost_factor"], 2),
                "ebitda_target_pct" : 15.0,
            })

    return pd.DataFrame(rows)


# ─────────────────────────────────────────────
# 4. SAVE TO CSV + SQLITE
# ─────────────────────────────────────────────

def save_all(ops_df, product_df, targets_df):
    # CSV files
    ops_df.to_csv("ops_data.csv", index=False)
    product_df.to_csv("product_data.csv", index=False)
    targets_df.to_csv("region_targets.csv", index=False)

    # SQLite database
    conn = sqlite3.connect("ops_finance.db")
    ops_df.to_sql("weekly_ops",      conn, if_exists="replace", index=False)
    product_df.to_sql("product_data", conn, if_exists="replace", index=False)
    targets_df.to_sql("region_targets", conn, if_exists="replace", index=False)

    # Create a useful summary view in SQLite
    conn.execute("""
        CREATE VIEW IF NOT EXISTS v_weekly_summary AS
        SELECT
            week_num,
            week_start,
            SUM(revenue)           AS total_revenue,
            SUM(budget_revenue)    AS total_budget,
            SUM(gross_profit)      AS total_gross_profit,
            SUM(ebitda)            AS total_ebitda,
            SUM(orders)            AS total_orders,
            ROUND(SUM(gross_profit)*100.0/SUM(revenue), 2) AS gpm_pct,
            ROUND(SUM(ebitda)*100.0/SUM(revenue), 2)       AS ebitda_margin_pct,
            ROUND(SUM(fulfillment_cost)*1.0/SUM(orders), 2) AS avg_cost_per_order,
            ROUND((SUM(revenue)-SUM(budget_revenue))*100.0/SUM(budget_revenue), 2) AS budget_var_pct
        FROM weekly_ops
        GROUP BY week_num, week_start
        ORDER BY week_num
    """)
    conn.commit()
    conn.close()

    print("✅ Files created:")
    for f in ["ops_data.csv", "product_data.csv", "region_targets.csv", "ops_finance.db"]:
        size = os.path.getsize(f)
        print(f"   {f:30s}  {size:,} bytes")


# ─────────────────────────────────────────────
# 5. PRINT SAMPLE SUMMARY
# ─────────────────────────────────────────────

def print_summary(ops_df):
    print("\n" + "═"*55)
    print("  DATASET SUMMARY")
    print("═"*55)
    print(f"  Weeks generated      : {ops_df['week_num'].nunique()}")
    print(f"  Regions              : {', '.join(ops_df['region'].unique())}")
    print(f"  Total rows           : {len(ops_df):,}")
    print(f"  Date range           : {ops_df['week_start'].min()} → {ops_df['week_start'].max()}")
    print()

    total = ops_df.groupby("week_num").agg(
        revenue=("revenue","sum"),
        orders=("orders","sum"),
        gpm=("gpm_pct","mean"),
        ebitda=("ebitda","sum"),
    ).reset_index()

    print(f"  Avg weekly revenue   : ₹{total['revenue'].mean():,.0f}")
    print(f"  Avg weekly orders    : {total['orders'].mean():,.0f}")
    print(f"  Avg GPM              : {total['gpm'].mean():.1f}%")
    print(f"  Total EBITDA (26wks) : ₹{ops_df['ebitda'].sum():,.0f}")
    print()

    print("  Top region by revenue:")
    top = ops_df.groupby("region")["revenue"].sum().sort_values(ascending=False)
    for region, rev in top.items():
        print(f"    {region:12s}  ₹{rev:,.0f}")

    print()
    print("  Sample KPIs (latest week):")
    latest = ops_df[ops_df["week_num"] == ops_df["week_num"].max()]
    print(f"    Revenue       : ₹{latest['revenue'].sum():,.0f}")
    print(f"    GPM %         : {latest['gpm_pct'].mean():.1f}%")
    print(f"    Cost/Order    : ₹{latest['cost_per_order'].mean():.1f}")
    print(f"    Budget Var %  : {latest['budget_var_pct'].mean():+.1f}%")
    print("═"*55)


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print("⏳ Generating India Ops data...")

    ops_df     = generate_weekly_ops()
    product_df = generate_product_data(ops_df)
    targets_df = generate_region_targets()

    save_all(ops_df, product_df, targets_df)
    print_summary(ops_df)

    print("\n🚀 Next step: run your SQL queries against ops_finance.db")
    print("   sqlite3 ops_finance.db")
    print("   SELECT * FROM v_weekly_summary LIMIT 5;")
