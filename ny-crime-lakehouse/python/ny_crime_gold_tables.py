# =============================================================================
# Daquan Morrison · Data Analyst · Syracuse City School District
# Role: Crime Analyst 2 · DCJS CNY Crime Analysis Center
# Project: NY Crime Analytics Lakehouse — Gold Layer (Analytical Tables)
# Notebook: ny_crime_gold_tables.py
# Description:
#   Builds all 6 Gold analytical tables from the Silver layer.
#   Gold tables are purpose-built for Power BI reporting, forecasting,
#   ranking, and trend analysis. County-level aggregations filter to
#   is_county_total = 1; agency-level tables use the full dataset.
# Source Table: MainLakehouse.dbo.silver_nys_crime
# Target Tables:
#   - gold_county_year_crime
#   - gold_county_rankings
#   - gold_violent_crime_trends
#   - gold_property_crime_trends
#   - gold_agency_activity
#   - gold_forecast_base
# =============================================================================

# --------------------------------------------------------------------------
# Cell 1 — Imports & Configuration
# --------------------------------------------------------------------------
from pyspark.sql import SparkSession, Window
from pyspark.sql import functions as F
from pyspark.sql.functions import (
    col, lit, sum as spark_sum, avg as spark_avg,
    round as spark_round, rank, lag, coalesce,
    dense_rank, first, last
)
from pyspark.sql.window import Window

spark = SparkSession.builder.getOrCreate()

SOURCE_TABLE = "silver_nys_crime"

GOLD_TABLES = {
    "county_year":        "gold_county_year_crime",
    "county_rankings":    "gold_county_rankings",
    "violent_trends":     "gold_violent_crime_trends",
    "property_trends":    "gold_property_crime_trends",
    "agency_activity":    "gold_agency_activity",
    "forecast_base":      "gold_forecast_base",
}

print(f"[INFO] Loading Silver source table: {SOURCE_TABLE}")


# --------------------------------------------------------------------------
# Cell 2 — Load Silver Data
# --------------------------------------------------------------------------
df_silver = spark.table(SOURCE_TABLE)

# County-level subset (is_county_total = 1) — used for most Gold tables
df_county = df_silver.filter(col("is_county_total") == 1)

# Agency-level subset (all rows) — used for gold_agency_activity
df_agency = df_silver  # full dataset

print(f"[INFO] Silver total rows:        {df_silver.count():,}")
print(f"[INFO] County-total rows:        {df_county.count():,}")
print(f"[INFO] Agency-level rows:        {df_agency.count():,}")


# --------------------------------------------------------------------------
# Cell 3 — gold_county_year_crime
# --------------------------------------------------------------------------
# Core analytical table: one row per county per year (county totals only)
# Columns: county, year, total_index_crimes, violent_crime_total,
#          property_crime_total, crime_category, decade, data_quality_flag
# --------------------------------------------------------------------------
print(f"\n[INFO] Building {GOLD_TABLES['county_year']}...")

df_gold_county_year = (
    df_county
    .select(
        "county",
        "year",
        "total_index_crimes",
        "violent_crime_total",
        "property_crime_total",
        "crime_category",
        "decade",
        "data_quality_flag",
    )
    .orderBy("county", "year")
)

df_gold_county_year.write \
    .format("delta") \
    .mode("overwrite") \
    .option("overwriteSchema", "true") \
    .saveAsTable(GOLD_TABLES["county_year"])

print(f"[SUCCESS] {GOLD_TABLES['county_year']}: {df_gold_county_year.count():,} rows written.")
df_gold_county_year.show(5, truncate=False)


# --------------------------------------------------------------------------
# Cell 4 — gold_county_rankings
# --------------------------------------------------------------------------
# For each year, ranks all counties by total_index_crimes DESC.
# Enables "Top 10 Counties" reporting and year-over-year rank movement.
# Columns: year, county, total_index_crimes, county_rank,
#          violent_crime_total, property_crime_total
# --------------------------------------------------------------------------
print(f"\n[INFO] Building {GOLD_TABLES['county_rankings']}...")

# Window: partition by year, order by total_index_crimes descending
window_rank = Window.partitionBy("year").orderBy(col("total_index_crimes").desc())

df_gold_rankings = (
    df_county
    .select(
        "year",
        "county",
        "total_index_crimes",
        "violent_crime_total",
        "property_crime_total",
    )
    .withColumn("county_rank", dense_rank().over(window_rank))
    .select(
        "year",
        "county",
        "total_index_crimes",
        "county_rank",
        "violent_crime_total",
        "property_crime_total",
    )
    .orderBy("year", "county_rank")
)

df_gold_rankings.write \
    .format("delta") \
    .mode("overwrite") \
    .option("overwriteSchema", "true") \
    .saveAsTable(GOLD_TABLES["county_rankings"])

print(f"[SUCCESS] {GOLD_TABLES['county_rankings']}: {df_gold_rankings.count():,} rows written.")
df_gold_rankings.show(10, truncate=False)


# --------------------------------------------------------------------------
# Cell 5 — gold_violent_crime_trends
# --------------------------------------------------------------------------
# Time-series trend table for violent crime components.
# Includes YoY change and 3-year rolling average.
# Columns: county, year, murder, rape_legacy, robbery, aggravated_assault,
#          violent_crime_total, yoy_change, rolling_3yr_avg
# --------------------------------------------------------------------------
print(f"\n[INFO] Building {GOLD_TABLES['violent_trends']}...")

# Window for YoY and rolling avg: partition by county, order by year
window_county_year = (
    Window
    .partitionBy("county")
    .orderBy("year")
)

window_rolling_3yr = (
    Window
    .partitionBy("county")
    .orderBy("year")
    .rowsBetween(-2, 0)    # Current row + 2 preceding = 3-year window
)

df_gold_violent = (
    df_county
    .select(
        "county",
        "year",
        col("murder_non_negl_manslaughter").alias("murder"),
        col("rape_legacy_definition").alias("rape_legacy"),
        "robbery",
        "aggravated_assault",
        "violent_crime_total",
    )
    .withColumn(
        "yoy_change",
        col("violent_crime_total") - lag("violent_crime_total", 1).over(window_county_year)
    )
    .withColumn(
        "rolling_3yr_avg",
        spark_round(spark_avg("violent_crime_total").over(window_rolling_3yr), 2)
    )
    .orderBy("county", "year")
)

df_gold_violent.write \
    .format("delta") \
    .mode("overwrite") \
    .option("overwriteSchema", "true") \
    .saveAsTable(GOLD_TABLES["violent_trends"])

print(f"[SUCCESS] {GOLD_TABLES['violent_trends']}: {df_gold_violent.count():,} rows written.")
df_gold_violent.show(10, truncate=False)


# --------------------------------------------------------------------------
# Cell 6 — gold_property_crime_trends
# --------------------------------------------------------------------------
# Time-series trend table for property crime components.
# Mirrors the violent crime trends structure.
# Columns: county, year, burglary, larceny, motor_vehicle_theft,
#          property_crime_total, yoy_change, rolling_3yr_avg
# --------------------------------------------------------------------------
print(f"\n[INFO] Building {GOLD_TABLES['property_trends']}...")

df_gold_property = (
    df_county
    .select(
        "county",
        "year",
        "burglary",
        "larceny",
        "motor_vehicle_theft",
        "property_crime_total",
    )
    .withColumn(
        "yoy_change",
        col("property_crime_total") - lag("property_crime_total", 1).over(window_county_year)
    )
    .withColumn(
        "rolling_3yr_avg",
        spark_round(spark_avg("property_crime_total").over(window_rolling_3yr), 2)
    )
    .orderBy("county", "year")
)

df_gold_property.write \
    .format("delta") \
    .mode("overwrite") \
    .option("overwriteSchema", "true") \
    .saveAsTable(GOLD_TABLES["property_trends"])

print(f"[SUCCESS] {GOLD_TABLES['property_trends']}: {df_gold_property.count():,} rows written.")
df_gold_property.show(10, truncate=False)


# --------------------------------------------------------------------------
# Cell 7 — gold_agency_activity
# --------------------------------------------------------------------------
# Agency-level detail table (NOT filtered to county totals).
# Used to analyze municipal/PD-level reporting completeness and crime data.
# Columns: county, agency, year, total_index_crimes, months_reported,
#          data_quality_flag
# --------------------------------------------------------------------------
print(f"\n[INFO] Building {GOLD_TABLES['agency_activity']}...")

df_gold_agency = (
    df_agency
    .select(
        "county",
        "agency",
        "year",
        "total_index_crimes",
        "months_reported",
        "data_quality_flag",
    )
    .orderBy("county", "agency", "year")
)

df_gold_agency.write \
    .format("delta") \
    .mode("overwrite") \
    .option("overwriteSchema", "true") \
    .saveAsTable(GOLD_TABLES["agency_activity"])

print(f"[SUCCESS] {GOLD_TABLES['agency_activity']}: {df_gold_agency.count():,} rows written.")
df_gold_agency.show(10, truncate=False)


# --------------------------------------------------------------------------
# Cell 8 — gold_forecast_base
# --------------------------------------------------------------------------
# Clean base table prepared for time-series forecasting (Prophet/ARIMA)
# or Power BI Analytics pane. One row per county per year.
# Columns: county, year, total_index_crimes, violent_crime_total,
#          property_crime_total — ordered by county, year ascending.
# --------------------------------------------------------------------------
print(f"\n[INFO] Building {GOLD_TABLES['forecast_base']}...")

df_gold_forecast = (
    df_county
    .select(
        "county",
        "year",
        "total_index_crimes",
        "violent_crime_total",
        "property_crime_total",
    )
    .orderBy("county", "year")
)

df_gold_forecast.write \
    .format("delta") \
    .mode("overwrite") \
    .option("overwriteSchema", "true") \
    .saveAsTable(GOLD_TABLES["forecast_base"])

print(f"[SUCCESS] {GOLD_TABLES['forecast_base']}: {df_gold_forecast.count():,} rows written.")
df_gold_forecast.show(10, truncate=False)


# --------------------------------------------------------------------------
# Cell 9 — Gold Layer Summary
# --------------------------------------------------------------------------
print("\n" + "=" * 60)
print("GOLD LAYER BUILD SUMMARY")
print("=" * 60)

for key, table_name in GOLD_TABLES.items():
    count = spark.table(table_name).count()
    print(f"  {table_name:<35} {count:>8,} rows")

print("=" * 60)
print("[SUCCESS] All 6 Gold tables written successfully.")
print("[INFO] Gold tables are ready for Power BI semantic model connection.")
print("[INFO] Connect via: Workspace → MainLakehouse → SQL Analytics Endpoint")


# --------------------------------------------------------------------------
# Cell 10 — Quick Business Validation (Onondaga County spot-check)
# --------------------------------------------------------------------------
print("\n=== SPOT CHECK: Onondaga County (Syracuse area) ===")

print("\nRecent Onondaga violent crime trends:")
spark.sql("""
    SELECT year, violent_crime_total, yoy_change, rolling_3yr_avg
    FROM gold_violent_crime_trends
    WHERE county = 'Onondaga'
    ORDER BY year DESC
    LIMIT 10
""").show()

print("\nOnondaga county ranking over last 5 years:")
spark.sql("""
    SELECT year, county, total_index_crimes, county_rank
    FROM gold_county_rankings
    WHERE county = 'Onondaga'
    ORDER BY year DESC
    LIMIT 5
""").show()

print("=== SPOT CHECK COMPLETE ===")
