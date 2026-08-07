# =============================================================================
# Daquan Morrison · Data Analyst · Syracuse City School District
# Role: Crime Analyst 2 · DCJS CNY Crime Analysis Center
# Project: NY Crime Analytics Lakehouse — Silver Transformation Layer
# Notebook: ny_crime_silver_transform.py
# Description:
#   Transforms Bronze raw data into a clean, typed, enriched Silver table.
#   Casts types, standardizes strings, derives calculated fields, and applies
#   data quality flags. Writes to silver_nys_crime (Delta format).
# Source Table: MainLakehouse.dbo.bronze_nys_crime
# Target Table: MainLakehouse.dbo.silver_nys_crime
# =============================================================================

# --------------------------------------------------------------------------
# Cell 1 — Imports & Configuration
# --------------------------------------------------------------------------
from pyspark.sql import SparkSession
from pyspark.sql import functions as F
from pyspark.sql.functions import (
    col, trim, initcap, cast,
    when, floor, lit,
    coalesce
)
from pyspark.sql.types import IntegerType, DoubleType, StringType

spark = SparkSession.builder.getOrCreate()

SOURCE_TABLE = "bronze_nys_crime"
TARGET_TABLE = "silver_nys_crime"

print(f"[INFO] Silver Transform: {SOURCE_TABLE} → {TARGET_TABLE}")


# --------------------------------------------------------------------------
# Cell 2 — Read Bronze Table
# --------------------------------------------------------------------------
print(f"[INFO] Reading Bronze table: {SOURCE_TABLE}")
df_bronze = spark.table(SOURCE_TABLE)

print(f"[INFO] Bronze row count: {df_bronze.count():,}")
print("[INFO] Bronze schema:")
df_bronze.printSchema()


# --------------------------------------------------------------------------
# Cell 3 — Type Casting (Bronze strings → proper types)
# --------------------------------------------------------------------------
print("[INFO] Casting column types...")

df_cast = (
    df_bronze
    # Integer fields
    .withColumn("year",             col("year").cast(IntegerType()))
    .withColumn("months_reported",  col("months_reported").cast(IntegerType()))

    # Double (numeric) fields — all crime count columns
    .withColumn("total_index_crimes",           col("total_index_crimes").cast(DoubleType()))
    .withColumn("violent_crimes_total",         col("violent_crimes_total").cast(DoubleType()))
    .withColumn("murder_non_negl_manslaughter", col("murder_non_negl_manslaughter").cast(DoubleType()))
    .withColumn("rape_legacy_definition",       col("rape_legacy_definition").cast(DoubleType()))
    .withColumn("robbery",                      col("robbery").cast(DoubleType()))
    .withColumn("aggravated_assault",           col("aggravated_assault").cast(DoubleType()))
    .withColumn("property_crimes_total",        col("property_crimes_total").cast(DoubleType()))
    .withColumn("burglary",                     col("burglary").cast(DoubleType()))
    .withColumn("larceny",                      col("larceny").cast(DoubleType()))
    .withColumn("motor_vehicle_theft",          col("motor_vehicle_theft").cast(DoubleType()))

    # String standardization: strip whitespace + Title Case
    .withColumn("county", initcap(trim(col("county"))))
    .withColumn("agency", initcap(trim(col("agency"))))
)

print("[INFO] Type casting complete.")


# --------------------------------------------------------------------------
# Cell 4 — Calculated Fields
# --------------------------------------------------------------------------
print("[INFO] Computing derived/calculated fields...")

df_enriched = (
    df_cast

    # --- Recalculated violent crime total from components ---
    # (more reliable than trusting the source aggregate column)
    .withColumn(
        "violent_crime_total",
        coalesce(col("murder_non_negl_manslaughter"), lit(0.0))
        + coalesce(col("rape_legacy_definition"),      lit(0.0))
        + coalesce(col("robbery"),                     lit(0.0))
        + coalesce(col("aggravated_assault"),          lit(0.0))
    )

    # --- Recalculated property crime total from components ---
    .withColumn(
        "property_crime_total",
        coalesce(col("burglary"),           lit(0.0))
        + coalesce(col("larceny"),          lit(0.0))
        + coalesce(col("motor_vehicle_theft"), lit(0.0))
    )

    # --- Crime Category: dominant crime type ---
    .withColumn(
        "crime_category",
        when(col("violent_crime_total") > col("property_crime_total"), "Violent-Heavy")
        .when(col("property_crime_total") > col("violent_crime_total"), "Property-Heavy")
        .otherwise("Mixed")
    )

    # --- Decade: group years into decades (1990, 2000, 2010, 2020, ...) ---
    .withColumn(
        "decade",
        (floor(col("year") / 10) * 10).cast(IntegerType())
    )

    # --- Is County Total: flag aggregated county-level rows ---
    # Agencies that represent county totals contain 'County Total' in the name
    .withColumn(
        "is_county_total",
        when(
            F.upper(col("agency")).contains("COUNTY TOTAL"), 1
        ).otherwise(0).cast(IntegerType())
    )

    # --- Data Quality Flag: full 12-month reporting or partial ---
    .withColumn(
        "data_quality_flag",
        when(col("months_reported") < 12, "INCOMPLETE")
        .when(col("months_reported").isNull(), "INCOMPLETE")
        .otherwise("COMPLETE")
    )

    # --- Crime Rate per 100k: placeholder — requires population join ---
    # Will be populated in a future enrichment notebook when census data is joined
    .withColumn(
        "crime_rate_per_100k",
        lit(None).cast(DoubleType())
    )
)

print("[INFO] Derived fields computed.")


# --------------------------------------------------------------------------
# Cell 5 — Column Selection & Ordering
# --------------------------------------------------------------------------
print("[INFO] Selecting and ordering final Silver columns...")

SILVER_COLUMNS = [
    # Identifiers
    "county",
    "agency",
    "year",
    "months_reported",

    # Raw crime totals (from source)
    "total_index_crimes",
    "violent_crimes_total",       # Original source aggregate
    "property_crimes_total",      # Original source aggregate

    # Raw crime components
    "murder_non_negl_manslaughter",
    "rape_legacy_definition",
    "robbery",
    "aggravated_assault",
    "burglary",
    "larceny",
    "motor_vehicle_theft",

    # Derived totals (computed from components)
    "violent_crime_total",
    "property_crime_total",

    # Derived categorical/analytic fields
    "crime_category",
    "decade",
    "is_county_total",
    "data_quality_flag",
    "crime_rate_per_100k",        # NULL placeholder

    # Audit
    "ingested_at",
    "source_api",
]

df_silver = df_enriched.select(SILVER_COLUMNS)


# --------------------------------------------------------------------------
# Cell 6 — Preview & Validation
# --------------------------------------------------------------------------
print("\n[INFO] Silver DataFrame preview (top 10 rows):")
df_silver.show(10, truncate=False)

print("\n[INFO] Silver schema:")
df_silver.printSchema()

print(f"\n[INFO] Total rows: {df_silver.count():,}")

# Distribution of crime_category
print("\n[INFO] Crime Category distribution:")
df_silver.groupBy("crime_category").count().orderBy("count", ascending=False).show()

# Distribution of data_quality_flag
print("\n[INFO] Data Quality Flag distribution:")
df_silver.groupBy("data_quality_flag").count().orderBy("count", ascending=False).show()

# is_county_total distribution
print("\n[INFO] Is County Total distribution:")
df_silver.groupBy("is_county_total").count().show()

# Year range
print("\n[INFO] Year range in Silver:")
df_silver.selectExpr("MIN(year) AS min_year", "MAX(year) AS max_year").show()


# --------------------------------------------------------------------------
# Cell 7 — Write Silver Table
# --------------------------------------------------------------------------
print(f"\n[INFO] Writing Silver table: {TARGET_TABLE}...")

(
    df_silver.write
             .format("delta")
             .mode("overwrite")
             .option("overwriteSchema", "true")
             .saveAsTable(TARGET_TABLE)
)

print(f"[SUCCESS] Silver table '{TARGET_TABLE}' written successfully.")
print(f"[SUCCESS] Rows written: {df_silver.count():,}")


# --------------------------------------------------------------------------
# Cell 8 — Post-Write Verification
# --------------------------------------------------------------------------
print("\n=== SILVER TABLE VERIFICATION ===")

df_verify = spark.table(TARGET_TABLE)
print(f"Rows in {TARGET_TABLE}: {df_verify.count():,}")

# Null check on critical fields
from pyspark.sql.functions import sum as spark_sum, when as spark_when

null_check_cols = ["county", "agency", "year", "total_index_crimes",
                   "violent_crime_total", "property_crime_total"]
null_counts = df_verify.select([
    spark_sum(spark_when(col(c).isNull(), 1).otherwise(0)).alias(c)
    for c in null_check_cols
]).collect()[0].asDict()

print("\nNull counts in Silver key columns:")
for c, n in null_counts.items():
    status = "OK" if n == 0 else f"WARNING: {n:,} nulls"
    print(f"  {c}: {status}")

# Decade distribution
print("\nDecade distribution:")
df_verify.groupBy("decade").count().orderBy("decade").show()

print("=== SILVER VERIFICATION COMPLETE ===")
