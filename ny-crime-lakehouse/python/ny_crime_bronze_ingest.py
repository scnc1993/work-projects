# =============================================================================
# Daquan Morrison · Data Analyst · Syracuse City School District
# Role: Crime Analyst 2 · DCJS CNY Crime Analysis Center
# Project: NY Crime Analytics Lakehouse — Bronze Ingestion Layer
# Notebook: ny_crime_bronze_ingest.py
# Description:
#   Ingests NYS Index Crimes by County and Agency data from the NYS Open Data
#   Socrata API into the Bronze Delta table in Microsoft Fabric (MainLakehouse).
#   Handles full paginated extraction, raw schema preservation, and audit columns.
# Source API: https://data.ny.gov/resource/ca8h-8gjq.json
# Target Table: MainLakehouse.dbo.bronze_nys_crime (Delta format)
# =============================================================================

# --------------------------------------------------------------------------
# Cell 1 — Imports & Configuration
# --------------------------------------------------------------------------
import requests
import json
from datetime import datetime, timezone
from pyspark.sql import SparkSession
from pyspark.sql.types import (
    StructType, StructField,
    StringType, TimestampType
)
from pyspark.sql.functions import lit, current_timestamp

# Fabric Spark session (already available in Fabric notebooks — no need to create)
spark = SparkSession.builder.getOrCreate()

# --------------------------------------------------------------------------
# API Configuration
# --------------------------------------------------------------------------
API_BASE_URL   = "https://data.ny.gov/resource/ca8h-8gjq.json"
APP_TOKEN      = ""          # Optional: set your Socrata app token to raise rate limits
PAGE_LIMIT     = 2000        # Records per page request
MAX_RECORDS    = 200_000     # Safety cap; actual dataset is ~100k rows
TARGET_TABLE   = "bronze_nys_crime"
LAKEHOUSE_NAME = "MainLakehouse"

# Raw column names as returned by the Socrata API
RAW_COLUMNS = [
    "county",
    "agency",
    "year",
    "months_reported",
    "total_index_crimes",
    "violent_crimes_total",
    "murder_non_negl_manslaughter",
    "rape_legacy_definition",
    "robbery",
    "aggravated_assault",
    "property_crimes_total",
    "burglary",
    "larceny",
    "motor_vehicle_theft",
]

# --------------------------------------------------------------------------
# Cell 2 — Paginated API Extraction
# --------------------------------------------------------------------------
def fetch_all_records(base_url: str, page_limit: int, max_records: int, app_token: str = "") -> list:
    """
    Fetches all records from a Socrata API endpoint using limit/offset pagination.

    Args:
        base_url    : Socrata JSON endpoint URL.
        page_limit  : Number of records to request per page.
        max_records : Maximum total records to retrieve (safety cap).
        app_token   : Optional Socrata application token.

    Returns:
        List of raw record dicts from the API.
    """
    headers = {}
    if app_token:
        headers["X-App-Token"] = app_token

    all_records = []
    offset = 0

    print(f"[INFO] Starting paginated extraction from {base_url}")
    print(f"[INFO] Page size: {page_limit} | Max records cap: {max_records:,}")

    while True:
        params = {
            "$limit":  page_limit,
            "$offset": offset,
            "$order":  ":id",          # Stable ordering for reproducible pagination
        }

        response = requests.get(base_url, headers=headers, params=params, timeout=60)
        response.raise_for_status()
        page_data = response.json()

        if not page_data:
            print(f"[INFO] Empty page received at offset {offset}. Extraction complete.")
            break

        all_records.extend(page_data)
        offset += len(page_data)
        print(f"[INFO] Fetched {len(page_data):,} records | Total so far: {len(all_records):,}")

        if len(page_data) < page_limit:
            print(f"[INFO] Partial page received — end of dataset reached.")
            break

        if len(all_records) >= max_records:
            print(f"[WARN] Max records cap ({max_records:,}) reached. Stopping pagination.")
            break

    print(f"[INFO] Extraction finished. Total records retrieved: {len(all_records):,}")
    return all_records


# --------------------------------------------------------------------------
# Cell 3 — Normalize & Build Spark DataFrame
# --------------------------------------------------------------------------
def normalize_records(records: list, raw_columns: list) -> list:
    """
    Normalizes raw API records to ensure all expected columns are present
    (fills missing keys with None for schema consistency).

    Args:
        records     : List of raw dicts from the API.
        raw_columns : Expected column names.

    Returns:
        List of normalized dicts.
    """
    normalized = []
    for rec in records:
        row = {col: rec.get(col, None) for col in raw_columns}
        normalized.append(row)
    return normalized


# Define explicit schema — all raw columns stored as StringType (Bronze = raw)
bronze_schema = StructType([
    StructField("county",                        StringType(), True),
    StructField("agency",                        StringType(), True),
    StructField("year",                          StringType(), True),
    StructField("months_reported",               StringType(), True),
    StructField("total_index_crimes",            StringType(), True),
    StructField("violent_crimes_total",          StringType(), True),
    StructField("murder_non_negl_manslaughter",  StringType(), True),
    StructField("rape_legacy_definition",        StringType(), True),
    StructField("robbery",                       StringType(), True),
    StructField("aggravated_assault",            StringType(), True),
    StructField("property_crimes_total",         StringType(), True),
    StructField("burglary",                      StringType(), True),
    StructField("larceny",                       StringType(), True),
    StructField("motor_vehicle_theft",           StringType(), True),
])


# --------------------------------------------------------------------------
# Cell 4 — Main Execution
# --------------------------------------------------------------------------
def run_bronze_ingest():
    """
    Orchestrates the full Bronze ingestion pipeline:
      1. Fetch all records from the NYS Open Data API via pagination.
      2. Normalize records against the expected raw schema.
      3. Build a Spark DataFrame with the Bronze schema.
      4. Append audit columns: ingested_at, source_api.
      5. Write to the Bronze Delta table in MainLakehouse (overwrite mode).
    """

    # Step 1: Extract
    raw_records = fetch_all_records(
        base_url   = API_BASE_URL,
        page_limit = PAGE_LIMIT,
        max_records= MAX_RECORDS,
        app_token  = APP_TOKEN,
    )

    if not raw_records:
        raise ValueError("[ERROR] No records returned from API. Aborting ingestion.")

    # Step 2: Normalize
    print(f"[INFO] Normalizing {len(raw_records):,} records...")
    normalized = normalize_records(raw_records, RAW_COLUMNS)

    # Step 3: Create Spark DataFrame
    print("[INFO] Creating Spark DataFrame...")
    df = spark.createDataFrame(normalized, schema=bronze_schema)

    # Step 4: Add audit columns
    ingestion_ts = datetime.now(timezone.utc).isoformat()
    df = (
        df
        .withColumn("ingested_at", current_timestamp())
        .withColumn("source_api",  lit(API_BASE_URL))
    )

    # Step 5: Preview
    print("[INFO] Sample records (top 5):")
    df.show(5, truncate=False)
    print(f"[INFO] Total rows in DataFrame: {df.count():,}")
    print(f"[INFO] Schema:")
    df.printSchema()

    # Step 6: Write to Bronze Delta table (full overwrite each run for idempotency)
    print(f"[INFO] Writing to Delta table: {TARGET_TABLE}...")
    (
        df.write
          .format("delta")
          .mode("overwrite")
          .option("overwriteSchema", "true")
          .saveAsTable(TARGET_TABLE)
    )

    print(f"[SUCCESS] Bronze table '{TARGET_TABLE}' written successfully.")
    print(f"[SUCCESS] Rows written: {df.count():,}")
    print(f"[SUCCESS] Ingestion timestamp: {ingestion_ts}")


# Run
run_bronze_ingest()


# --------------------------------------------------------------------------
# Cell 5 — Validation Checks
# --------------------------------------------------------------------------
print("\n=== BRONZE TABLE VALIDATION ===")

df_validate = spark.sql(f"SELECT * FROM {TARGET_TABLE}")

# Row count
row_count = df_validate.count()
print(f"Total rows in {TARGET_TABLE}: {row_count:,}")

# Null check on key columns
print("\nNull counts per key column:")
from pyspark.sql.functions import col, sum as spark_sum, when

null_counts = df_validate.select([
    spark_sum(when(col(c).isNull(), 1).otherwise(0)).alias(c)
    for c in ["county", "agency", "year", "total_index_crimes"]
]).collect()[0].asDict()

for col_name, null_count in null_counts.items():
    print(f"  {col_name}: {null_count:,} nulls")

# Year range
print("\nYear range:")
spark.sql(f"""
    SELECT MIN(year) AS min_year, MAX(year) AS max_year
    FROM {TARGET_TABLE}
""").show()

# County sample
print("Distinct county count:")
spark.sql(f"SELECT COUNT(DISTINCT county) AS distinct_counties FROM {TARGET_TABLE}").show()

# Latest ingestion timestamp
print("Latest ingested_at:")
spark.sql(f"SELECT MAX(ingested_at) AS latest_ingest FROM {TARGET_TABLE}").show()

print("=== VALIDATION COMPLETE ===")
