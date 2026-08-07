# NY Crime Analytics Lakehouse — Fabric Setup Guide
**Daquan Morrison · Data Analyst · Syracuse City School District**  
**Crime Analyst 2 · DCJS CNY Crime Analysis Center**  
Portfolio Project: NY Crime Analytics Lakehouse (Microsoft Fabric)

---

## Prerequisites

- Microsoft Fabric workspace with at least **Contributor** role
- A Fabric capacity assigned to your workspace (Trial, F2, or higher)
- Python familiarity for notebook execution
- NYS Open Data access (public API — no auth required for read-only)

---

## Step 1 — Create MainLakehouse (if not yet created)

1. Open your **Microsoft Fabric workspace** at `https://app.fabric.microsoft.com`.
2. Click **+ New** in the top-left toolbar.
3. Select **Lakehouse** from the items list.
4. Name it exactly: **`MainLakehouse`**
   - This name is referenced in all three PySpark notebooks. If you use a different name, update the `saveAsTable()` calls accordingly, or prefix table references with the lakehouse name.
5. Click **Create**.
6. Your new lakehouse opens with an empty **Tables** section and a **Files** section.
7. Confirm the lakehouse is connected to your workspace capacity.

> **Tip:** If you already have a lakehouse, verify its name in the workspace explorer. Rename it to `MainLakehouse` or update the notebook table references.

---

## Step 2 — Import the Notebooks

Repeat for all three `.py` notebook files:

1. In your **Fabric workspace**, click **+ New → Import notebook** (or use **Upload** from the toolbar).
2. Select the `.py` file from your local machine:
   - `ny_crime_bronze_ingest.py`
   - `ny_crime_silver_transform.py`
   - `ny_crime_gold_tables.py`
3. After import, open each notebook and attach it to **MainLakehouse**:
   - Click the **Lakehouse** icon in the left rail of the notebook editor.
   - Select **Add** → choose **MainLakehouse** → click **Confirm**.
4. This ensures all `spark.table()` and `saveAsTable()` calls resolve to the correct lakehouse.

> **Alternative import method:** Paste the `.py` file contents directly into a new Fabric notebook (New → Notebook → paste code cell by cell using the cell separators marked by `# ----------`).

---

## Step 3 — Running Order (Bronze → Silver → Gold)

Run the notebooks **in sequence**. Each layer depends on the previous.

```
Run 1:  ny_crime_bronze_ingest.py
         ↓
         Creates: bronze_nys_crime

Run 2:  ny_crime_silver_transform.py
         ↓
         Reads:   bronze_nys_crime
         Creates: silver_nys_crime

Run 3:  ny_crime_gold_tables.py
         ↓
         Reads:   silver_nys_crime
         Creates: gold_county_year_crime
                  gold_county_rankings
                  gold_violent_crime_trends
                  gold_property_crime_trends
                  gold_agency_activity
                  gold_forecast_base
```

### Running a notebook:
- Open the notebook in Fabric.
- Click **Run All** (▶▶) from the top toolbar.
- Monitor cell outputs for `[SUCCESS]` messages and validation tables.
- Expected run times:
  - Bronze ingest: ~3–8 minutes (API pagination + Delta write)
  - Silver transform: ~1–2 minutes
  - Gold tables: ~2–5 minutes (6 tables with window functions)

---

## Step 4 — Verify bronze_nys_crime Table Exists

After running `ny_crime_bronze_ingest.py`:

**Option A — Lakehouse Explorer:**
1. Open **MainLakehouse** from the workspace.
2. Expand **Tables** in the left panel.
3. You should see `bronze_nys_crime` listed.
4. Click the table → **Preview data** to see sample rows.

**Option B — SQL Query Editor:**
1. Open **MainLakehouse** → click **New SQL query** (top toolbar).
2. Run:
```sql
SELECT COUNT(*) AS total_rows FROM bronze_nys_crime;

SELECT TOP 5 * FROM bronze_nys_crime;

SELECT MIN(year) AS min_year, MAX(year) AS max_year FROM bronze_nys_crime;
```
3. Expected result: ~100,000+ rows, years ranging 1990–present.

**Option C — In a Fabric Notebook:**
```python
df = spark.table("bronze_nys_crime")
print(f"Row count: {df.count():,}")
df.show(5)
```

---

## Step 5 — Verify All Silver and Gold Tables

After all three notebooks complete:

```sql
-- In the SQL Analytics Endpoint query editor:

SELECT 'bronze_nys_crime'           AS table_name, COUNT(*) AS rows FROM bronze_nys_crime
UNION ALL
SELECT 'silver_nys_crime',                          COUNT(*) FROM silver_nys_crime
UNION ALL
SELECT 'gold_county_year_crime',                    COUNT(*) FROM gold_county_year_crime
UNION ALL
SELECT 'gold_county_rankings',                      COUNT(*) FROM gold_county_rankings
UNION ALL
SELECT 'gold_violent_crime_trends',                 COUNT(*) FROM gold_violent_crime_trends
UNION ALL
SELECT 'gold_property_crime_trends',                COUNT(*) FROM gold_property_crime_trends
UNION ALL
SELECT 'gold_agency_activity',                      COUNT(*) FROM gold_agency_activity
UNION ALL
SELECT 'gold_forecast_base',                        COUNT(*) FROM gold_forecast_base;
```

All 8 tables should return non-zero row counts.

---

## Step 6 — Connect Power BI to the SQL Analytics Endpoint

The **SQL Analytics Endpoint** exposes all Lakehouse Delta tables as a queryable SQL database — no data duplication needed.

1. In your **Fabric workspace**, open **MainLakehouse**.
2. In the top-right, switch from **Lakehouse** view to **SQL analytics endpoint** (toggle in the top-right corner).
3. Copy the **SQL Connection String** from the Properties panel (format: `<workspace>.datawarehouse.fabric.microsoft.com`).
4. Open **Power BI Desktop**.
5. Click **Get Data → More → Microsoft Fabric → Lakehouse**.
   - Or use **Get Data → SQL Server** and paste the connection string directly.
6. Sign in with your Microsoft account (same one used for Fabric).
7. Select **MainLakehouse** → expand **Tables** → check the boxes for all Gold tables:
   - `gold_county_year_crime`
   - `gold_county_rankings`
   - `gold_violent_crime_trends`
   - `gold_property_crime_trends`
   - `gold_agency_activity`
   - `gold_forecast_base`
8. Click **Load** (or **Transform Data** if you want to apply additional Power Query steps).
9. In the Power BI Data model, create relationships if needed:
   - `gold_county_year_crime[county]` ↔ `gold_county_rankings[county]`
   - `gold_county_year_crime[year]` ↔ `gold_county_year_crime[year]` (self-join not needed — use DAX)
   - Create a Date dimension table: `Date = CALENDAR(DATE(1990,1,1), DATE(2030,12,31))`
10. Add the DAX measures from `ny_crime_pbi_measures.md`.

> **Publish to Power BI Service:** File → Publish → Your Workspace. The report will query the Fabric lakehouse live via DirectLake mode.

---

## Step 7 — Dataflow Gen2 Alternative for Bronze Ingest

If you want a **no-code / low-code** alternative to the Bronze PySpark notebook, use **Dataflow Gen2** in Fabric:

### Step-by-step Dataflow Gen2 setup:

1. In your Fabric workspace, click **+ New → Dataflow Gen2**.
2. Name it: `bronze_nys_crime_ingest`.
3. In the Power Query editor, click **Get Data → Web**.
4. Enter the API URL:
   ```
   https://data.ny.gov/resource/ca8h-8gjq.json?$limit=2000&$offset=0
   ```
5. Power Query will parse the JSON array into a table.
6. To handle pagination — in Power Query's **Advanced Editor**, use a recursive function:
   ```powerquery-m
   let
       FetchPage = (offset as number) =>
           let
               url = "https://data.ny.gov/resource/ca8h-8gjq.json?$limit=2000&$offset=" & Number.ToText(offset),
               raw = Web.Contents(url),
               parsed = Json.Document(raw),
               table = Table.FromList(parsed, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
               expanded = Table.ExpandRecordColumn(table, "Column1", 
                   {"county","agency","year","months_reported","total_index_crimes",
                    "violent_crimes_total","murder_non_negl_manslaughter","rape_legacy_definition",
                    "robbery","aggravated_assault","property_crimes_total","burglary",
                    "larceny","motor_vehicle_theft"})
           in
               expanded,
       
       AllPages = List.Generate(
           () => [page = FetchPage(0), offset = 0],
           each Table.RowCount([page]) > 0,
           each [page = FetchPage([offset] + 2000), offset = [offset] + 2000],
           each [page]
       ),
       
       Combined = Table.Combine(AllPages)
   in
       Combined
   ```
7. Add a custom column for `ingested_at`:
   - Add Column → Custom Column → `= DateTime.LocalNow()`
8. Add a custom column for `source_api`:
   - Add Column → Custom Column → `= "https://data.ny.gov/resource/ca8h-8gjq.json"`
9. Set the **Data Destination**:
   - Click the **+** in the query's data destination field.
   - Choose **Lakehouse**.
   - Select **MainLakehouse** → **New table** → name it `bronze_nys_crime`.
   - Set update method to **Replace** (full refresh).
10. Click **Publish**.
11. Schedule a refresh in the Dataflow settings (e.g., monthly, or after data is updated on NYS Open Data).

> **Note:** The PySpark notebook approach is preferred for this use case because it handles 100k+ rows more efficiently and gives you better error handling and logging. Use Dataflow Gen2 if you want a scheduled, no-code refresh that runs automatically.

---

## Step 8 — Refresh Schedule (Optional)

To keep the lakehouse current as NYS publishes new annual data:

1. Open the **Dataflow Gen2** or set up a **Fabric Pipeline**.
2. Schedule the pipeline to run **annually in Q1** (when NYS typically publishes prior-year crime data).
3. Pipeline structure:
   ```
   Trigger (Scheduled / Manual)
       → Run ny_crime_bronze_ingest.py (Notebook Activity)
       → Run ny_crime_silver_transform.py (Notebook Activity)
       → Run ny_crime_gold_tables.py (Notebook Activity)
   ```
4. Add email notifications on pipeline failure.

---

## Resume Language (Add Once Built)

After completing this project, add the following to your resume under your DCJS Crime Analyst 2 role or a **Projects** section:

---

**NY Crime Analytics Lakehouse** | *Microsoft Fabric, PySpark, Delta Lake, Power BI, NYS Open Data*

> Designed and implemented a medallion lakehouse architecture (Bronze/Silver/Gold) in Microsoft Fabric to analyze NYS Index Crime data (1990–present, 100,000+ records). Built PySpark ETL pipelines for paginated API ingestion, type casting, data quality flagging, and analytical aggregations across 8 Delta tables. Created Power BI semantic model with DAX measures including YoY crime change, rolling 3-year averages, and county rankings. Gold layer supports time-series forecasting (Prophet/ARIMA-ready) and is used for CNY regional crime trend reporting.

**Key skills demonstrated:**
- Microsoft Fabric (Lakehouse, Notebook, SQL Analytics Endpoint, Dataflow Gen2)
- PySpark / Spark SQL (Delta Lake, window functions, LAG, RANK, rolling aggregations)
- Data engineering (medallion architecture, schema evolution, data quality flagging)
- Power BI (DirectLake mode, DAX measures, semantic model design)
- Public sector data (NYS Open Data Socrata API, crime analytics)

---

**LinkedIn Project Post (draft):**
> Just built a full medallion lakehouse on Microsoft Fabric using NYS Index Crime data — Bronze API ingest → Silver enrichment → 6 Gold analytical tables → Power BI with DAX measures and forecasting-ready exports. 100k+ records, PySpark window functions, and Delta Lake throughout. Built to support my work at the DCJS CNY Crime Analysis Center. #MicrosoftFabric #DataEngineering #CrimeAnalytics #PowerBI #PySpark

---

## Troubleshooting Common Issues

| Issue | Cause | Fix |
|---|---|---|
| `Table not found` error in Silver notebook | Bronze notebook did not complete successfully | Re-run bronze notebook, check for API errors |
| `AnalysisException: Cannot overwrite` | Delta table locked by another process | Wait 1–2 min and retry, or restart Spark session |
| API returns empty data | Socrata rate limit hit | Add `X-App-Token` header or reduce page size to 1000 |
| Power BI can't find tables | Not connected to SQL Analytics Endpoint | Switch from Lakehouse to SQL Analytics Endpoint view in Fabric |
| `months_reported` null in many rows | Some agencies don't report this field | Expected — data_quality_flag handles this |
| YoY measure returns BLANK | Year slicer has no prior-year data | Normal for 1990 (first year in dataset) |
