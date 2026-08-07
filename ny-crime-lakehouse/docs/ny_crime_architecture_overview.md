# NY Crime Analytics Lakehouse — Architecture Overview
**Daquan Morrison · Data Analyst · Syracuse City School District**  
**Crime Analyst 2 · DCJS CNY Crime Analysis Center**  
Portfolio Project: NY Crime Analytics Lakehouse (Microsoft Fabric)

---

## Project Summary

A medallion lakehouse architecture built in **Microsoft Fabric** to ingest, transform, and analyze New York State Index Crime data from 1990 to present. The system processes 100,000+ records from the NYS Open Data Socrata API through three layered Delta tables (Bronze → Silver → Gold) and exposes analytical Gold tables to Power BI for interactive reporting, county rankings, trend analysis, and time-series forecasting.

**Primary analyst audience:** DCJS CNY Crime Analysis Center, Syracuse City School District leadership, and public safety stakeholders in the Central New York region.

---

## Full Data Flow Diagram

```
╔══════════════════════════════════════════════════════════════════════════════╗
║                     NY CRIME ANALYTICS LAKEHOUSE                            ║
║                      Microsoft Fabric / MainLakehouse                       ║
╚══════════════════════════════════════════════════════════════════════════════╝

  ┌─────────────────────────────────────┐
  │        NYS OPEN DATA API            │
  │  NYS Index Crimes by County/Agency  │
  │  Endpoint: data.ny.gov              │
  │  Format: JSON (Socrata REST API)    │
  │  Records: ~100,000+  (1990–now)     │
  └──────────────────┬──────────────────┘
                     │
                     │  HTTP GET (paginated, limit/offset)
                     │  ~2,000 records/page
                     ▼
  ┌══════════════════════════════════════════════════════╗
  ║  BRONZE LAYER  ·  ny_crime_bronze_ingest.py          ║
  ║  PySpark Notebook · MainLakehouse                    ║
  ╠══════════════════════════════════════════════════════╣
  ║  Table: bronze_nys_crime  (Delta)                    ║
  ║                                                      ║
  ║  • All columns stored as StringType (raw/untyped)    ║
  ║  • No transformations applied                        ║
  ║  • Full overwrite on each run (idempotent)           ║
  ║  • Adds: ingested_at (timestamp), source_api (str)   ║
  ║  • Purpose: immutable source-of-truth audit layer    ║
  ╚══════════════════════════════════════════════════════╝
                     │
                     │  spark.table("bronze_nys_crime")
                     ▼
  ┌══════════════════════════════════════════════════════╗
  ║  SILVER LAYER  ·  ny_crime_silver_transform.py       ║
  ║  PySpark Notebook · MainLakehouse                    ║
  ╠══════════════════════════════════════════════════════╣
  ║  Table: silver_nys_crime  (Delta)                    ║
  ║                                                      ║
  ║  • Cast: year → INT, months_reported → INT           ║
  ║  • Cast: all crime columns → DOUBLE                  ║
  ║  • Standardize: county/agency → strip + Title Case   ║
  ║  • Derive: violent_crime_total (sum of components)   ║
  ║  • Derive: property_crime_total (sum of components)  ║
  ║  • Derive: crime_category (Violent-Heavy /           ║
  ║            Property-Heavy / Mixed)                   ║
  ║  • Derive: decade (FLOOR(year/10)*10)                ║
  ║  • Derive: is_county_total (1/0 flag)                ║
  ║  • Derive: data_quality_flag (COMPLETE/INCOMPLETE)   ║
  ║  • Derive: crime_rate_per_100k (NULL placeholder)    ║
  ╚══════════════════════════════════════════════════════╝
                     │
                     │  spark.table("silver_nys_crime")
                     │  (filtered: is_county_total = 1
                     │   for county-level gold tables)
                     ▼
  ┌══════════════════════════════════════════════════════╗
  ║  GOLD LAYER  ·  ny_crime_gold_tables.py              ║
  ║  PySpark Notebook · MainLakehouse                    ║
  ╠══════════════════════════════════════════════════════╣
  ║                                                      ║
  ║  ┌─────────────────────────────────────────────────┐ ║
  ║  │ gold_county_year_crime                          │ ║
  ║  │ Core county × year analytical table             │ ║
  ║  └─────────────────────────────────────────────────┘ ║
  ║  ┌─────────────────────────────────────────────────┐ ║
  ║  │ gold_county_rankings                            │ ║
  ║  │ Annual county ranking by total crimes           │ ║
  ║  └─────────────────────────────────────────────────┘ ║
  ║  ┌─────────────────────────────────────────────────┐ ║
  ║  │ gold_violent_crime_trends                       │ ║
  ║  │ Violent crime time series + YoY + rolling avg   │ ║
  ║  └─────────────────────────────────────────────────┘ ║
  ║  ┌─────────────────────────────────────────────────┐ ║
  ║  │ gold_property_crime_trends                      │ ║
  ║  │ Property crime time series + YoY + rolling avg  │ ║
  ║  └─────────────────────────────────────────────────┘ ║
  ║  ┌─────────────────────────────────────────────────┐ ║
  ║  │ gold_agency_activity                            │ ║
  ║  │ Agency-level detail with reporting completeness │ ║
  ║  └─────────────────────────────────────────────────┘ ║
  ║  ┌─────────────────────────────────────────────────┐ ║
  ║  │ gold_forecast_base                              │ ║
  ║  │ Clean base for Prophet/ARIMA/PBI forecasting    │ ║
  ║  └─────────────────────────────────────────────────┘ ║
  ╚══════════════════════════════════════════════════════╝
                     │
                     │  SQL Analytics Endpoint
                     │  (DirectLake / Import mode)
                     ▼
  ┌═══════════════════════════════════════╗
  ║  POWER BI SEMANTIC MODEL              ║
  ╠═══════════════════════════════════════╣
  ║  DAX Measures (ny_crime_pbi_measures) ║
  ║  Report Pages:                        ║
  ║  • Statewide Overview                 ║
  ║  • County Deep Dive                   ║
  ║  • CNY / Onondaga Focus               ║
  ║  • Data Quality Dashboard             ║
  ║  • Forecast Visualization             ║
  ╚═══════════════════════════════════════╝
```

---

## Table Schemas

### bronze_nys_crime
Raw ingestion table. All columns stored as strings (no type enforcement).

| Column | Type | Description |
|---|---|---|
| county | STRING | County name as returned by API |
| agency | STRING | Reporting law enforcement agency name |
| year | STRING | Reporting year (raw string) |
| months_reported | STRING | Number of months reported that year (raw string) |
| total_index_crimes | STRING | Total index crimes reported (raw string) |
| violent_crimes_total | STRING | Source-reported violent crime aggregate |
| murder_non_negl_manslaughter | STRING | Murder and non-negligent manslaughter count |
| rape_legacy_definition | STRING | Rape (legacy UCR definition) count |
| robbery | STRING | Robbery count |
| aggravated_assault | STRING | Aggravated assault count |
| property_crimes_total | STRING | Source-reported property crime aggregate |
| burglary | STRING | Burglary count |
| larceny | STRING | Larceny-theft count |
| motor_vehicle_theft | STRING | Motor vehicle theft count |
| ingested_at | TIMESTAMP | UTC timestamp when the row was ingested |
| source_api | STRING | Source API endpoint URL |

---

### silver_nys_crime
Cleaned, typed, and enriched table. Ready for analytical queries.

| Column | Type | Description |
|---|---|---|
| county | STRING | Title-cased, whitespace-stripped county name |
| agency | STRING | Title-cased, whitespace-stripped agency name |
| year | INT | Reporting year as integer |
| months_reported | INT | Months of data reported (1–12) |
| total_index_crimes | DOUBLE | Total index crimes reported |
| violent_crimes_total | DOUBLE | Source-reported violent crime aggregate |
| property_crimes_total | DOUBLE | Source-reported property crime aggregate |
| murder_non_negl_manslaughter | DOUBLE | Murder and non-negligent manslaughter |
| rape_legacy_definition | DOUBLE | Rape (legacy definition) |
| robbery | DOUBLE | Robbery |
| aggravated_assault | DOUBLE | Aggravated assault |
| burglary | DOUBLE | Burglary |
| larceny | DOUBLE | Larceny-theft |
| motor_vehicle_theft | DOUBLE | Motor vehicle theft |
| violent_crime_total | DOUBLE | Derived: sum of murder + rape + robbery + assault |
| property_crime_total | DOUBLE | Derived: sum of burglary + larceny + MVT |
| crime_category | STRING | Violent-Heavy / Property-Heavy / Mixed |
| decade | INT | Decade (1990, 2000, 2010, 2020, ...) |
| is_county_total | INT | 1 if agency = County Total row, 0 otherwise |
| data_quality_flag | STRING | COMPLETE (12 months) / INCOMPLETE (<12 months or null) |
| crime_rate_per_100k | DOUBLE | NULL placeholder — requires population join |
| ingested_at | TIMESTAMP | Passed through from Bronze |
| source_api | STRING | Passed through from Bronze |

---

### gold_county_year_crime
Core analytical table. One row per county per year (county totals only).

| Column | Type | Description |
|---|---|---|
| county | STRING | County name |
| year | INT | Reporting year |
| total_index_crimes | DOUBLE | Total index crimes |
| violent_crime_total | DOUBLE | Derived violent crime total |
| property_crime_total | DOUBLE | Derived property crime total |
| crime_category | STRING | Violent-Heavy / Property-Heavy / Mixed |
| decade | INT | Decade grouping |
| data_quality_flag | STRING | COMPLETE / INCOMPLETE |

---

### gold_county_rankings
County rankings by total crimes for each year. One row per county per year.

| Column | Type | Description |
|---|---|---|
| year | INT | Reporting year |
| county | STRING | County name |
| total_index_crimes | DOUBLE | Total index crimes |
| county_rank | INT | Dense rank within year (1 = highest crimes) |
| violent_crime_total | DOUBLE | Violent crime total |
| property_crime_total | DOUBLE | Property crime total |

---

### gold_violent_crime_trends
Time-series analysis of violent crime components. Includes trend indicators.

| Column | Type | Description |
|---|---|---|
| county | STRING | County name |
| year | INT | Reporting year |
| murder | DOUBLE | Murder and non-negligent manslaughter |
| rape_legacy | DOUBLE | Rape (legacy definition) |
| robbery | DOUBLE | Robbery |
| aggravated_assault | DOUBLE | Aggravated assault |
| violent_crime_total | DOUBLE | Sum of all violent crime components |
| yoy_change | DOUBLE | Year-over-year change (current − prior year); NULL for first year |
| rolling_3yr_avg | DOUBLE | 3-year rolling average of violent_crime_total |

---

### gold_property_crime_trends
Time-series analysis of property crime components. Mirrors violent trends structure.

| Column | Type | Description |
|---|---|---|
| county | STRING | County name |
| year | INT | Reporting year |
| burglary | DOUBLE | Burglary count |
| larceny | DOUBLE | Larceny-theft count |
| motor_vehicle_theft | DOUBLE | Motor vehicle theft count |
| property_crime_total | DOUBLE | Sum of all property crime components |
| yoy_change | DOUBLE | Year-over-year change; NULL for first year per county |
| rolling_3yr_avg | DOUBLE | 3-year rolling average of property_crime_total |

---

### gold_agency_activity
Agency-level detail table. Includes all rows — NOT filtered to county totals.  
Used for municipal/PD-level reporting completeness analysis.

| Column | Type | Description |
|---|---|---|
| county | STRING | County name |
| agency | STRING | Reporting law enforcement agency |
| year | INT | Reporting year |
| total_index_crimes | DOUBLE | Agency-reported total index crimes |
| months_reported | INT | Months of data reported |
| data_quality_flag | STRING | COMPLETE / INCOMPLETE |

---

### gold_forecast_base
Minimal, clean table for forecasting pipelines. One row per county per year.

| Column | Type | Description |
|---|---|---|
| county | STRING | County name |
| year | INT | Reporting year (ordered ASC for time-series models) |
| total_index_crimes | DOUBLE | Total index crimes |
| violent_crime_total | DOUBLE | Violent crime total |
| property_crime_total | DOUBLE | Property crime total |

---

## Business Use Cases by Gold Table

| Gold Table | Primary Use Cases |
|---|---|
| `gold_county_year_crime` | Statewide overview dashboard, decade trend analysis, crime category distribution, data quality monitoring |
| `gold_county_rankings` | "Top 10 Counties" bar chart, rank movement tracking (did Onondaga move up/down?), year-specific snapshots |
| `gold_violent_crime_trends` | Violent crime trend line charts, YoY change cards, 3-year smoothing for policy reporting, component breakdown (what's driving violent crime?) |
| `gold_property_crime_trends` | Property crime trend line charts, burglary vs. larceny vs. MVT breakdown, cyclical pattern detection |
| `gold_agency_activity` | Agency reporting completeness audit, identification of non-reporting or partial-reporting agencies, municipal-level drill-down |
| `gold_forecast_base` | Prophet/ARIMA time-series forecasting, Power BI Analytics pane forecasting, academic research base dataset |

---

## Technology Stack

| Component | Technology | Role |
|---|---|---|
| Platform | Microsoft Fabric | Unified analytics platform (compute + storage + BI) |
| Storage Format | Delta Lake | ACID-compliant columnar storage, time travel, schema enforcement |
| Compute | PySpark (Apache Spark) | Distributed data processing for ETL and window functions |
| Orchestration | Fabric Notebooks | Cell-based ETL pipelines; can be wrapped in Fabric Pipelines |
| Ingestion Alternative | Dataflow Gen2 | No-code/low-code paginated API ingestion |
| Data Source | NYS Open Data (Socrata) | Public API — NYS Index Crimes by County and Agency (1990–present) |
| BI Layer | Microsoft Power BI | Semantic model, DAX measures, interactive dashboards |
| Connectivity | SQL Analytics Endpoint | Exposes Delta tables as SQL for Power BI DirectLake mode |
| Forecasting (future) | Prophet / ARIMA / PBI Analytics | Time-series forecasting from `gold_forecast_base` |
| Architecture Pattern | Medallion (Bronze/Silver/Gold) | Industry-standard lakehouse data quality layering |

---

## Future Enhancements

1. **Population Join**: Connect Census Bureau ACS population data by county and year to compute `crime_rate_per_100k` (currently NULL in Silver).
2. **Prophet Forecasting Notebook**: Add `ny_crime_forecast.py` that reads `gold_forecast_base`, fits a Prophet model per county, and writes predictions to `gold_forecast_predictions`.
3. **Fabric Pipeline**: Wrap all three notebooks in a Fabric Data Pipeline with error handling, retry logic, and email alerts for scheduled annual refresh.
4. **Streaming Ingest**: If near-real-time data becomes available, replace batch Bronze ingest with Fabric Eventstream.
5. **Geospatial Layer**: Add county centroid coordinates to enable Power BI map visuals and choropleth crime density maps.
6. **NIBRS Migration**: Prepare schema evolution plan for transition from UCR (Summary Reporting System) to NIBRS (National Incident-Based Reporting System) as NYS agencies transition.

---

*Architecture designed by Daquan Morrison — Data Analyst, Syracuse City School District / Crime Analyst 2, DCJS CNY Crime Analysis Center*  
*Data source: NYS Open Data — [Index Crimes by County and Agency](https://data.ny.gov/Public-Safety/Index-Crimes-by-County-and-Agency-Beginning-1990/ca8h-8gjq)*
