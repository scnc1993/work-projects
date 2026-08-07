# NY Crime Analytics — Power BI DAX Measures
**Daquan Morrison · Data Analyst · Syracuse City School District**  
**Crime Analyst 2 · DCJS CNY Crime Analysis Center**  
Portfolio Project: NY Crime Analytics Lakehouse (Microsoft Fabric)

---

## How to Add These Measures

1. Open Power BI Desktop (or Power BI Service Report Editor).
2. Connect to the **SQL Analytics Endpoint** of `MainLakehouse` (see setup guide).
3. In the **Data** pane, right-click on the target table → **New Measure**.
4. Paste the DAX expression.
5. Set the **Format** as specified in each measure block.

> All measures should be added to a dedicated hidden table called `_Measures` for clean model organization. Create a calculated table with `_Measures = {0}` and add all measures there.

---

## Table Reference Map

| Gold Table | Use In These Measures |
|---|---|
| `gold_county_year_crime` | Total Crimes, Violent Rate, Property Rate, Top County, Onondaga |
| `gold_violent_crime_trends` | YoY Change, 3-Year Moving Avg (violent) |
| `gold_property_crime_trends` | YoY Change, 3-Year Moving Avg (property) |
| `gold_county_rankings` | Ranking visuals |
| `gold_agency_activity` | Agency-level filters |
| `gold_forecast_base` | Forecasting visuals |

---

## Core Measures

---

### 1. Total Crimes

```dax
Total Crimes = 
SUM( gold_county_year_crime[total_index_crimes] )
```

**Format:** Whole number, comma separator  
**Chart Type:** Card visual (KPI), Clustered Bar Chart (by county or year)  
**Notes:** Place in a card at the top of the report page as a headline KPI. Use with a Year slicer for dynamic filtering.

---

### 2. Violent Crime Rate

```dax
Violent Crime Rate = 
DIVIDE(
    SUM( gold_county_year_crime[violent_crime_total] ),
    SUM( gold_county_year_crime[total_index_crimes] ),
    0
)
```

**Format:** Percentage, 1 decimal place (e.g., `24.3%`)  
**Chart Type:** Clustered Bar Chart (county comparison), Line Chart (trend over time)  
**Notes:** Use DIVIDE to safely handle division by zero. Pair with a county slicer to show Onondaga vs. statewide average. Bar charts preferred — avoid pie charts.

---

### 3. Property Crime Rate

```dax
Property Crime Rate = 
DIVIDE(
    SUM( gold_county_year_crime[property_crime_total] ),
    SUM( gold_county_year_crime[total_index_crimes] ),
    0
)
```

**Format:** Percentage, 1 decimal place (e.g., `75.7%`)  
**Chart Type:** Clustered Bar Chart (county comparison), Stacked Bar Chart (alongside Violent Crime Rate)  
**Notes:** Violent Crime Rate + Property Crime Rate will not always sum to 100% because some totals include rounding differences from component sums vs. the source aggregate. Communicate this to stakeholders.

---

### 4. YoY Crime Change %

```dax
YoY Crime Change % = 
VAR CurrentYear = 
    CALCULATE(
        SUM( gold_county_year_crime[total_index_crimes] )
    )
VAR PriorYear = 
    CALCULATE(
        SUM( gold_county_year_crime[total_index_crimes] ),
        DATEADD( 'Date'[Date], -1, YEAR )
    )
RETURN
    DIVIDE(
        CurrentYear - PriorYear,
        PriorYear,
        BLANK()
    )
```

**Format:** Percentage, 1 decimal place. Conditional formatting: red for positive (more crime), green for negative (less crime).  
**Chart Type:** Card visual (current year context), Line Chart (trend), Matrix (county × year)  
**Notes:** Requires a `Date` dimension table in your model with a continuous date range. If you don't have one yet, create a calculated table:

```dax
Date = CALENDAR( DATE(1990,1,1), DATE(2030,12,31) )
```

Then relate `Date[Year]` to `gold_county_year_crime[year]` (mark as date table or use `YEAR()` extraction). Alternatively, a simpler version using an integer year column:

```dax
YoY Crime Change % (Simple) = 
VAR SelectedYear = SELECTEDVALUE( gold_county_year_crime[year] )
VAR CurrentTotal = 
    CALCULATE(
        SUM( gold_county_year_crime[total_index_crimes] ),
        gold_county_year_crime[year] = SelectedYear
    )
VAR PriorTotal = 
    CALCULATE(
        SUM( gold_county_year_crime[total_index_crimes] ),
        gold_county_year_crime[year] = SelectedYear - 1
    )
RETURN
    DIVIDE( CurrentTotal - PriorTotal, PriorTotal, BLANK() )
```

---

### 5. Top County by Crime

```dax
Top County = 
CALCULATE(
    SELECTEDVALUE( gold_county_year_crime[county] ),
    TOPN(
        1,
        VALUES( gold_county_year_crime[county] ),
        [Total Crimes],
        DESC
    )
)
```

**Format:** Text (no formatting needed)  
**Chart Type:** Card visual  
**Notes:** Returns the county name with the highest total crimes in the current filter context (year slicer, etc.). Returns BLANK if context is ambiguous. Pair with a year slicer so it dynamically updates.

---

### 6. Onondaga Crimes

```dax
Onondaga Crimes = 
CALCULATE(
    [Total Crimes],
    gold_county_year_crime[county] = "Onondaga"
)
```

**Format:** Whole number, comma separator  
**Chart Type:** Card visual, Line Chart (trend)  
**Notes:** Hardcoded filter for Onondaga County — relevant for the Syracuse/CNY Crime Analysis Center context. Place prominently on the CNY-focused report page. Pair with a year slicer and trend line. Use conditional formatting to highlight years above the 10-year average.

---

### 7. Onondaga Violent Crime Rate

```dax
Onondaga Violent Crime Rate = 
CALCULATE(
    [Violent Crime Rate],
    gold_county_year_crime[county] = "Onondaga"
)
```

**Format:** Percentage, 1 decimal place  
**Chart Type:** Card visual, comparison bar alongside statewide rate

---

### 8. 3-Year Moving Average (Violent Crime)

```dax
3-Year Moving Avg Violent = 
VAR SelectedYear = SELECTEDVALUE( gold_violent_crime_trends[year] )
VAR SelectedCounty = SELECTEDVALUE( gold_violent_crime_trends[county] )
RETURN
    CALCULATE(
        AVERAGE( gold_violent_crime_trends[violent_crime_total] ),
        gold_violent_crime_trends[county] = SelectedCounty,
        gold_violent_crime_trends[year] >= SelectedYear - 2,
        gold_violent_crime_trends[year] <= SelectedYear
    )
```

**Format:** Decimal number, 1 decimal place  
**Chart Type:** Line Chart (overlay on bar chart showing annual totals)  
**Notes:** The pre-computed `rolling_3yr_avg` column in `gold_violent_crime_trends` is the recommended approach for performance — use that column directly in visuals. This measure provides dynamic context-aware computation when you need slicing flexibility.

---

### 9. 3-Year Moving Average (Property Crime)

```dax
3-Year Moving Avg Property = 
VAR SelectedYear = SELECTEDVALUE( gold_property_crime_trends[year] )
VAR SelectedCounty = SELECTEDVALUE( gold_property_crime_trends[county] )
RETURN
    CALCULATE(
        AVERAGE( gold_property_crime_trends[property_crime_total] ),
        gold_property_crime_trends[county] = SelectedCounty,
        gold_property_crime_trends[year] >= SelectedYear - 2,
        gold_property_crime_trends[year] <= SelectedYear
    )
```

**Format:** Decimal number, 1 decimal place  
**Chart Type:** Line Chart (overlay on annual property crime bar chart)

---

### 10. Statewide Average (Total Crimes per County)

```dax
Statewide Avg Crimes Per County = 
AVERAGEX(
    VALUES( gold_county_year_crime[county] ),
    [Total Crimes]
)
```

**Format:** Whole number, comma separator  
**Chart Type:** Reference line on bar charts comparing counties  
**Notes:** Use as a reference line in bar charts via Analytics pane. Enables instant visual identification of above/below average counties.

---

### 11. Data Completeness Rate

```dax
Data Completeness Rate = 
DIVIDE(
    COUNTROWS(
        FILTER(
            gold_county_year_crime,
            gold_county_year_crime[data_quality_flag] = "COMPLETE"
        )
    ),
    COUNTROWS( gold_county_year_crime ),
    0
)
```

**Format:** Percentage, 1 decimal place  
**Chart Type:** Card visual on a data quality/governance page  
**Notes:** Essential for reporting data reliability to stakeholders. A rate below 90% in a given year/county warrants investigation.

---

### 12. COMPLETE Records Count

```dax
Complete Records = 
CALCULATE(
    COUNTROWS( gold_county_year_crime ),
    gold_county_year_crime[data_quality_flag] = "COMPLETE"
)
```

**Format:** Whole number  
**Chart Type:** Card visual, paired with Data Completeness Rate

---

## Report Page Design Recommendations

> No pie charts. Use bar charts, line charts, and card visuals throughout.

### Page 1 — Statewide Overview
- Card: Total Crimes | Violent Crime Rate | Property Crime Rate | Top County
- Clustered Bar: Top 10 Counties by Total Crimes (filtered by year slicer)
- Line Chart: Statewide Total Crimes trend (1990–present)
- Slicer: Year (or Year range)

### Page 2 — County Deep Dive
- Slicer: County (single select)
- Line Chart: Total Crimes trend with 3-year rolling avg overlay
- Clustered Bar: Violent vs. Property breakdown by year
- Card: YoY Crime Change %

### Page 3 — CNY / Onondaga Focus
- Card: Onondaga Crimes | Onondaga Violent Crime Rate | YoY Change
- Line Chart: Onondaga vs. statewide average trend
- Clustered Bar: Onondaga crime component breakdown (murder, rape, robbery, assault, burglary, larceny, MVT)
- Reference line: Statewide average on bar charts

### Page 4 — Data Quality
- Card: Data Completeness Rate | Complete Records
- Bar: Completeness rate by county (highlight incomplete agencies)
- Matrix: County × Year with data_quality_flag conditional formatting

### Page 5 — Forecast
- Connect `gold_forecast_base` to Analytics pane for built-in Power BI forecasting
- Or export to Python/R visual with Prophet model
- Line Chart: Historical + forecast horizon for Onondaga and top 5 counties
