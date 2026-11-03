# Project Documentation: Global Residential Property Price Analysis

**Workbook:** `Property_Price_Analysis_Dashboard.xlsx`
**Tool:** Microsoft Excel (no macros, no add-ins, no Power Query)
**Version:** 1.0

---

## 1. Project overview

| Item | Detail |
|---|---|
| Purpose | Show end-to-end Excel analysis: cleaning, modelling, analysis, dashboard, reporting |
| Dataset | Quarterly residential property price data: nominal index, real index, nominal YoY %, real YoY % |
| Coverage | 58 countries plus the Euro area; 7,599 market-quarter rows; 1948 to Sep 2023 |
| Index base | 2010 = 100 |
| Deliverable | One workbook with an interactive dashboard, cleaning log and written insights |

## 2. Business context

**Client (fictional):** an international real-estate advisory team that allocates capital across housing markets.

**Problem:** headline (nominal) price indices overstate gains in high-inflation markets and hide markets where prices are falling in real terms.

**Business question:** which markets delivered genuine price growth, how did they behave in booms and busts, and which are under pressure now?

**Objectives**
1. Turn four messy source tables into one reliable analysis table.
2. Separate real growth from nominal growth and measure the inflation gap.
3. Compare markets and regions across five market cycles.
4. Flag markets furthest below their peak or in falling-price territory.
5. Deliver an interactive dashboard a non-technical manager can use.

**Business questions and answers**

| Question | Answered by | Sheet |
|---|---|---|
| Q1 Which markets grew fastest / slowest in real terms? | Average real YoY %, ranked | Country_Analysis, Dashboard |
| Q2 How much growth does inflation absorb? | Nominal minus real YoY | Country_Analysis, Trend_Analysis |
| Q3 How did regions behave in each cycle? | Region x period heat map | Trend_Analysis, Dashboard |
| Q4 Which markets are weakest now? | Latest vs peak, latest YoY | Country_Analysis, Insights |
| Q5 Is there seasonality? | Average QoQ change by quarter | Trend_Analysis |

**KPIs**

| KPI | Definition |
|---|---|
| Real Price Index | Inflation-adjusted price index, 2010 = 100 |
| Avg Real YoY % | Mean quarterly real year-on-year change over the chosen period |
| Inflation Gap (pp) | Nominal YoY % minus Real YoY % |
| Latest vs Peak % | Latest real index / highest real index - 1 |
| Volatility | Standard deviation of real YoY % |
| % Falling Quarters | Share of quarters with real YoY % below zero |

## 3. Data description

### 3.1 Source sheets (kept untouched as `Raw_*`)

| Raw sheet | Content | Original sheet name |
|---|---|---|
| Raw_Nominal_Index | Nominal price index (2010 = 100) | nominal_price |
| Raw_Real_Index | Real price index (2010 = 100) | real_index |
| Raw_Nominal_YoY | Nominal year-on-year % change | nominal_year |
| Raw_Real_YoY | Real year-on-year % change | real_year |

Each has: Date, Country_Code, Country, a value column, and a Key column.

### 3.2 Data dictionary: `Clean_Data` (fact table)

| Field | Description | Logic |
|---|---|---|
| Date, Country_Code, Country | Quarter-end date and market | Spine from Raw_Real_Index, blanks removed, sorted by country then date |
| Key | Unique row id | `=TEXT(A2,"yyyy-mm-dd")&"_"&C2` |
| Region, Economy_Type | Analyst grouping | Lookup to Country_Ref |
| Year, Quarter, Decade | Time attributes | `YEAR`, `ROUNDUP(MONTH/3)`, `INT(year/10)*10` |
| Nominal_Index, Real_Index | Price index | `VALUE(INDEX/MATCH)` from Raw sheets |
| Nominal_YoY_pct, Real_YoY_pct | YoY change, percent | `VALUE(INDEX/MATCH)` from Raw sheets |
| Inflation_Gap_pp | Nominal minus real YoY | Formula |
| Real_QoQ_pct | Quarter-on-quarter real change | Formula (see 6.4) |
| Real_YoY_Sq | Helper for standard deviation | Real YoY squared |

### 3.3 Dimension table: `Country_Ref`

Country, Country_Code, Region (Europe, Asia-Pacific, Americas, Middle East & Africa, Aggregate), Economy_Type (Advanced, Emerging, Aggregate) and Is_Country (1/0). **Region and Economy_Type are my assumptions**, shown as blue editable inputs.

## 4. Data cleaning

Every issue is logged in `Cleaning_Log` with a live formula as evidence.

| # | Issue | Evidence | Fix | Excel technique |
|---|---|---|---|---|
| 1 | Blank rows inside data tables | 653 rows | Model spine built from non-blank rows only | COUNTBLANK, filter |
| 2 | Numbers stored as text | 22,885 cells | Converted in the lookup | VALUE, ISTEXT, SUMPRODUCT |
| 3 | Untidy / inconsistent headers | Value column named 4 ways; stray spaces | Standard field names in Clean_Data | LEN vs TRIM |
| 4 | Inconsistent key formulas | 2 of 4 sheets lacked TRIM | One standard key everywhere | TEXT, TRIM |
| 5 | Four tables, no common model | 4 sheets | Merged into one fact table | INDEX/MATCH |
| 6 | Lookup not backward compatible | Original used XLOOKUP | Bounded INDEX/MATCH | INDEX/MATCH |
| 7 | Aggregates mixed with countries | World (59 rows), Euro area | World excluded; Euro area flagged | COUNTIF, flag column |
| 8 | Different history by sheet | 88 extra rows in nominal YoY | Model built on the real-series spine | COUNTA |
| 9 | Extreme values (hyperinflation) | 5 nominal YoY values above 100% | Kept and flagged; focus on real growth | COUNTIF |
| 10 | Markets stop at different dates | 6 markets end before latest quarter | Each market measured at its own latest quarter | MAXIFS |
| 11 | Unbalanced panel | 30 markets start after 2000 | Markets-reporting count shown; rankings need 8+ quarters | COUNTIFS |
| 12 | Stray scratch formula | `=7865-7777` in original | Removed | Manual review |

### Validation checks (all PASS)

| Check | Expected | Result |
|---|---|---|
| Rows in Clean_Data = non-blank rows in Raw_Real_Index | 7,599 | 7,599 |
| Duplicate keys | 0 | 0 |
| Numeric measure cells filled | 30,396 | 30,396 |
| Markets without Region mapping | 0 | 0 |
| Markets in Clean_Data | 59 | 59 |
| Average real index in 2010 is about 100 | 100 | 100.0 |

## 5. Data model

```
Raw_Nominal_Index ─┐
Raw_Real_Index   ──┤  Key lookups   ┌────────────┐
Raw_Nominal_YoY  ──┼──────────────► │ Clean_Data │ ◄── Country_Ref
Raw_Real_YoY     ──┘                │  (fact)    │     (dimension)
                                    └─────┬──────┘
                          ┌───────────────┴───────────────┐
                  Country_Analysis                 Trend_Analysis
                          └───────────────┬───────────────┘
                                      Dash_Data ──► Dashboard
                                                    Insights
```

## 6. Key formulas

### 6.1 Standard key (raw sheets)
```excel
=IF(A2="","",TEXT(A2,"yyyy-mm-dd")&"_"&TRIM(C2))
```

### 6.2 Merging the four tables (Clean_Data)
```excel
=IFERROR(VALUE(INDEX(Raw_Real_Index!$D$2:$D$7778,
        MATCH($D2,Raw_Real_Index!$E$2:$E$7778,0))),"")
```
`MATCH` finds the row for the key, `INDEX` returns the value, `VALUE` converts text to a number, and `IFERROR` returns blank if the key is missing.

### 6.3 Region lookup
```excel
=IFERROR(INDEX(Country_Ref!$C$5:$C$63,MATCH(C2,Country_Ref!$A$5:$A$63,0)),"Unmapped")
```

### 6.4 Quarter-on-quarter change
```excel
=IF(AND(C3=C2,A2=EOMONTH(A3,-3),ISNUMBER(K3),ISNUMBER(K2)),(K3/K2-1)*100,"")
```
Only calculated when the previous row is the same market and exactly one quarter earlier.

### 6.5 Country analysis examples
```excel
Average real growth : =AVERAGEIFS(Clean_Data!$M$2:$M$7600,Clean_Data!$C$2:$C$7600,$A5)
Last observation    : =MAXIFS(Clean_Data!$A$2:$A$7600,Clean_Data!$C$2:$C$7600,$A5)
Latest real index   : =SUMIFS(Clean_Data!$K$2:$K$7600,Clean_Data!$C$2:$C$7600,$A5,
                              Clean_Data!$A$2:$A$7600,$F5)
Latest vs peak %    : =(H5/I5-1)*100
Volatility          : =SQRT((SUMIFS(sq_col,country_col,$A5)-n*mean^2)/(n-1))
```
`MAXIFS` and `MINIFS` need Excel 2019 or Microsoft 365.

### 6.6 Dashboard engine (Dash_Data)
- **Named ranges:** `Sel_Country`, `Cmp_Country`, `From_Yr`, `To_Yr`, `Peer_Grp` point to the dashboard dropdown cells.
- **Annual series:** `=IF(AND(year>=From_Yr,year<=To_Yr),IFERROR(AVERAGEIFS(...),NA()),NA())`. `NA()` makes charts skip years with no data.
- **Ranking:** an eligibility flag (right peer group and at least 8 quarters in range), then `LARGE` / `SMALL` plus `INDEX/MATCH`. A tiny `ROW()/1E9` tie-breaker avoids duplicate ranks.
- **Dynamic text:** takeaway sentences use `TEXT()` and `&` so they change with the selection.

## 7. Analysis methodology

| Analysis | Method |
|---|---|
| Market summary | One row per market: first/last date, latest and peak index, drawdown, averages, volatility, rank |
| Annual trend | Simple average across markets reporting that year, with a markets-reporting count |
| Decade view | Average real and nominal growth and inflation gap per decade |
| Market cycles | Pre-GFC boom 2003-07, GFC 2008-11, Recovery 2012-19, COVID boom 2020-21, Rate-hike shock 2022-23 |
| Region / group comparison | Average real YoY by region and economy type for each cycle, shown as a colour-scale heat map |
| Seasonality | Average real QoQ change by calendar quarter since 2000 |

## 8. Dashboard design

**Layout (top to bottom):** title banner; filter panel; five KPI cards; real vs nominal index chart and market comparison chart; year-by-year growth chart and Top 10 chart; Bottom 10 chart and heat map; three live takeaways.

**Controls (orange-bordered cells with dropdowns)**

| Control | Affects |
|---|---|
| Market | KPI cards, first three charts, takeaways |
| Compare with | Comparison chart, takeaway 2 |
| From year / To year | All charts, KPIs and rankings |
| Peer group | Top 10 and Bottom 10 rankings |

**Design choices:** one colour palette (navy, teal, orange); red only for negative values; charts read from `Dash_Data`, so the dashboard sheet holds no heavy formulas; dropdown cells are visibly different from outputs.

## 9. Key findings

1. Nominal prices rose about 6.8% a year versus 2.3% in real terms. The inflation gap was 10.0 pp in the 1970s and 2.3 pp in the 2010s.
2. Average real growth by cycle: +6.8% (pre-GFC), -1.2% (GFC, with 56% of observations falling), +2.2% (recovery), +4.5% (COVID boom), -0.6% (2022-23, with 60% falling).
3. Europe was the most cyclical region: +8.9% before the GFC, -3.6% during it. Asia-Pacific was steadier.
4. In 2022-23 advanced economies averaged -1.6% real growth versus +0.9% for emerging markets.
5. At the latest quarter, 41 of 58 countries show falling real prices and 18 are more than 20% below their peak (Russia is furthest, about -42%).
6. Since 2000, Turkiye has the highest average real growth (8.6% a year) and Romania the lowest (-2.8%).
7. Seasonality is mild: Q3 strongest (0.70% average QoQ), Q4 weakest (0.20%).

**Recommendations (illustrative):** compare markets on real growth; keep a watch list of markets with falling prices well below peak; read volatility next to growth; spread exposure across regions; refresh quarterly.

## 10. Assumptions and limitations

- Region and Economy_Type are analyst assumptions, not source data.
- Averages are unweighted and the panel is unbalanced (markets start at different dates, from Italy in 1948 to Poland in 2011).
- High-inflation markets (for example Turkiye) can distort emerging-market averages.
- Euro area is an aggregate and is excluded from rankings. World is excluded.
- The client scenario is fictional.
- Confirm the dataset source and licence before publishing the raw data.

## 11. How to reproduce this project step by step

1. **Profile:** for each raw sheet check row count, blanks, data types and the list of countries and dates.
2. **Standardise keys:** add `=IF(A2="","",TEXT(A2,"yyyy-mm-dd")&"_"&TRIM(C2))` to each raw sheet.
3. **Build the spine:** copy Date, Country_Code and Country from the real index sheet, remove blanks, sort by country then date.
4. **Merge:** add four `VALUE(INDEX/MATCH)` columns pulling each measure by key.
5. **Add time fields:** Year, Quarter, Decade, inflation gap, QoQ change.
6. **Build Country_Ref** and add Region and Economy_Type through a lookup.
7. **Validate:** row counts, duplicate keys, blanks, base-year check.
8. **Country_Analysis:** one row per market with the formulas in section 6.5.
9. **Trend_Analysis:** annual, decade, cycle and quarter tables with `AVERAGEIFS` and `COUNTIFS`.
10. **Dash_Data:** name the dropdown cells, build the annual series and the ranking block.
11. **Dashboard:** add dropdowns (Data Validation list), KPI cards, charts, conditional formatting and takeaway text.
12. **Document:** Business_Problem, Cleaning_Log and Insights sheets.

## 12. Interview preparation

**Q: How did you merge four tables?** I built a unique key (date + country), standardised it with `TRIM`, and used `INDEX/MATCH` against bounded ranges.

**Q: Why INDEX/MATCH instead of XLOOKUP?** It works in every Excel version and recalculates faster on bounded ranges.

**Q: Why focus on real growth?** Nominal growth includes inflation, which inflates high-inflation markets. Real growth compares markets fairly.

**Q: How did you handle outliers?** I kept them because they are real data, flagged them in the log, and used real growth for comparisons.

**Q: How is the dashboard interactive?** Five dropdown cells (data validation) are named ranges; `AVERAGEIFS`, `MAXIFS` and `INDEX/MATCH` formulas on `Dash_Data` read them, and the charts read from `Dash_Data`.

**Q: How did you check your work?** Six validation checks (row counts, duplicates, missing values, mapping, market count, base-year check), plus comparing totals against the raw data.

**Q: What are the limitations?** Unweighted averages, unbalanced panel, assumed region groups.

## 13. Possible improvements

- Weight averages by market size or housing stock.
- Add Power Query to automate the merge, and a PivotTable with slicers.
- Add affordability data (income or rents) to complement prices.
- Build the same dashboard in Power BI for comparison.

## 14. Version history

| Version | Notes |
|---|---|
| 1.0 | Initial full build: cleaning, model, analysis, dashboard, documentation |
