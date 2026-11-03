# Global Residential Property Price Analysis

An end-to-end **Microsoft Excel analytics project** analysing residential property price movements across 58 countries and the Euro area, covering historical market trends from 1948 to September 2023.

The project transforms four raw datasets into a structured analytical model, distinguishes nominal price growth from inflation-adjusted real growth, evaluates housing market cycles, and delivers an interactive Excel dashboard to support international real-estate market comparisons.

## Project Overview

**Business Scenario:** A fictional international real-estate advisory team evaluating residential property markets to understand historical performance, identify market risks, and compare investment environments.

**Business Problem:** Nominal property price growth can create a misleading picture of market performance, particularly in high-inflation economies. Markets showing strong nominal appreciation may experience much weaker growth, or even declining prices, after adjusting for inflation.

The objective was to build an Excel-based analytical solution that separates genuine price growth from inflation-driven increases and identifies markets experiencing the strongest growth, volatility, and price declines.

## Project Objectives

- Consolidate four separate residential property datasets into one structured analytical table.
- Clean and standardise inconsistent data, headers, keys, and numerical values.
- Compare nominal and inflation-adjusted real property price growth.
- Evaluate residential property performance across historical economic cycles.
- Identify markets experiencing declining real prices and significant drawdowns from previous peaks.
- Analyse regional differences, economic group performance, and seasonal trends.
- Develop an interactive Excel dashboard for non-technical business stakeholders.

## Dataset Overview

| Attribute | Details |
|---|---|
| Dataset | Quarterly residential property price data |
| Geographic Coverage | 58 countries + Euro area |
| Total Records | 7,599 market-quarter rows |
| Historical Period | 1948 – September 2023 |
| Index Base | 2010 = 100 |
| Key Measures | Nominal Index, Real Index, Nominal YoY %, Real YoY % |

The original dataset consisted of four separate tables:

1. Nominal Residential Property Price Index
2. Real Residential Property Price Index
3. Nominal Year-over-Year Growth
4. Real Year-over-Year Growth

The original source tables were preserved separately, while the cleaned and integrated data was structured into a single fact table.

## Tools & Technical Skills

- **Microsoft Excel:** Data cleaning, transformation, modelling, analysis, and dashboard development.
- **Excel Formulas:** INDEX/MATCH, VALUE, IFERROR, SUMIFS, AVERAGEIFS, COUNTIFS, MAXIFS, MINIFS, STDEV-related calculations, TEXT, TRIM, EOMONTH, and logical functions.
- **Data Modelling:** Fact and dimension table structure, unique keys, lookup relationships, and time-based attributes.
- **Data Quality:** Missing-value checks, duplicate detection, data type validation, consistency checks, and reconciliation.
- **Data Visualisation:** Interactive charts, KPI cards, conditional formatting, ranking tables, and heat maps.
- **Dashboard Interactivity:** Data Validation dropdowns and named ranges.

*The solution was developed entirely in standard Excel without macros, add-ins, or Power Query.*

## Project Workflow

### 1. Data Profiling & Cleaning

Inspected the four source tables to identify data quality issues and inconsistencies.

Key challenges addressed:

- Blank rows within source tables.
- Numerical values stored as text.
- Inconsistent column headers and extra spaces.
- Inconsistent key construction across datasets.
- Different historical coverage across source tables.
- Country-level data mixed with aggregate observations.
- Extreme nominal growth values in high-inflation markets.
- Different reporting end dates across markets.

A dedicated Cleaning_Log sheet documents the identified issues, evidence, corrections, and Excel techniques used.

### 2. Data Integration & Modelling

Created a standardised composite key using the reporting date and country name.

Used bounded INDEX/MATCH formulas to integrate the four datasets into a consolidated Clean_Data fact table.

Developed a separate Country_Ref dimension table containing country classifications, regional groupings, and economy types.

Additional analytical fields included:

- Year, Quarter, and Decade.
- Nominal and Real Price Indices.
- Nominal and Real YoY Growth.
- Inflation Gap (percentage points).
- Real Quarter-over-Quarter Growth.

### 3. Exploratory & Comparative Analysis

Developed country-level and time-based analytical views to investigate:

- Average real property price growth by country.
- Nominal versus real growth differences.
- Inflation impact across decades.
- Market performance across five historical cycles.
- Regional and advanced-versus-emerging economy comparisons.
- Latest real property price levels relative to historical peaks.
- Real price volatility and the proportion of declining quarters.
- Seasonal patterns in quarterly price movements.

### 4. Interactive Dashboard Development

Built an interactive Excel dashboard designed to help users explore global residential property market performance.

**Dashboard features
