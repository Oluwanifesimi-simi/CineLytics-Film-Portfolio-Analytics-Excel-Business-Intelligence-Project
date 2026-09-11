# CineLytics-Film-Portfolio-Analytics-Excel-Business-Intelligence-Project
A full-scale data analytics project simulating a real-world entertainment industry engagement — built entirely in Microsoft Excel across 15,000 film records, 26 variables, and 91,219 live formula calculations.

## Table of Contents 

- [📌 Project Overview](#project-overview)
- [🗂️ Repository Structure](#repository-structure)
- [📋 Dataset Specifications](#dataset-specifications)
- [🗃️ Workbook Structure (8 Sheets)](#workbook-structure-8-sheets)
- [🔍 Key Analyses Performed](#key-analysis-performed)
- [💡 Key Findings](#key-findings)
- [🛠️ Excel Techniques Reference](#excel-techniques-reference)
- [📈 Business Questions Answered (20 Total)](#business-questions-answered-20-total)
- [🚀 How to Use This Workbook](#how-to-use-this-workbook)
- [📐 Project Specifications](#project-specifications)
- [👩🏽‍💻 About This Project](#about-this-project)
- [🔗 Connect](#connect)
- [📄 License](#license)


## 📌 Project Overview

This project was developed as a portfolio case study simulating an analytics engagement for CineLytics Inc., a fictional global entertainment consultancy commissioned to evaluate a 15,000-title film portfolio spanning 2000–2024 across theatrical and streaming distribution channels.

The business problem: the client's executive team had no clear framework for identifying which genres, studios, platforms, directors, and budget ranges were actually generating value — and which were quietly destroying it.

The deliverable: a fully structured Microsoft Excel workbook containing a data cleaning audit, 20 answered business questions, five analytical frameworks, a statistical outlier detection engine, a 3-year revenue forecast, a platform-genre strategy matrix, and an executive dashboard — all powered by live formulas with zero hardcoded answers.


## 🗂️ Repository Structure

### [📊CineLytics_Film_Analytics_Dataset - Original.xlsx](https://github.com/Oluwanifesimi-simi/CineLytics-Film-Portfolio-Analytics-Excel-Business-Intelligence-Project/blob/main/CineLytics_Film_Analytics_Dataset%20-%20Original.xlsx)

| 📁 screenshots | Purpose |
|-----------------|-------------|
| 01_executive_dashboard.png | Chart Visualization |
| 02_platform_genre_matrix.png | Best performing streaming platforms by genre |
| 03_outlier_detection_engine.png | Statistical outliers- blockbusters & flops |
| 04_studio_efficiency_scorecard.png | ROI efficiency for each company |
| 05_revenue_forecast_chart.png | 3-year industry revenue forecast |
| 06_Country_profitability_table.png | Countries with the most profitable films |

| 📁 docs/Paths | Purpose |
|-----------------|-----------------|
| `CineLytics_Film_Analytics_Dataset - Original.xlsx` | All 20 questions with techniques · Column-level documentation · Key findings write-up · Techniques and formulas used
| `CineLytics-Film-Portfolio-Analytics-Excel-Business-Intelligence-Project/README.md` | Project narrative, table of contents, Structures, insights, and visual documentation. |

## 📋 Dataset Specifications
| Attribute |	Detail |
|-----------|----------|
| Total Records	| 15,000 films |
| Time Period	| 2000 – 2024 |
| Total Columns |	26 variables |
| Live Formulas	| 91,219 calculations |
| Formula Errors	| 0 |
| Workbook Sheets	| 9 |
### Columns included
Movie ID · Movie Title · Genre · Production Budget ($M) · Box Office Revenue ($M) · Profit ($M) · Release Year · Release Month · Release Season · Runtime (min) · IMDb Rating · Rotten Tomatoes Score (%) · Director · Lead Actor/Actress · Production Company · Country · Streaming Platform · Number of Awards · Marketing Spend ($M) · Opening Weekend Sales ($M) · Audience Age Group · Viewer Rating (1–10) · Subscription Growth (%) · ROI (%) · Revenue per $ Budget · Profitability Tier

## 🗃️ Workbook Structure (8 Sheets)
**1. RAW DATA:**
The full 15,000-row dataset with frozen headers, zebra-row formatting, and five additional derived columns added during analysis:
| Column |	Name |	Purpose |
|--------------|-----------|------------|
| AA  |	Franchise Flag        |	IFERROR/INDEX-MATCH lookup against director reference list |
| AB	| Revenue Outlier Flag	| ±2σ rule: "Blockbuster" / "Flop" / blank |
| AC	| Profit Outlier Flag	  | ±2σ rule: "High Profit" / "Deep Loss" / blank |
| AD	| ROI Outlier Flag      | ±2σ rule: "High ROI" / "Low ROI" / blank |
| AE	| Master Outlier Label	| Priority-ranked composite label |
| AF	| Severity Score (0–3)	| Count of metrics that triggered a flag per film |

**2. DATA DICTIONARY:**
Full column-level documentation covering data type, description, value range, example value, and analytical notes for all 26 original columns.

**3. PIVOT ANALYSIS:**
contains 20 full answered business questions with summaries, working formulars and pivot tables.

**4. BUSINESS QUESTIONS:**
20 structured analyst questions, each documented with:

- Analytical category
- Columns to use
- Excel technique required
- Difficulty level (Beginner / Intermediate / Advanced)
- Expected output format

**5. EXCEL SKILLS GUIDE:**
22 Excel tools mapped to specific project tasks — from Pivot Tables through FORECAST.ETS, DAX Measures, and Dashboard Design — with real application examples from this dataset.

**6. INSIGHTS GUIDE:**
15 annotated findings structured across three tiers:

- 5 Beginner insights
- 5 Intermediate insights
- 5 Advanced insights

Each entry includes the finding, why it matters, how to find it, and the key metric to report.

**7. FRANCHISE DIRECTORS**
A lookup table that contains franchise directors information reference list

**8. DASHBOARD TEMPLATE:**
Executive dashboard shell with:

- 7 chart placeholder zones
- Live KPI formula cards (Total Films, Total Revenue, Total Profit, Avg ROI, Avg IMDb, Blockbusters, Loss Films)
- Full setup instructions for Pivot Chart placement, Slicer connection, and Sparkline addition

## 🔍 Key Analyses Performed

**1. Data Cleaning & Audit**
- Structural null identification using Go To Special → Blanks
- Hidden whitespace detection via SUMPRODUCT(--(LEN(TRIM(range))=0))
- Data type validation across all 26 columns
- Duplicate check on Movie ID primary key
- Outlier flagging using ±2σ — preserved for analysis, not deleted
  
**2. Genre & Profitability Analysis**
- 15-genre performance matrix across 8 metrics
- Budget bin analysis across 7 tiers ($0–25M through $200M+)
- ROI sweet spot identification: $40–75M budget range
- Profitability tier segmentation: Loss / Low / Medium / High
  
![Country profitability table](https://github.com/Oluwanifesimi-simi/CineLytics-Film-Portfolio-Analytics-Excel-Business-Intelligence-Project/blob/main/Screenshot1/06_Country_profitability_table.png.png)

**3. Statistical Outlier Detection**

Upper Outlier = Value > AVERAGE + 2 × STDEV

Lower Outlier = Value < AVERAGE - 2 × STDEV

*Applied independently across Revenue, Profit, and ROI. Results:*

- 800 Revenue Blockbusters (5.3% of portfolio)
- 79 Mega Blockbusters — flagged on all three metrics simultaneously
- 90.6% of films classified as Normal — consistent with ±2σ statistical expectation

![Statistical Outliers](https://github.com/Oluwanifesimi-simi/CineLytics-Film-Portfolio-Analytics-Excel-Business-Intelligence-Project/blob/main/Screenshot1/03_outlier_detection_engine.png.png)

**4. Platform-Genre Strategy Matrix**
- 750 live AVERAGEIFS/COUNTIFS cells
- Independent colour-scale conditional formatting per matrix
- Dynamic best/worst genre lookup per platform using INDEX/MATCH

![Platform genre strategy](https://github.com/Oluwanifesimi-simi/CineLytics-Film-Portfolio-Analytics-Excel-Business-Intelligence-Project/blob/main/Screenshot1/02_platform_genre_matrix.png.png)
  
**5. Studio Efficiency Scorecard**

ROI Consistency Score formula:

```excel
=IFERROR(ROUND(100 - (StdDev_ROI / Avg_ROI * 100), 1), 0)
```

*Efficiency label using nested IF/AND:*
```
excel
=IF(AND(Avg_ROI > IndustryAvgROI,
        Avg_Budget < AVERAGE(All_Budgets),
        StdDev_ROI < 50),
   "✅ Highly Efficient",
IF(AND(Avg_ROI > IndustryAvgROI, StdDev_ROI < 50),
   "⚡ Profitable but High Spend",
IF(AND(Avg_ROI > IndustryAvgROI, StdDev_ROI >= 50),
   "⚠️ Profitable but Inconsistent",
IF(AND(Avg_ROI <= IndustryAvgROI,
        Avg_Budget < AVERAGE(All_Budgets)),
   "🔻 Underspending but Underperforming",
   "❌ Inefficient"))))
```

![Studio efficiency matrix](https://github.com/Oluwanifesimi-simi/CineLytics-Film-Portfolio-Analytics-Excel-Business-Intelligence-Project/blob/main/Screenshot1/04_studio_efficiency_scorecard.png.png)

**6. Franchise Director Analysis**
```
excel
=IFERROR(INDEX(FranchiseList_Flag,
         MATCH(M2, FranchiseList_Director, 0)),
         "Non-Franchise")
```
*Finding:* No meaningful portfolio-wide revenue premium — but franchise directors appear disproportionately among Severity Score = 3 (Mega Blockbuster) films.

**7. Revenue Forecasting**
```
excel
=FORECAST.ETS(target_year, revenue_history, year_history, 1, 1)
=FORECAST.ETS.CONFINT(target_year, revenue_history, year_history, 0.95, 1, 1)
```
3-year projection (2025–2027) with 95% confidence intervals. Confidence bands reported as widening at year 3 — documented honestly rather than smoothed.

![A 3-year revenue forecast](https://github.com/Oluwanifesimi-simi/CineLytics-Film-Portfolio-Analytics-Excel-Business-Intelligence-Project/blob/main/Screenshot1/05_revenue_forecast_chart.png.png)

**8. Critical-to-Commercial Success Ratio**
```
excel
Critical Score  = ROUND(((Avg_IMDb / 10) * 100 + Avg_RT) / 2, 1)
Commercial Score = ROUND((Avg_ROI / MAX(All_ROI)) * 100, 1)
C2C Ratio       = ROUND(Critical_Score / Commercial_Score, 2)
```
Plotted on a quadrant scatter chart: Blockbuster Darlings · Critical Darlings · Commercial Hits · Underperformers.

## 💡 Key Findings
| S/N | Finding |	Implication|
|-----------|-----------|----------|
| 1	| ROI peaks at $40–75M budget; declines above $150M	| Challenges the tentpole investment strategy |
| 2	| Horror averages 4.5× revenue multiple on sub-$20M budgets	| Most capital-efficient genre in the portfolio |
| 3	| Theatrical releases outperform all streaming platforms on avg profit	| Complicates streaming-first pivot strategy |
| 4	| Documentary ranks last on profit across 9 of 10 platforms	| Cannot be justified on commercial grounds alone |
| 5	| Opening weekend > 1.8× portfolio average predicts top-quartile revenue	| Unused early-warning signal for distribution decisions |
| 6	| Animation leads streaming profitability across most platforms	| Strongest genre-platform fit for streaming investment |
| 7	| Franchise director premium disappears at portfolio level	| Blanket premium fees not supported by average-level data |
| 8	| 79 Mega Blockbusters carry Severity Score = 3 across all outlier metrics	| Extreme outliers follow distinct genre and budget patterns |

## 🛠️ Excel Techniques Reference
| Category	| Techniques |
|------------|-----------|
| Lookup & Reference	| XLOOKUP · VLOOKUP · INDEX-MATCH · Named Ranges |
| Aggregation	| AVERAGEIF · AVERAGEIFS · SUMIF · SUMIFS · COUNTIF · COUNTIFS |
| Statistical	| CORREL() · STDEV(IF()) · RANK() · LARGE() · SMALL() · SUMPRODUCT() |
| Forecasting	| FORECAST.ETS · FORECAST.ETS.CONFINT · FORECAST.LINEAR |
| Logic	| Nested IF · IF/AND · IFERROR · Array Formulas |
| Data Cleaning	| TRIM() · LEN() · COUNTBLANK() · Go To Special |
| Visualization	| Pivot Charts · Combo Charts · Scatter + Trendline · Benchmark Line · Heatmap |
| Interactivity	| Pivot Tables · Slicers · Conditional Formatting · Data Validation |
| Transformation | Power Query · Histogram & Bin Analysis · Dashboard Design |

## 📈 Business Questions Answered (20 Total)
| S/N |	Question	| Difficulty |
|-----------|-----------|-------------|
| 1	| Which genres generate the highest average profit and ROI?	| Beginner |
| 2	| Does marketing spend significantly affect box office revenue?	| Intermediate |
| 3	| Which directors and actors consistently deliver high ROI?	| Intermediate |
| 4	| What seasonal trends exist in releases and revenue?	| Beginner |
| 5	| Which streaming platform performs best by genre?	| Intermediate |
| 6	| What is the relationship between IMDb rating and revenue?	| Intermediate |
| 7	| How has average production budget changed over the years?	| Intermediate |
| 8	| Which production companies have the highest market share?	| Beginner |
| 9	| Is there a budget sweet spot that maximizes ROI?	 | Advanced |
| 10 | Which countries produce the most profitable films?	| Intermediate |
| 11	| Do higher-rated films win more awards?	| Intermediate |
| 12	| What percentage of films in each genre are profitable?	| Beginner |
| 13	| How does opening weekend predict total revenue?	| Advanced |
| 14	| Which audience age groups drive the most revenue per genre?	| Intermediate |
| 15	| Do streaming films generate higher subscription growth when critically acclaimed?	| Advanced |
| 16	| What is the ROI efficiency of each production company?	| Intermediate |
| 17	| How do franchise directors compare to non-franchise directors in revenue?	| Advanced |
| 18	| Which genre has the best critical-to-commercial success ratio?	| Advanced |
| 19	| Forecast total industry revenue for the next 3 years	| Advanced |
| 20	| Identify statistical outliers — blockbusters and flops | Advanced |

## 🚀 How to Use This Workbook

*Step 1 — Start with DATA DICTIONARY* 
Read the column definitions before touching any formula. Understanding what each column measures and its valid range prevents misinterpretation before it happens.

*Step 2 — Run the data audit*
Go to RAW DATA → select the full range → Home → Find & Select → Go To Special → Blanks. Confirm only column W (Subscription Growth) returns blanks, and only for Theatrical Only platform rows.

*Step 3 — Work through the BUSINESS QUESTIONS sheet*
Each question includes the columns to use, the technique required, and the expected output. Build your own analysis first before checking the pre-built answers in SUMMARY STATS.

Step 4 — Build your own Pivot Tables
Select the RAW DATA range → Insert → PivotTable → Add to Data Model. Work through Questions 1, 4, 5, 8, and 12 using Pivot Tables alone before moving to formula-based analysis.

*Step 5 — Explore the advanced sheets*
Once comfortable with Pivot outputs, open OUTLIER ANALYSIS and PLATFORM-GENRE MATRIX to study how the formula architecture is structured — particularly the AVERAGEIFS pattern inside the 10×15 matrix and the IFERROR/INDEX/MATCH franchise lookup.

*Step 6 — Build the Dashboard*
Use the DASHBOARD TEMPLATE sheet as your canvas. Create Pivot Charts in SUMMARY STATS, then move them to the dashboard using Chart Tools → Move Chart → Object In → DASHBOARD TEMPLATE. Add Slicers via PivotChart Analyze → Insert Slicer.

## 📐 Project Specifications
| Specification	| Detail |
|-----------|------------|
| Tool	| Microsoft Excel (365 compatible) |
| Dataset Size	| 15,000 rows × 26 columns |
| Derived Columns Added	| 6 (Franchise Flag, 3 Outlier Flags, Master Label, Severity Score) |
| Conditional Formatting Rules	| 18 rules across 4 sheets |
| Named Ranges	| 2 (FranchiseList_Director, FranchiseList_Flag) |
| Charts	| 6 embedded charts across 3 sheets |
| Total Formula Calculations	| 91,219 |
| Formula Error Rate	| 0% |
| Time to Complete	| 4 weeks |

## 👩🏽‍💻 About This Project

This project was built as a portfolio piece to demonstrate end-to-end data analytics capability in Microsoft Excel — from raw data ingestion and cleaning through statistical analysis, business intelligence, and executive-level reporting.

It was designed to simulate the kind of analytical engagement a Financial Data Analyst or Business Intelligence Analyst would be expected to deliver for a media or entertainment industry client — with realistic business questions, honest findings (including ones that contradicted the expected narrative), and outputs built for a non-technical executive audience.

**Author:** Oyinlola Oluwanifesimi Oladeji

**Role:** Financial Data Analyst & Accountant

**Location:** Ado, Ekiti State, Nigeria

**Education:** M.Sc. Financial Engineering — WorldQuant University (expected 2027)

**Open to:** Financial Data Analyst · Business Intelligence Analyst · Accounting & Finance Analyst roles (remote and international)

## 🔗 Connect

[LinkedIn](https://www.linkedin.com/in/oyinlola-oladeji-430108294/)

[Email](http://mail.google.com/oluwanifesimi988@gmail.com)

## 📄 License

This project is shared for portfolio and educational purposes.
Dataset is synthetically generated — no real film industry data or proprietary information is contained in this repository.

*If this project is useful to you, please consider leaving a ⭐ — it helps others find it.*

