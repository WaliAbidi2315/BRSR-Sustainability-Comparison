# BRSR Sustainability Comparison

**ESG comparison of Indian financial services companies using Business Responsibility and Sustainability Report (BRSR) data, extracted with R and Python.**

[![Report](https://img.shields.io/badge/Report-Word%20(.docx)-2B579A)](reports/BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx) [![Data](https://img.shields.io/badge/Data-Excel%20%2B%20CSV-217346)](data/) [![Companies](https://img.shields.io/badge/Companies-4-555555)](#companies-included) [![Section C rows](https://img.shields.io/badge/Section%20C%20rows-542-4C6A92)](#data)

**Project links:** [Portfolio page](https://waliabidi2315.github.io/BRSR-Sustainability-Comparison/) · [Full report (.docx)](reports/BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx) · [Section C workbook](data/Bajaj_Finserv_BRSR_SectionC_clean.xlsx)

## Overview

SEBI's BRSR is the mandatory ESG disclosure format for India's top listed companies, organised around nine principles (P1–P9) of the National Guidelines on Responsible Business Conduct. This project extracts BRSR disclosures with R and compares financial services companies on environmental and workforce indicators, then goes deep on one company's Section C (principle-wise) disclosures.

Financial services firms are a useful test case: with no manufacturing footprint, their material issues are workforce, governance, customer treatment and financed emissions, and BRSR's manufacturing-style indicators fit them imperfectly.

The analysis is descriptive. It is not an overall ESG rating.

## Report

The full write-up is in [`reports/`](reports/BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx). It covers:

1. A cross-company comparison of headline environmental and workforce metrics (FY2024-25)
2. A principle-by-principle deep dive (P1–P9), FY2024 vs FY2023
3. A multi-year trend view of Aditya Birla Capital
4. A completeness check of FY2025 leadership disclosures
5. Key findings, limitations and data notes

## Companies included

- Bajaj Finserv
- ICICI
- Aditya Birla Capital
- Muthoot Finance

## Data

All files are in [`data/`](data/).

| File | What it holds |
|---|---|
| [`Bajaj_Finserv_BRSR_SectionC_clean.xlsx`](data/Bajaj_Finserv_BRSR_SectionC_clean.xlsx) | Cleaned Section C extraction for Bajaj Finserv (FY2024 with FY2023 comparatives). Five sheets: Read Me, Section C Data, Wide Tables, Qualitative, Raw Text |
| [`bajaj_finserv_sectionC_data.csv`](data/bajaj_finserv_sectionC_data.csv) | Section C Data sheet: 542 rows covering principles P1–P3 and P5–P9, tagged Essential or Leadership, with period, unit and the value exactly as reported |
| [`bajaj_finserv_wide_tables.csv`](data/bajaj_finserv_wide_tables.csv) | Wide tables that did not fit the one-number-per-row layout (well-being measures, hazardous waste reclaimed) |
| [`bajaj_finserv_qualitative.csv`](data/bajaj_finserv_qualitative.csv) | Written answers for all nine principles |

Section C columns: `Principle`, `Q No`, `Indicator`, `Essential/Leadership`, `Item`, `Period`, `Column`, `Col No`, `Value`, `Unit`, `As Reported`. The `As Reported` column keeps exactly what the report printed (NIL, N.A., "8.50 lakh"), so blank `Value` cells mean the report gave no number. Nothing is estimated or filled in. The `Raw Text` sheet in the workbook holds the source text for checking extractions.

## Comparison scope

The cross-company figures (four companies, FY2024-25) are reported in Section 3 of the [report](reports/BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx).

| Indicator | Unit |
|---|---|
| Company size | Employees |
| GHG emissions intensity | tCO₂e / INR crore revenue |
| Water intensity | KL / INR crore revenue |
| GHG per employee | tCO₂e / employee |
| Water per employee | KL / employee |
| Employee turnover | % |
| Female wage share proxy | % |

## Reading the comparison responsibly

- Reporting periods differ slightly between companies (FY2024, FY2025, FY2024-25), so strict year-to-year comparability is limited.
- Group-level and entity-level disclosures are mixed, and intensity denominators are not identical across firms.
- Bajaj Finserv's water per employee (0.000113 KL) is thousands of times below the others and is probably a unit error. Missing values are left blank, not imputed.
- Most figures are self-reported and unaudited.
- The female wage-share proxy is **not** a direct gender pay gap or equal-pay measure.
- The report's deep-dive sections are headed "Bajaj Finance", but the figures (139,000 employees, windmill and insurance operations, "BFS Group") match Bajaj Finserv, and the workbook is named for Bajaj Finserv. The report's own Section 8 flags this for confirmation.
- The project does not build a composite score or rank companies overall.

## Code

Extraction scripts are in [`code/`](code/):

- [`code/R/scrape_brsr_sectionC.R`](code/R/scrape_brsr_sectionC.R)
- [`code/python/brsr_sectionC_clean_scraper.py`](code/python/brsr_sectionC_clean_scraper.py)

## Repository structure

```
BRSR-Sustainability-Comparison/
├── README.md
├── index.html
├── code/
│   ├── R/
│   └── python/
├── data/
│   ├── Bajaj_Finserv_BRSR_SectionC_clean.xlsx
│   ├── bajaj_finserv_sectionC_data.csv
│   ├── bajaj_finserv_wide_tables.csv
│   └── bajaj_finserv_qualitative.csv
└── reports/
    └── BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx
```

## Tools and methods

R · Python · PDF text extraction · Excel · Comparative ESG metrics · Cross-company analysis

---

*Portfolio project by Wali, BS Analytics and Sustainability Studies, TISS Mumbai.*
