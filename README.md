# BRSR Sustainability Comparison

**Cross-company comparison of selected Business Responsibility and Sustainability Report (BRSR) metrics.**

[![Excel](https://img.shields.io/badge/Data-Excel%20workbook-217346)](data/cross_company_comparison.csv) [![Companies](https://img.shields.io/badge/Companies-4-555555)](#companies-included) [![Metrics](https://img.shields.io/badge/Metrics-7-4C6A92)](#comparison-scope)

**Project links:** [Portfolio page](https://waliabidi2315.github.io/BRSR-Sustainability-Comparison/) · [Comparison data CSV](data/cross_company_comparison.csv)


## Overview

This project compares selected environmental and workforce indicators across four financial companies using the cross-company comparison worksheet from the supplied Excel workbook. It is intended as a descriptive comparison of reported metrics—not an overall ESG rating.

**Source selection:** only the `Cross Company Comparision` worksheet was used. The other two comparison sheets were excluded at the user's direction because they were incorrect.

## Companies included

- Bajaj Finserv
- ICICI
- Aditya Birla Capital
- Muthoot Finance

## Comparison scope

| Indicator | Unit as recorded in the workbook |
|---|---|
| Company size | Employees |
| GHG emissions intensity | tCO₂e / INR crore revenue |
| Water intensity | KL / INR crore revenue |
| GHG per employee | tCO₂e / employee |
| Water per employee | KL / employee |
| Employee turnover | % |
| Female wage share proxy | % |

## Data

The selected worksheet has been extracted to a tidy CSV for easy review:

- [Open the cross-company comparison dataset](data/cross_company_comparison.csv)

Blank cells remain blank; no missing values were estimated or filled. Reporting periods are retained as written in the workbook.

## Reading the comparison responsibly

- Reporting periods differ slightly between companies, so strict year-to-year comparability is limited.
- Metrics and units are carried over from the worksheet headers. Verify them against source BRSR disclosures before making decisions.
- Missing entries are not imputed.
- The project does not construct a composite sustainability score or rank companies overall.
- The female wage-share proxy is **not** a direct gender pay gap or equal-pay measure.


## Repository structure

```
BRSR-Sustainability-Comparison/
├── README.md
├── index.html
└── data/
    └── cross_company_comparison.csv
```


## Tools and methods

Excel · Comparative ESG metrics · Data extraction · Cross-company analysis · Data presentation


---

*Portfolio project based exclusively on the selected cross-company comparison worksheet from the supplied Excel workbook.*
