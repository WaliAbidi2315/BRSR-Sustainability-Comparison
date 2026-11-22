<div align="center">

[![Python](https://img.shields.io/badge/Python-scraping-3776AB?logo=python&logoColor=white)](https://www.python.org/) [![R](https://img.shields.io/badge/R-analysis-276DC3?logo=r&logoColor=white)](https://www.r-project.org/) [![Pandas](https://img.shields.io/badge/Pandas-data-150458?logo=pandas&logoColor=white)](#) [![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel%20outputs-217346?logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io/) [![pdftools](https://img.shields.io/badge/R%20pdftools-PDF%20extraction-276DC3?logo=r&logoColor=white)](https://docs.ropensci.org/pdftools/) [![Poppler](https://img.shields.io/badge/Poppler-pdftotext-555555)](https://poppler.freedesktop.org/)

</div>

# BRSR Sustainability Comparison

A data-driven review of **Business Responsibility and Sustainability Report (BRSR)** disclosures from Indian financial-services companies. The project first compares four companies on a common set of headline metrics, then goes deep into one company's full Section C disclosures.

**Portfolio:** [Open the project page](https://waliabidi2315.github.io/BRSR-Sustainability-Comparison/)  
**Full report:** [Download the Word report](reports/BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx)

## Project scope

BRSR is SEBI's mandatory ESG disclosure format for India's top listed companies. It is organised around nine principles (P1–P9) of the National Guidelines on Responsible Business Conduct, and separates **Essential** indicators, which every company must report, from voluntary **Leadership** indicators. Financial-services firms are a useful test case: with no manufacturing footprint, their material issues are workforce, governance, customer treatment and financed emissions, and BRSR's manufacturing-style indicators fit them imperfectly.\n\nThe work is descriptive. It is **not an overall ESG rating or investment recommendation**.

The project has two parts:
- **Cross-company comparison:** four financial-services companies, using a common set of headline environmental and workforce measures.
- **Company deep dive:** structured extraction and qualitative analysis of Bajaj Finserv's Section C, covering the NGRBC principles, year comparisons, wider tables and narrative answers.

\n## Project parts\n\n| Part | Question it answers | Scope |\n|---|---|---|\n| **Part 1: Cross-company comparison** | How do comparable companies look on the same headline measures? | Four companies, FY2024–25, seven environmental and workforce indicators |\n| **Part 2: Bajaj Finserv deep dive** | What does one company's full BRSR disclosure contain, and how good is it? | Section C across the NGRBC principles, FY2024 vs FY2023, quantitative and qualitative answers |\n\n## Companies in the comparison

- Bajaj Finserv
- ICICI
- Aditya Birla Capital
- Muthoot Finance

## Headline comparison

| Company | Reporting period | Employees | GHG intensity (tCO₂e / INR Cr revenue) | Water intensity (KL / INR Cr revenue) | Employee turnover | Female wage-share proxy |
|---|---|---:|---:|---:|---:|---:|
| Bajaj Finserv | 2024–2025 | 108,742 | — | 12.29 | 18% | 6.48% |
| ICICI | FY2025 | 20,035 | 0.10 | 2.24 | 14% | 26% |
| Aditya Birla Capital | 2024–2025 | 60,187 | 2.00 | 16.68 | 38.87% | 21% |
| Muthoot Finance | 2024–25 | 29,221 | 4.40 | 23.10 | 26% | 17% |

The table reproduces the repository comparison dataset; blank cells are missing values, not zeroes. Reporting periods and disclosure boundaries differ.

\n### Additional interpretation\n\n- **ICICI** has the lowest GHG and water intensity and lowest turnover in this four-company dataset, with the highest female wage-share proxy (26%).\n- **Muthoot Finance** reports the highest GHG and water intensity in this sample; this may reflect differences in branch footprint and operations, but the dataset alone cannot establish the cause.\n- **Aditya Birla Capital** has the highest turnover (38.87%). The report tracks it over four years; turnover peaked in 2023–24 and eased thereafter, while emissions and water series fluctuate in ways that complicate trend interpretation.\n- The **female wage-share proxy** is uneven and reflects workforce composition and pay distribution; it is not a gender-pay-gap measure.\n\nThis is a small sample, not a composite sustainability score or overall ranking.\n\n## Bajaj Finserv Section C data

The cleaned extraction includes **542 rows** across Principles P1–P9, in Essential and Leadership indicator categories. It preserves the report's raw “As Reported” values and includes numeric values, units, periods, column labels, table items and qualitative answers.

### Data files

| File | Description |
|---|---|
| [Section C workbook](data/Bajaj_Finserv_BRSR_SectionC_clean.xlsx) | Multi-sheet workbook: Read Me, Section C Data, Wide Tables, Qualitative and Raw Text |
| [Section C numeric/structured data (CSV)](data/bajaj_finserv_sectionC_data.csv) | Row-level extracted indicators and values |
| [Wide tables (CSV)](data/bajaj_finserv_wide_tables.csv) | Complex multi-column disclosures preserved in table-row form |
| [Qualitative responses (CSV)](data/bajaj_finserv_qualitative.csv) | Narrative answers linked to principle/question and indicator |

\n### Environment (Principle 6): FY2024 vs FY2023\n\n| Indicator | FY2023 | FY2024 | Change |\n|---|---:|---:|---:|\n| Total energy (GJ) | 217,216 | 270,423 | +24.5% |\n| Water consumption (KL) | 704,829 | 870,931 | +23.6% |\n| Scope 1 (tCO₂e) | 14,843 | 6,057 | −59.2% |\n| Scope 2 (tCO₂e) | 36,793 | 42,089 | +14.4% |\n| Scope 3 (tCO₂e) | 18,747 | 28,753 | +53.4% |\n| Waste generated (MT) | 19.5 | 103.59 | +431% |\n\nAbsolute energy, water and Scope 2 and 3 emissions rose while energy intensity per rupee of turnover improved. Renewable energy is reported as under 1% of total energy consumed.\n\n### Findings from other principles\n\n- **Ethics (P1):** training coverage is 100% for the Board, key managerial personnel and employees; anti-corruption review is internal, with no external assessment.\n- **Employee well-being (P3):** women are about 12.5% of permanent employees. Return-to-work after parental leave is 99%, while subsequent retention is 77%.\n- **Human rights (P5):** sexual-harassment complaints rose from 14 to 38 and upheld complaints from 10 to 21; the data cannot determine whether reporting trust or incidence changed.\n- **Inclusive growth (P8):** MSME/small-producer sourcing fell from 9.80% to 8.43%; wage cost is concentrated in metropolitan areas (66.4%), with 1.6% in rural areas.\n- **Consumers (P9):** complaints rose about 19%, from 10,551 to 12,547.\n\n### Disclosure quality\n\nAll 40 FY2025 leadership indicators are marked disclosed, but answers range from numeric data to narrative, “not applicable” and “nil.” Completeness alone therefore does not establish decision-useful disclosure quality.\n\n## Report

The [full analysis report](reports/BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx) includes:
1. Four-company headline comparison for FY2024–25 (as disclosed).
2. Principle-wise Bajaj Finserv Section C review (FY2024 vs FY2023).
3. A multi-year Aditya Birla Capital trend analysis.
4. FY2025 leadership-disclosure completeness review.
5. Findings, caveats and extraction/data notes.

## Scraping code

Both scripts target Section C of BRSR PDFs, using PDF text extraction with layout preserved so values can be assigned to the right columns/periods. The workflow separates structured numeric rows, wide table rows, qualitative text and cleaned source text for manual validation.

- [Python scraper](code/python/brsr_sectionC_clean_scraper.py)
- [R scraper](code/R/scrape_brsr_sectionC.R)

**Python prerequisites:** Python, pandas, openpyxl, and Poppler's `pdftotext` (preferred); the script has a pdfplumber fallback for PDF text extraction.

**R prerequisites:** R, `pdftools`, `openxlsx`, and Poppler's `pdftotext` (preferred).

The scripts are designed for the report layout and may require adjustment for other companies or differently formatted PDFs. Review outputs against the original PDF before using results.

## Key observations and cautions

- The cross-company metrics are a small descriptive sample and are not a composite sustainability score.
- ICICI reports the lowest GHG intensity and water intensity among the four rows in the comparison dataset; Muthoot Finance reports the highest values for these two intensity measures. This is not a comprehensive performance ranking.
- Bajaj Finserv's extracted trend analysis shows absolute energy, water and Scope 2/3 emissions rising while energy intensity per turnover improved; Scope 1 emissions fell in the period described in the report.
- The dataset's Bajaj Finserv water-per-employee figure is unusually small and may reflect a unit inconsistency—treat it as a warning, not as a valid comparative signal.
- Reporting periods, organizational boundaries and intensity denominators vary, limiting direct comparisons.
- The female wage-share proxy is not a gender pay gap or equal-pay measure.
- The report uses “Bajaj Finance” in some deep-dive headings, while the source figures and context appear to relate to Bajaj Finserv/BFS Group; this naming inconsistency is noted for validation.
- Data are self-reported disclosures and should be checked against the source report before interpretation.

## Repository structure

```
BRSR-Sustainability-Comparison/
├── README.md
├── index.html
├── code/
│   ├── R/scrape_brsr_sectionC.R
│   └── python/brsr_sectionC_clean_scraper.py
├── data/
│   ├── Bajaj_Finserv_BRSR_SectionC_clean.xlsx
│   ├── bajaj_finserv_sectionC_data.csv
│   ├── bajaj_finserv_wide_tables.csv
│   ├── bajaj_finserv_qualitative.csv
│   └── cross_company_comparison.csv
└── reports/
    └── BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx
```

## Tools

R · Python · PDF text extraction · Excel · CSV · BRSR/ESG disclosure analysis · Data cleaning · Comparative analysis

---

Portfolio project | BS Analytics and Sustainability Studies, TISS Mumbai
