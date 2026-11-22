# BRSR Sustainability Comparison

A data-driven review of **Business Responsibility and Sustainability Report (BRSR)** disclosures from Indian financial-services companies, combining a cross-company comparison with a detailed Section C extraction and qualitative review of Bajaj Finserv.

**Portfolio:** [Open the project page](https://waliabidi2315.github.io/BRSR-Sustainability-Comparison/)  
**Full report:** [Download the Word report](reports/BRSR_Analysis_of_Five_Indian_Financial_Services_Companies.docx)

## Project scope

The project uses company BRSR disclosures to explore how environmental, workforce, governance and responsible-business indicators are reported—and how comparable those disclosures are in a financial-services context. The work is descriptive; it is **not an overall ESG rating or investment recommendation**.

The project has two parts:
- **Cross-company comparison:** four financial-services companies, using a common set of headline environmental and workforce measures.
- **Company deep dive:** structured extraction and qualitative analysis of Bajaj Finserv's Section C, covering the NGRBC principles, year comparisons, wider tables and narrative answers.

## Companies in the comparison

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

## Bajaj Finserv Section C data

The cleaned extraction includes **542 rows** across Principles P1–P9, in Essential and Leadership indicator categories. It preserves the report's raw “As Reported” values and includes numeric values, units, periods, column labels, table items and qualitative answers.

### Data files

| File | Description |
|---|---|
| [Section C workbook](data/Bajaj_Finserv_BRSR_SectionC_clean.xlsx) | Multi-sheet workbook: Read Me, Section C Data, Wide Tables, Qualitative and Raw Text |
| [Section C numeric/structured data (CSV)](data/bajaj_finserv_sectionC_data.csv) | Row-level extracted indicators and values |
| [Wide tables (CSV)](data/bajaj_finserv_wide_tables.csv) | Complex multi-column disclosures preserved in table-row form |
| [Qualitative responses (CSV)](data/bajaj_finserv_qualitative.csv) | Narrative answers linked to principle/question and indicator |

## Report

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
