# Crop Yield Prediction in India (District Level)

**Intern:** Shivbhushan Gaikwad | **Program:** Yuva Intern

Machine learning project for agriculture and agribusiness. This repository holds the **Week 1: Data Acquisition and Problem Scoping** deliverable and the project skeleton for later weeks.

## Problem statement
Predict seasonal yield (tonnes per hectare) of major crops at district level in India using historical yield records combined with weather, soil and satellite vegetation data, and flag districts at elevated risk of below-normal yield.

## Workflow
Problem Definition -> Identify Data Requirements -> Collect Data -> Understand Dataset -> Check Data Quality -> Document Findings

## Repository structure
```
.
├── docs/
│   ├── Week1_Data_Acquisition_and_Problem_Scoping.docx   # Week 1 report
│   ├── data_catalogue.csv                                # candidate data sources
│   └── decision_log.md                                   # decisions and reasoning
├── data/
│   ├── raw/          # original downloads (git-ignored)
│   └── processed/    # cleaned datasets (git-ignored)
├── notebooks/        # exploration and modelling notebooks
├── src/              # reusable Python code
├── scripts/
│   └── build_report.js   # regenerates the Word report
├── requirements.txt  # Python packages for later weeks
└── package.json      # Node dependency for the report script
```

## Timeline (8 weeks)
| Week | Focus |
|------|-------|
| 1 | Problem scoping and data research (done) |
| 2 | Data acquisition |
| 3 | Cleaning and integration |
| 4 | Exploratory analysis |
| 5 | Feature engineering and baselines |
| 6 | Modelling |
| 7 | Evaluation and interpretation |
| 8 | Final reporting |

## Rebuild the report
```bash
npm install
npm run build:report
```

## Data sources (planned)
Official district yield statistics (Ministry of Agriculture / data.gov.in), IMD or NASA POWER weather, SoilGrids soil, MODIS NDVI. Verify each source current access and licence before download.
