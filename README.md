# Survey QC Automation

Automated quality control pipeline for large-scale survey field data.

Built from real QC workflows applied across 10+ national surveys in Brazil, this project replaces manual verification processes that typically take one to two days per study. The pipeline runs in under an hour and produces a structured report of all flagged cases, ready for review by the research team.

---

## The problem

In large-scale face-to-face (F2F) and telephone (CATI) surveys, quality control is traditionally done manually: analysts review interview records one by one, checking for inconsistencies, suspicious patterns, and logical errors. At scale — hundreds or thousands of interviews per wave — this is slow, error-prone, and doesn't catch everything.

This pipeline automates that process across 18 different checks.

---

## What it checks

| # | Check | Logic |
|---|---|---|
| 1 | Household members | Mismatch between selection grid and questionnaire |
| 2 | Age | Discrepancy > 2 years between selection grid and questionnaire |
| 3 | Gender | Mismatch between selection grid and questionnaire |
| 4 | Education | Years of study inconsistent with declared education level |
| 5 | Minor authorization | Underage respondent interviewed without consent record |
| 6 | Income | Declared value outside expected range |
| 7 | Duration (short) | Interview completed below minimum threshold |
| 8 | Duration (long) | Interview exceeded maximum expected duration |
| 9 | System flags | Issues flagged by the data collection platform |
| 10 | Partial racing | High percentage of questions answered in under 3 seconds |
| 11 | Members under 15 | Unusually high count of household members under 15 |
| 12 | Food spending vs income | Monthly food spending exceeds or is inconsistent with declared income |
| 13 | NEC visits | 3 or more "nobody home" visits to the same household on the same day |
| 14 | Visit intervals | Two consecutive attempts to the same household less than 2 hours apart |
| 15 | First-attempt rate | PSU/cluster with more than 50% of interviews completed on first attempt |
| 16 | Daily interview count | Interviewer with 10 or more completed interviews in a single day |
| 17 | Audio authorization | PSU/cluster with less than 50% of respondents authorizing audio recording |
| 18 | Nighttime interviews | Interview completed between 21h and 6h |

---

## Outputs

- **QC report** (`Voltas_DDMM.xlsx`): one row per flagged interview, with a `Problem` column describing the specific issue found
- **Field schedule report** (`Relatorio_Horarios.xlsx`): summary of visit attempts by time of day and day of week, per interviewer
- **Visit results report** (`Resultados_Tentativas.xlsx`): frequency of visit outcome codes per interviewer, and full list of "other reason" attempts

---

## Project structure

```
survey-qc-automation/
│
├── notebooks/
│   └── survey_qc_walkthrough.ipynb   # Full pipeline walkthrough with synthetic data
│
├── scripts/
│   └── survey_qc.py                  # Production script (configurable for any project)
│
└── README.md
```

---

## How to adapt to a new project

The script uses a `VAR` dictionary inside the main processing function to map project-specific column names to the internal names used by the pipeline. To adapt it to a new survey, update only that dictionary:

```python
VAR = {
    'col_entrevistador':       'YOUR_INTERVIEWER_ID_COLUMN',
    'col_duracao':             'YOUR_DURATION_COLUMN',
    'col_idade_selecao':       'YOUR_SELECTION_AGE_COLUMN',
    # ... and so on
}
```

No other changes to the code are required.

---

## Stack

- `Python 3.x`
- `Pandas` — data wrangling, consistency checks, aggregations
- `NumPy` — numerical operations
- `Matplotlib` — visualizations in the walkthrough notebook
- `openpyxl` — Excel output
- `pytz` — timezone handling for nighttime and scheduling checks

---

## Context

These routines were developed and applied across more than 10 national surveys in Brazil, including studies conducted for international academic and research organizations. The `survey_qc.py` script is designed for production use with real field data; the notebook uses synthetic data to demonstrate the full pipeline without exposing any respondent information.

--- 

## Author

**Teresa De Bastiani**
Senior Market Research Analyst · Florianópolis, Brazil
[LinkedIn](https://linkedin.com/in/mteresadebastiani)
