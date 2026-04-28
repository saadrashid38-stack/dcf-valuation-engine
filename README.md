# DCF Valuation Engine

A professional-grade Discounted Cash Flow (DCF) valuation tool combining a structured Excel financial model with a Python automation layer that pulls live market data from Yahoo Finance.

Built by **Saad Rashid** — CFA Charterholder | MSc Business Analytics & AI, ESCP Paris

---

## What This Does

- Models a 5-year DCF with explicit revenue growth, EBITDA margin, capex, and working capital assumptions
- Calculates Free Cash Flow, WACC, Terminal Value, Enterprise Value, and implied share price
- Runs three scenarios (Base, Bull, Bear) and a sensitivity table (WACC × Terminal Growth Rate)
- Pulls **live** comparable company data (EV/EBITDA, EV/Revenue, P/E) via Python and writes it into Excel automatically

---

## Project Structure
---

## Excel Model — Tab Overview

| Tab | Description |
|-----|-------------|
| Cover | Project title and navigation |
| Assumptions | All input drivers (growth, margins, WACC) |
| DCF Model | 5-year FCF build, PV calculations, EV bridge |
| Scenarios | Base / Bull / Bear comparison |
| Sensitivity | WACC × Terminal Growth Rate matrix |
| Output Summary | One-page valuation summary |
| Comps | Live comparable company multiples (Python-generated) |

---

## How to Run

```bash
# 1. Clone the repo
git clone https://github.com/saadrashid38-stack/dcf-valuation-engine.git
cd dcf-valuation-engine

# 2. Set up virtual environment
python -m venv venv
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Fetch live comps and update Excel
python scripts/fetch_comps.py
```

---

## Key Output (Base Case)

| Metric | Value |
|--------|-------|
| Revenue CAGR (5yr) | ~27% |
| EBITDA Margin (Yr 5) | ~22% |
| WACC | 10% |
| Terminal Growth Rate | 3% |
| Enterprise Value | Calculated dynamically |
| Implied Share Price | Calculated dynamically |

---

## Tech Stack

- **Python 3.14** — data automation and Excel writing
- **yfinance** — live market data
- **openpyxl** — Excel formatting and output
- **pandas** — data manipulation
- **Excel** — financial model and scenario analysis

---

## About

This project demonstrates the intersection of institutional finance knowledge and technical implementation — the combination that quant-adjacent PE, FP&A, and asset management roles require.

**LinkedIn:** [Saad Rashid](https://linkedin.com/in/saad-rashid-cfa)  
**CFA Charterholder | ESCP MSc Business Analytics & AI**
