# Intuitive Surgical driver-based valuation

A formula-driven forecast with a Python numerical cross-check, bull/base/bear cases, discount-rate sensitivity and a reverse-growth solver. This is an educational research model; it does not establish an investment recommendation.

## Model basis

The model forecasts 2026-2035 and values 2027-2035 cash flows at **December 31, 2026**. Its equity bridge combines reported September 2026 cash with a simplified Q4 cash-flow estimate (25% of annual modeled FCF), settlement and noncontrolling-interest adjustments. It is a forward valuation, not a value discounted to the November benchmark quote date.

The supplied $407.24 benchmark is retained as an input and is **not independently quote-verified**. It must not be described as today's market price.

## Run

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-dev.txt
python -m pytest -q
python scripts/reverse_dcf.py
```

Eight tests cover source revenue reconciliation, reverse-solver round trips, discount-rate sensitivity, invalid terminal assumptions, unattainable targets and scenario ordering. `reports/valuation_results.json` contains numerical results.

Open `ISRG_Valuation_Model.xlsx` in Excel. Core forecast and sensitivity cells are formulas. Reverse-growth outputs are build-time solver results and must be rerun after model inputs change, or solved with Excel Goal Seek. The original openpyxl rebuild script is retained and updated; changing it and rerunning `python scripts/build_model.py` generates an uncached workbook that Excel must recalculate. The supplied reviewed workbook has been recalculated and checked against the Python engine.

## Source corrections

FY2025 revenue is $10,064.7M: instruments/accessories $6,018.9M, systems $2,473.7M, and services $1,572.1M. The supplied model had rounded segment inputs that did not match the filing exactly. Latest core revenue/KPI, cash and share-count inputs were checked against the filings below. Legacy expense and depreciation inputs have not all been individually source-verified.

- [FY2025 SEC 10-K](https://www.sec.gov/Archives/edgar/data/1035267/000103526726000010/isrg-20251231.htm)
- [September 2026 SEC 10-Q](https://www.sec.gov/Archives/edgar/data/1035267/000103526726000071/isrg-20260930.htm)

## Boundaries

The 14.5% first-year procedure-growth rate is a modeling assumption near the midpoint of Q2 guidance, not a precisely observed growth result. Da Vinci operating drivers proxy revenue that also includes Ion. Working-capital/leased-system investment, discount rates, terminal growth and compensation treatment are assumptions. Adding compensation back without dilution is a sensitivity, not free cash.

`Comps` remains an incomplete peer-research template. Peer quotes are unverified, and the reviewed workbook suppresses peer-median outputs. Do not claim a validated multi-company comparable valuation.

Developed with AI coding assistance. See `WALKTHROUGH_CN.md` for the local explanation checklist.
