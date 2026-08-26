# Retirement & Personal Investment Planner

An interactive Streamlit app for long-term retirement planning: it projects portfolio growth from today through retirement, compares that to a required nest egg under the 4% withdrawal rule, stress-tests the plan with market-return scenarios and Monte Carlo simulation, and can export the full model to a multi-tab Excel workbook.

---

## What it does

1. **Accumulation projection** — year-by-year portfolio growth from current age to retirement age, combining starting balance, user contributions (with an annual raise), employer match, and investment gains, alongside an inflation-adjusted ("real") value.
2. **Retirement readiness** — compares the projected nest egg at retirement to the required nest egg (net annual spending after Social Security, divided by 4%, per the 4% withdrawal rule) and computes a funding ratio, classified as At Risk / Needs Improvement / On Track / Strong.
3. **Growth Projection tab** — nominal vs. inflation-adjusted portfolio value chart, plus a breakdown of how much of the final balance came from contributions, employer match, and compound growth (pie chart).
4. **Scenario Analysis tab** — compares projected portfolio value under Conservative (5%), Base Case (8%), and Aggressive (10%) annual return assumptions.
5. **Monte Carlo tab** — runs 5,000 simulated return paths (Gaussian, based on expected return and volatility) to estimate the probability of meeting the retirement goal, with 10th/25th/50th/75th/90th percentile outcomes and a spaghetti chart of sample trajectories.
6. **Full Schedule tab** — the complete year-by-year accumulation table (starting balance, contributions, match, gains, ending portfolio, real purchasing power).
7. **Excel export engine** — builds a 7-tab formula-linked Excel workbook (Dashboard, Assumptions, Accumulation, Retirement, Scenarios, Monte Carlo, Summary) mirroring the in-app model, generated in memory from the current sidebar inputs.

---

## Project structure

```
.
└── app.py   # Streamlit app — single-file, all logic and UI
```

No input files are required — all assumptions are entered through the sidebar and held in the Streamlit session.

---

## Installation

```bash
pip install streamlit pandas numpy matplotlib openpyxl
```

---

## Usage

```bash
streamlit run app.py
```

From the **sidebar**, enter:
- **Personal** — current age, retirement age, life expectancy
- **Portfolio & Savings** — current portfolio balance, annual income, annual contribution, employer match %, annual raise on savings
- **Market Assumptions** — expected annual return, portfolio volatility (std dev), expected inflation rate
- **Retirement Spending** — desired annual spending goal, expected Social Security income

The dashboard KPIs (projected nest egg, required nest egg, funding ratio, planning status) update live, and the four tabs — **Growth Projection**, **Scenario Analysis**, **Monte Carlo**, **Full Schedule** — provide progressively deeper analysis of the plan.

> **Note:** The app builds a full Excel workbook of the model (`build_excel_in_memory`) on every run, but the sidebar download button that offers it to the user is currently commented out in the code. Uncomment the `st.sidebar.download_button(...)` block near the top of the script to re-enable the Excel export.

---

## Methodology

**Accumulation** (per year): `Ending Portfolio = Starting Balance + Contributions + Employer Match + Investment Gain`, where the investment gain is computed on the average of starting balance and mid-year contributions at the expected return rate. Real (inflation-adjusted) value discounts the nominal ending balance by the inflation rate compounded over elapsed years.

**Required nest egg**: `(Retirement Spending − Social Security) / 0.04` — the standard 4% safe-withdrawal-rate rule of thumb.

**Funding ratio**: `Projected Assets / Required Assets`. < 80% → At Risk, 80–99% → Needs Improvement, 100–119% → On Track, ≥ 120% → Strong.

**Monte Carlo**: simulates 5,000 independent paths using annual returns drawn from a Normal distribution parameterized by the expected return and volatility inputs, applied to the same contribution schedule as the deterministic projection. Success probability is the share of simulated paths that meet or exceed the required nest egg.

---

## Notes

- This is an educational/portfolio financial planning model, not personalized investment, tax, or retirement advice — the app is explicitly labeled "for educational use."
- The Monte Carlo simulation uses a simplified Gaussian return model; it doesn't account for sequence-of-returns risk beyond simple compounding, fat tails, or fee drag.
- All inputs and results persist only for the current Streamlit session — nothing is saved to disk unless the Excel export is re-enabled and downloaded.
