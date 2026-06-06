<p align="center">
  <img src="https://raw.githubusercontent.com/Nishanth20/India-home-loan-simulator-excel/main/screenshots/dashboard.png" width="820" alt="Home Loan Strategy Lab — Dashboard"/>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Excel-2016%2B-217346?logo=microsoftexcel&logoColor=white" alt="Excel"/>
  <img src="https://img.shields.io/badge/VBA-Macro--Enabled-blue" alt="VBA"/>
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License"/>
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white" alt="Platform"/>
  <img src="https://img.shields.io/badge/Made%20for-Indian%20Home%20Loans-FF9933" alt="Made for India"/>
  <img src="https://img.shields.io/badge/No%20Add--ins-100%25%20Native%20Excel-success" alt="Free"/>
</p>

# 🏠 Home Loan Strategy Lab

*India's most complete Excel-native home loan simulator. Simulate EMI, tenure, prepayments, floating rate changes, and Section 24b/80C tax benefits — entirely in VBA, no add-ins.*

## Table of contents

| Section | Link |
| --- | --- |
| What this does | [What this does](#what-this-does) |
| Screenshots | [Screenshots](#screenshots) |
| Features | [Features](#features) |
| How it works | [How it works](#how-it-works) |
| Getting started | [Getting started](#getting-started) |
| Sheet guide | [Sheet guide](#sheet-guide) |
| Prepayment modes | [Prepayment modes](#prepayment-modes) |
| Tax benefits tracked | [Tax benefits tracked](#tax-benefits-tracked) |
| Scenario compare — what gets simulated | [Scenario compare — what gets simulated](#scenario-compare--what-gets-simulated) |
| FAQ | [FAQ](#faq) |
| Contributing | [Contributing](#contributing) |
| License | [License](#license) |

## What this does

It estimates your month-by-month home loan journey inside Excel.  
It updates EMI, tenure, tax benefit, and balance after your prepayments.  
It compares normal, planned, and stressed loan paths before you commit.  

**Key capabilities:**
- ✅ Full amortization schedule with EMI calculations
- ✅ Multiple prepayment strategies (lump sum, extra EMI)
- ✅ Floating rate tracking and stress testing
- ✅ Section 24(b) and Section 80C tax benefit calculations
- ✅ Scenario comparison engine
- ✅ Annual summary rollups
- ✅ 100% Excel-native — no external dependencies

## Screenshots

### Loan_Input
<p align="center"><img src="screenshots/loan-input.png" width="750"/></p>
*This sheet captures your loan inputs and prepayment plan.*  

### Loan_Schedule
<p align="center"><img src="screenshots/loan-schedule.png" width="750"/></p>
*This sheet shows the full amortization schedule and editable prepayment columns.*  

### Dashboard
<p align="center"><img src="screenshots/dashboard.png" width="750"/></p>
*This sheet shows headline loan numbers and a quick recommendation.*  

### Scenario_Compare
<p align="center"><img src="screenshots/scenario-compare.png" width="750"/></p>
*This sheet compares four loan strategies on interest, payoff, and prepayment.*  

### Rate_Changes
<p align="center"><img src="screenshots/rate-changes.png" width="750"/></p>
*This sheet stores month-wise floating rate changes for the simulation.*  

### Annual_Summary
<p align="center"><img src="screenshots/annual-summary.png" width="750"/></p>
*This sheet rolls monthly output into a year-by-year summary.*  

## Features

| Feature | What it does |
| --- | --- |
| **Tool** | Microsoft Excel macro-enabled workbook (`.xlsm`) |
| **Language** | VBA (100% codebase) |
| **Code structure** | Single `LoanModule.bas` file for all business logic |
| **Purpose** | Simulates Indian home loan cash flow and prepayment strategies |
| **Sheets** | `Loan_Input`, `Loan_Schedule`, `Rate_Changes`, `Dashboard`, `Scenario_Compare`, `Annual_Summary` |
| **Core inputs** | Loan amount, annual interest rate, tenure (years), start date, prepayment mode |
| **Prepayment modes** | `REDUCE_EMI` (lower monthly burden) and `REDUCE_TENURE` (faster closure) |
| **Lump sum support** | Month-wise lump sum prepayments |
| **Extra EMI support** | Month-wise extra EMI counts |
| **Floating rate support** | Month-wise rate changes with automatic recalculation |
| **Tax tracking** | Section 24(b) up to ₹2,00,000 and Section 80C up to ₹1,50,000 |
| **Comparison engine** | Base Loan, Your Plan EMI, Your Plan Tenure, and Rate Stress +2% scenarios |
| **Dashboard view** | Loan Amount, Rate %, Tenure, Payoff Months, Total Interest, Total Prepayment, Years Saved, Interest Saved |
| **Visualization** | Interest vs Principal line chart by month |
| **Number format** | Indian lakh and crore style with ₹ symbol |
| **Date labels** | `MMM-YYYY` format based on loan start date |
| **Interactive buttons** | Generate Loan Schedule, Recalculate Schedule, Refresh Annual Summary |
| **Compatibility** | Excel 2016+ on Windows only |
| **Dependencies** | None — 100% native Excel, no add-ins or external tools |

## How it works

1. **Setup**: Run the setup macro once to build the workbook structure
2. **Input**: Enter loan details, start date, and any prepayment plan in `Loan_Input`
3. **Generate**: Run the schedule macro to compute sheets, chart, taxes, and comparisons

The VBA engine automatically:
- Calculates month-by-month amortization
- Applies prepayments in your chosen mode
- Tracks floating rate changes
- Computes tax-eligible amounts
- Runs scenario comparisons
- Generates summary reports

## Getting started

### Prerequisites

| Requirement | Details |
| --- | --- |
| **Operating system** | Windows only |
| **Excel version** | Excel 2016 or later |
| **Macro setting** | Macros must be enabled |
| **Additional software** | None required |

### Installation

1. Download `HomeLoanStrategyLab.xlsm` from the [Releases](https://github.com/Nishanth20/India-home-loan-simulator-excel/releases) tab
2. Open in Excel. Click **Enable Macros** when prompted
3. Go to **Developer** tab → **Macros** → run **SetupLoanCalculator** once
4. Fill your loan details in the **Loan_Input** sheet
5. Click **Generate Loan Schedule**
6. Review the **Dashboard** and **Scenario_Compare** sheets

### Quick start example

```
Loan Amount: ₹50,00,000
Annual Rate: 6.5%
Tenure: 20 years
Start Date: Jan 2024
Mode: REDUCE_TENURE
Prepayment: ₹50,000 every 6 months
```

Run the schedule and see your payoff months, interest savings, and tax benefits automatically calculated.

## Sheet guide

### Loan_Input

The control panel for the entire simulation.

- **Loan Amount**: Principal borrowed (in ₹)
- **Annual Interest Rate**: Fixed or current floating rate (%)
- **Tenure**: Loan duration in years
- **Start Date**: When the loan begins
- **Prepayment Mode**: `REDUCE_EMI` or `REDUCE_TENURE`
- **Prepayment Schedule**: Month-wise lump sum or extra EMI counts

### Loan_Schedule

The full amortization schedule showing every month of the loan.

- **Month**: Sequence number from loan start
- **Date**: Calendar month and year
- **Principal**: Amount paid toward principal
- **Interest**: Amount paid toward interest
- **EMI**: Equated monthly installment
- **Balance**: Outstanding principal after the month
- **Lump Sum**: Optional prepayment (editable)
- **Extra EMI**: Optional additional EMI counts (editable)
- **Tax tracking**: Section 24(b) and 80C eligible amounts

After editing lump sums or extra EMIs, click **Recalculate Schedule** to update dependent columns.

### Rate_Changes

Stores future floating rate changes.

- **Month Number**: When the rate change takes effect (e.g., month 13)
- **New Annual Rate**: The revised interest rate (%)

Example: If your bank cuts rates at month 25, enter 25 in Month Number and the new rate in New Annual Rate. Click **Recalculate Schedule** to regenerate the loan plan.

### Dashboard

The executive summary with key metrics and a quick recommendation.

Shows:
- **Payoff period** (months and years)
- **Total interest** over the loan life
- **Total prepayment** amount
- **Years saved** vs. base loan
- **Interest saved** vs. base loan
- **Estimated tax benefit** for the fiscal year
- **Interest vs Principal chart** showing cash flow composition by month

### Scenario_Compare

Side-by-side comparison of four loan strategies.

- **Base Loan**: No prepayments, baseline tenure and interest
- **Your Plan EMI**: Your prepayment plan applied with EMI reduction
- **Your Plan Tenure**: Your prepayment plan applied with tenure reduction
- **Rate Stress +2%**: Your prepayment plan under a hypothetical +2% rate increase

Use this to evaluate trade-offs in cash flow relief vs. interest savings.

### Annual_Summary

Groups monthly results into loan years (financial year or calendar year).

Shows year-by-year:
- **Annual EMI** (sum of 12 months)
- **Annual interest** paid
- **Annual principal** paid
- **Section 24(b) tax benefit**
- **Section 80C deduction** eligible
- **Year-end balance**

Useful for tax planning and year-by-year reviews.

## Prepayment modes

| Mode | EMI | Tenure | Best for |
| --- | --- | --- | --- |
| **REDUCE_EMI** | Falls after prepayment | Stays fixed | Lower monthly cash burden |
| **REDUCE_TENURE** | Stays steady | Shortens after prepayment | Faster loan closure and interest savings |

**Example:**
- Original loan: ₹50 lakh at 7% over 20 years = ₹388/month EMI
- After ₹10 lakh prepayment under REDUCE_EMI: EMI drops to ~₹310/month, tenure stays 20 years
- After ₹10 lakh prepayment under REDUCE_TENURE: EMI stays ₹388/month, tenure drops to ~15 years

## Tax benefits tracked

### Section 24(b) — Home Loan Interest Deduction

- **Annual cap**: ₹2,00,000 per financial year
- **Eligibility**: Interest paid on home loans (self-occupied or let-out property)
- **How it's computed**: Monthly interest is tracked and capped at ₹2,00,000 per FY
- **In the tool**: Visible in `Loan_Schedule` column P and yearly `Annual_Summary`

### Section 80C — Principal Repayment (Principal Portion Only)

- **Annual cap**: ₹1,50,000 per financial year
- **Eligibility**: Principal repayment amount (only)
- **How it's computed**: Monthly principal is tracked and capped at ₹1,50,000 per FY
- **In the tool**: Visible in `Loan_Schedule` column Q and yearly `Annual_Summary`

**Note**: These are indicative calculations. Consult a tax advisor for your specific situation, as eligibility depends on property classification and lender type.

## Scenario compare — what gets simulated

| Scenario | Prepayments used | Mode | What it tests |
| --- | --- | --- | --- |
| **Base Loan** | None | REDUCE_TENURE | Baseline path — no prepayment impact |
| **Your Plan EMI** | Your lump sum + extra EMI plan | REDUCE_EMI | Cash flow relief after prepayment |
| **Your Plan Tenure** | Your lump sum + extra EMI plan | REDUCE_TENURE | Fastest payoff with your prepayments |
| **Rate Stress +2%** | Your lump sum + extra EMI plan | REDUCE_TENURE | Resilience under rate increase scenario |

**Use case**: Compare interest saved vs. EMI reduced across all four paths to make an informed prepayment strategy.

## FAQ

<details>
<summary><strong>Do I need to install anything beyond Excel?</strong></summary>
<br>
No. It runs as a native Excel VBA workbook. No add-ins, no Python, no external tools, no dependencies.
</details>

<details>
<summary><strong>Does this work on Mac?</strong></summary>
<br>
No. VBA macro execution and the Scripting.Dictionary object require Windows. Excel on Mac does not support this workbook.
</details>

<details>
<summary><strong>Can I use this for a loan that has already started?</strong></summary>
<br>
Yes. Enter your original loan start date and current outstanding principal as the loan amount. The schedule will run from month 1 of your inputs.
</details>

<details>
<summary><strong>What happens when my bank changes my floating rate?</strong></summary>
<br>
Open the `Rate_Changes` sheet. Enter the month number when the change takes effect and the new annual rate. Click **Recalculate Schedule**.
</details>

<details>
<summary><strong>Is my data sent to any server?</strong></summary>
<br>
No. This is a local Excel file with no network calls. Nothing leaves your machine.
</details>

<details>
<summary><strong>Can I modify the VBA code?</strong></summary>
<br>
Yes. All business logic lives in `LoanModule.bas` — a single, clearly structured file. Fork, modify, and use freely under the MIT License.
</details>

<details>
<summary><strong>What if I find a bug or want to suggest a feature?</strong></summary>
<br>
Please [open an issue](https://github.com/Nishanth20/India-home-loan-simulator-excel/issues/new) on GitHub with a description, and we'll review it.
</details>

<details>
<summary><strong>Is there a way to compare multiple loan offers?</strong></summary>
<br>
Yes. Create separate copies of the workbook, fill in each bank's terms, and run the simulations. Compare the `Dashboard` sheets side by side.
</details>

## Use cases

This tool is ideal for:

- **New homebuyers**: Evaluate prepayment strategies before signing the loan agreement
- **Existing borrowers**: Plan prepayments and understand tax deductions
- **Financial planners**: Model loan scenarios for clients
- **Excel enthusiasts**: Learn VBA macros and financial modeling in Excel
- **Tax planning**: Track Section 24(b) and 80C deductions year by year

## Technical details

- **Codebase**: 100% VBA, single module (`LoanModule.bas`)
- **Data structure**: Uses Scripting.Dictionary for efficient lookups
- **Recalculation**: Click buttons to regenerate entire simulation
- **Performance**: Handles up to 360-month loans (30 years) instantly

## Contributing

Pull requests are welcome. All business logic lives in `LoanModule.bas` — it is one file, clearly structured, and easy to read and fork.

**Areas for contribution:**
- Enhanced tax calculations (Section 80D, other deductions)
- Additional scenarios (co-borrower, joint loans)
- Mac compatibility exploration
- New visualization formats
- Bug fixes and optimizations

If you find a bug or want to suggest a feature, please [open an issue](https://github.com/Nishanth20/India-home-loan-simulator-excel/issues/new).

## Disclaimer

This tool is provided as-is for educational and planning purposes only. The calculations are indicative and based on the inputs you provide. Please consult:
- Your bank or lender for official loan terms
- A tax advisor for deduction eligibility and compliance
- A financial advisor for personalized loan strategy recommendations

## License

MIT — See [LICENSE](LICENSE) file for details.

---

**Made for Indian homebuyers. Entirely in Excel. No add-ins. No dependencies.**

Questions? Open an issue or reach out. Happy loan planning! 🏠
