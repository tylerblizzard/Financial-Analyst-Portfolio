# Investment Banking 3-Statement Financial Model

A comprehensive, integrated 3-statement financial model built to investment banking standards with full scenario analysis, proper debt schedule, and model validation checks.

## Model Overview

This Excel-based financial model provides:
- **Fully integrated** Income Statement, Balance Sheet, and Cash Flow Statement
- **Scenario analysis** with Base, Upside, and Downside cases
- **Proper debt schedule** with Term Loan and Revolver facilities
- **Automated checks** to validate model integrity
- **Executive summary** with charts and scenario comparison

### Time Periods
- **Historical:** 2021-2024 (4 years)
- **Forecast:** 2025-2029 (5 years)

## File Structure

```
3_Statement_Financial_Model.xlsx - The complete financial model
├── Summary - Executive dashboard with charts, scenario comparison, and valuation
├── Assumptions & Drivers - All model inputs, scenario toggles, and DCF assumptions
├── Income Statement - P&L with segment revenue
├── Balance Sheet - Assets, liabilities, and equity
├── Debt Schedule - Term loan and revolver with average balance interest
├── Cash Flow - Indirect method with revolver plug
├── DCF - Unlevered FCF valuation with WACC, terminal values, and sensitivity tables
└── Checks - Model validation and integrity checks (includes DCF checks)
```

## Key Features

### 1. Scenario Analysis
Change the scenario in **Assumptions & Drivers (Cell B4)** to see the model recalculate:
- **Base Case:** Moderate growth, stable margins
- **Upside Case:** Higher growth, margin expansion
- **Downside Case:** Conservative growth, margin compression

Scenario-driven assumptions include:
- Revenue growth rates by segment (Product/Service)
- Gross margin %
- SG&A % of revenue
- R&D % of revenue
- CapEx % of revenue
- Working capital days (AR, Inventory, AP)

### 2. Debt Schedule
**Two debt facilities:**
- **Term Loan:** Fixed principal, 6% interest rate
- **Revolver:** Flexible draw/paydown, 5% interest rate

**Key features:**
- Interest calculated on average balances (more accurate than period-end)
- Revolver automatically draws when cash < $20mm minimum
- Excess cash automatically pays down revolver

### 3. Revenue Model
**Two segments:**
- **Product Revenue:** Tangible goods with separate growth drivers
- **Service Revenue:** Services/subscriptions with independent assumptions

### 4. Working Capital Drivers
Scenario-sensitive operating assumptions:
- **Accounts Receivable:** Days sales outstanding
- **Inventory:** Days inventory on hand (based on COGS)
- **Accounts Payable:** Days payable outstanding (based on COGS)
- Other current assets/liabilities as % of revenue

### 5. PP&E Roll-forward
Standard capital expenditure modeling:
- Beginning PP&E (net)
- \+ Capital Expenditures (% of revenue, scenario-driven)
- \- Depreciation & Amortization (% of revenue)
- = Ending PP&E (net)

### 6. Model Checks
The **Checks** tab validates:
- ✓ Balance Sheet balances (Assets = Liabilities + Equity)
- ✓ Cash Flow reconciles to Balance Sheet cash
- Status shows "PASS" when all checks succeed

### 7. Executive Summary
The **Summary** tab includes:
- Current scenario indicator
- Key metrics comparison table (ready for manual scenario capture)
- Three charts tracking:
  - Revenue growth over time
  - EBITDA trend
  - Free Cash Flow
- **DCF Valuation Summary** with perpetuity and exit multiple methods
- Enterprise Value and Equity Value per Share
- Key DCF assumptions display

### 8. DCF Valuation
The **DCF** tab provides comprehensive valuation analysis:

**Unlevered Free Cash Flow Calculation:**
- NOPAT (EBIT × (1-Tax))
- Plus: Depreciation & Amortization
- Less: Capital Expenditures
- Less: Increase in Net Working Capital
- = Unlevered FCF (cash available to all capital providers)

**WACC (Weighted Average Cost of Capital):**
- **Cost of Equity via CAPM:** Risk-free rate + Beta × Equity risk premium
- **After-tax Cost of Debt:** Pre-tax rate × (1 - Tax rate)
- **Capital Structure:** Target debt/equity weights
- **WACC Formula:** (% Equity × Cost of Equity) + (% Debt × After-tax Cost of Debt)

**Terminal Value - Two Methods:**
1. **Perpetuity Growth Method:**
   - Terminal FCF × (1 + g) / (WACC - g)
   - Default: 2.5% perpetual growth rate

2. **Exit Multiple Method:**
   - Terminal year EBITDA × Exit EV/EBITDA multiple
   - Default: 8.5x EBITDA multiple

**Valuation Output:**
- Present value of forecast period FCFs (2025-2029)
- Present value of terminal value
- **Enterprise Value** = PV(FCFs) + PV(Terminal Value)
- Less: Net Debt (Total Debt - Cash)
- **= Equity Value**
- **Equity Value per Share** = Equity Value / Fully Diluted Shares

**Sensitivity Analysis:**
Two sensitivity table frameworks (use Excel Data Table feature):
- **Table 1:** WACC vs Terminal Growth Rate → Equity Value per Share
- **Table 2:** WACC vs Exit EBITDA Multiple → Equity Value per Share

## Color Coding

Consistent formatting throughout:
- **Blue cells** = User inputs (editable assumptions)
- **Black text** = Formulas (calculated values)
- **Green text** = Check results and validations
- **Gray fill** = Section headers
- **Dark blue headers** = Column/row labels

## How to Use

### Basic Usage
1. Open `3_Statement_Financial_Model.xlsx`
2. Go to **Summary** tab for high-level overview
3. Navigate to **Assumptions & Drivers** tab
4. Change **Scenario** (Cell B4) to Base/Upside/Downside
5. Model automatically recalculates all statements
6. Review **Checks** tab to ensure model integrity

### Customizing Assumptions
1. Go to **Assumptions & Drivers** tab
2. Modify any **blue cells** to change inputs
3. Historical data (2021-2024) can be updated
4. Scenario assumptions (rows 12-23) drive forecast behavior
5. Fixed assumptions (rows 26-47) apply across all scenarios

### Scenario Comparison
To compare scenarios:
1. Set scenario to "Base" in Assumptions & Drivers
2. Record metrics from Summary tab (or specific statements)
3. Change scenario to "Upside"
4. Record metrics again
5. Repeat for "Downside"
6. Use the comparison table in Summary (rows 6-11) to track differences

### Understanding the Debt Schedule
The **Debt Schedule** tab shows:
- **Term Loan:** Annual beginning/ending balances, interest expense
- **Revolver:** Automatic draws/paydowns based on cash needs
- **Total Interest:** Feeds into Income Statement

The revolver logic ensures:
- Minimum cash balance of $20mm is maintained
- Excess cash automatically pays down revolver
- Draws occur when operating cash flow insufficient

### Using the DCF Valuation
The **DCF** tab calculates company valuation:

**Key DCF Inputs to Adjust:**
1. **WACC Components:**
   - Risk-free rate (default: 4.5%)
   - Equity risk premium (default: 6.5%)
   - Beta (default: 1.2)
   - Target debt % (default: 30%)

2. **Terminal Value Assumptions:**
   - Perpetual growth rate (default: 2.5%)
   - Exit EV/EBITDA multiple (default: 8.5x)

3. **Share Count:**
   - Fully diluted shares (default: 100mm)

**Reading DCF Output:**
- **Enterprise Value (Perpetuity):** Cell C60
- **Equity Value (Perpetuity):** Cell C63
- **Value per Share (Perpetuity):** Cell C66
- **Enterprise Value (Exit Multiple):** Cell C73
- **Equity Value (Exit Multiple):** Cell C75
- **Value per Share (Exit Multiple):** Cell C76

**Creating Sensitivity Tables:**
1. Select the sensitivity table range (e.g., B81:G86 for first table)
2. Go to Data → What-If Analysis → Data Table
3. For WACC vs Growth table:
   - Row input cell: C38 (Terminal Growth Rate)
   - Column input cell: C32 (WACC)
4. Click OK to populate the table
5. Repeat for Exit Multiple table (B93:I98)

The DCF automatically updates when you change scenarios, as operating assumptions affect FCF generation.

## Model Logic Flow

```
Assumptions & Drivers (Operating + DCF assumptions)
        ↓
Income Statement (Revenue → EBITDA → EBIT → Net Income)
        ↓
Balance Sheet (Assets = Liabilities + Equity)
        ↓
Cash Flow Statement (CFO → CFI → CFF)
        ↓
Debt Schedule (Revolver plug based on cash needs)
        ↓ (circular reference resolved)
Back to Balance Sheet & Income Statement
        ↓
DCF Analysis (pulls EBIT, D&A, CapEx, NWC changes)
        ↓
Unlevered FCF → WACC → Terminal Value → Enterprise Value
        ↓
Less: Net Debt → Equity Value → Value per Share
        ↓
Summary Tab (integrates all outputs + valuation)
```

## Technical Details

### Circular Reference Resolution
The model handles circularity through the revolver:
1. Cash Flow calculates cash before financing
2. If cash < $20mm minimum, revolver draws
3. Revolver balance updates in Debt Schedule
4. Interest expense updates in Income Statement
5. Net Income updates, affecting equity in Balance Sheet
6. Model iterates until balanced

Excel's iterative calculation feature handles this automatically.

### Free Cash Flow Calculation
**Unlevered FCF = CFO + CapEx**

(CapEx is negative, so this is CFO - |CapEx|)

This represents cash available to all capital providers before financing activities.

## Generating the Model

Three Python scripts are included for reproducibility:

### build_model.py
Creates the initial 3-statement model from scratch.

```bash
python3 build_model.py
```

### enhance_model.py
Enhances an existing model with IB-level features (debt schedule, scenarios, checks).

```bash
python3 enhance_model.py
```

### add_dcf_valuation.py
Adds comprehensive DCF valuation analysis to the model.

```bash
python3 add_dcf_valuation.py
```

**To build the complete model from scratch:**
```bash
python3 build_model.py
python3 enhance_model.py
python3 add_dcf_valuation.py
```

**Requirements:**
- Python 3.7+
- openpyxl library: `pip install openpyxl`

## Model Validation

Before relying on model outputs:

1. **Check the Checks tab** - All checks should show "PASS"
2. **Verify Balance Sheet balances** - Check row should be ~0
3. **Test scenario switching** - Change scenarios and verify recalculation
4. **Review debt schedule** - Ensure revolver draws/pays down logically
5. **Spot-check formulas** - Sample random cells to verify formula logic

## Common Issues & Troubleshooting

### Balance Sheet Doesn't Balance
- Check Checks tab for specific errors
- Verify all formulas link correctly between sheets
- Ensure circular references enabled in Excel (File → Options → Formulas)

### Negative Cash with No Revolver Draw
- Verify minimum cash balance assumption (should be $20mm)
- Check revolver draw formula in Cash Flow statement
- Ensure Debt Schedule links are correct

### Scenario Not Changing Outputs
- Verify Cell B4 in Assumptions & Drivers contains valid scenario name
- Check that forecast formulas reference scenario-driven assumptions (column E in rows 12-23)
- Recalculate workbook manually (Ctrl+Alt+F9)

## Best Practices

1. **Always use the scenario dropdown** rather than editing individual forecast assumptions
2. **Document any manual overrides** if you edit blue cells differently per scenario
3. **Keep historical data separate** from forecast assumptions
4. **Review Checks tab after any major changes**
5. **Use Save As with version numbers** when making significant modifications
6. **Test all scenarios** after model changes to ensure consistent logic

## Extensions & Customizations

This model already includes comprehensive DCF valuation. Additional extensions could include:
- **Trading/Transaction Comps** - Comparable company analysis tab
- Additional revenue segments or business units
- More granular operating expense detail
- Multiple debt tranches with varying terms and covenants
- Dividend/distribution modeling
- Detailed tax schedules (NOLs, deferred taxes, deferred tax assets/liabilities)
- Quarterly working capital forecasts for more granular modeling
- Credit metrics analysis (leverage ratios, coverage ratios)
- Returns analysis (IRR, MOIC for PE use cases)
- Management case vs adjusted case reconciliation
- Accretion/dilution analysis for M&A scenarios

## License

Open source - feel free to use and modify for your needs.

## Contact

For questions or issues, please open a GitHub issue.
