# normalize_amfi_csv.py

A Python script to normalize AMFI (Association of Mutual Funds in India) Excel reports into clean, flat CSV files. It handles both **MF AUM (Mutual Fund Assets Under Management)** and **NSR (New Schemes Report)** data.

## Usage

```
python normalize_amfi_csv.py <path_to_excel_file>
```

### Examples

```
python normalize_amfi_csv.py amsep2025repo.xls
python normalize_amfi_csv.py MF_AUM_AMFI.xls
python normalize_amfi_csv.py Sample_NSR_input_Mar.xls
```

## Prerequisites

- Python 3.x
- pandas
- xlrd (for .xls files)
- openpyxl (for .xlsx files)

Install dependencies:
```
pip install pandas xlrd openpyxl
```

## Supported Input Formats

| Format | Description |
|--------|-------------|
| `.xls` | Legacy Excel format (uses xlrd engine) |
| `.xlsx` | Modern Excel format (uses openpyxl engine) |

## Input File Types

### 1. Multi-Sheet Files (e.g., amsep2025repo.xls)
Files containing both AUM and NSR data on separate sheets. The script auto-detects each sheet type and generates **2 CSV files**.

### 2. Single-Sheet MF AUM Files (e.g., MF_AUM_AMFI.xls)
Monthly report with hierarchical scheme data (Open ended, Close Ended, Interval Schemes). Generates **1 CSV file**: `Normalized_MF_AUM.csv`

### 3. Single-Sheet NSR Files
New Schemes Report files in various layouts:

| Variant | Example File | Description |
|---------|--------------|-------------|
| Mar style | Sample_NSR_input_Mar.xls | Has both Open and Closed ended columns (6 cols) |
| Apr style | Sample_NSR_input_Apr.xls | Has only Open ended columns (4 cols) |
| Sep style | Sample_NSR_input_Sep25.xls | Has `*NEW SCHEMES LAUNCHED` section with scheme names listed separately |
| Dec style | Sample_NSR_input_dec25.xls | Header says "Open-ended" or "Close-ended", mapping columns accordingly |

Generates **1 CSV file**: `Normalized_NSR.csv`

## Output Files

Output CSVs are saved in the same directory as the input file.

### Normalized_MF_AUM.csv

| Column | Description |
|--------|-------------|
| Scheme_Category | Top-level category (Open ended Schemes, Close Ended Schemes, Interval Schemes, Fund of Funds Scheme (Domestic)) |
| Sub_Scheme_Category | Sub-category (Income/Debt Oriented Schemes, Growth/Equity Oriented Schemes, Hybrid Schemes, etc.) |
| Fund_Name | Individual fund name (Overnight Fund, Liquid Fund, etc.) |
| No_of_Schemes | Number of schemes (integer) |
| No_of_Folios | Number of folios (integer) |
| Funds_Mobilized | Funds mobilized in INR crore (2 decimal places) |
| Repurchase_Redemption | Repurchase/Redemption amount (2 decimal places) |
| Net_Inflow_Outflow | Net inflow/outflow (2 decimal places, negative for outflow) |
| Net_AUM | Net Assets Under Management (2 decimal places) |
| Average_Net_AUM | Average Net AUM for the month (2 decimal places) |
| No_of_segregated_portfolios | Number of segregated portfolios |
| Net_AUM_in_segregated_portfolio | Net AUM in segregated portfolio |

### Normalized_NSR.csv

| Column | Description |
|--------|-------------|
| Scheme_Category | Category (Income/Debt Oriented Schemes, Growth/Equity Oriented Schemes, Hybrid Schemes, Other Schemes) |
| Scheme_Name | Individual scheme name(s), semicolon-separated if multiple |
| Fund_name | Fund type (Index Funds, Other ETFs, Sectoral/Thematic Funds, etc.) |
| Open_no_of_schemes | Number of open ended schemes launched |
| Open_Funds_mobilized | Funds mobilized for open ended schemes (in crore) |
| Closed_no_of_schemes | Number of closed ended schemes launched |
| Closed_Funds_mobilized | Funds mobilized for closed ended schemes (in crore) |

## Data Cleaning Rules

- All subtotal, total, and grand total rows are removed
- Numbers are stored without commas (numeric datatype)
- For NSR: `-` or NaN values become blank; `0` is kept only when explicitly `0` in the input
- For AUM: NaN values default to `0.00` for numeric fields
- Footnote rows (e.g., "Data in respect Fund of Funds...") are excluded
- Open/Closed ended columns are always present in NSR output (blank if not in input)

## Auto-Detection Logic

The script automatically determines:
1. Whether the file has multiple sheets (combined report) or a single sheet
2. Whether a sheet contains AUM data (looks for "Monthly Report" or "Net Assets Under Management")
3. Whether a sheet contains NSR data (looks for "NEW SCHEMES" in headers)
4. NSR format variant (Sep-style with `*NEW SCHEMES LAUNCHED` marker vs Mar/Apr/Dec style)
5. Whether NSR data maps to Open or Closed ended fields (based on header text)
