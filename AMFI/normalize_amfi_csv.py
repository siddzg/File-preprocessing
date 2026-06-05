import pandas as pd
import sys
import os

# Take file path as input argument
if len(sys.argv) < 2:
    print("Usage: python normalize_amfi.py <path_to_excel_file>")
    sys.exit(1)

file_path = sys.argv[1]

# Determine engine based on file extension
ext = os.path.splitext(file_path)[1].lower()
if ext == '.xls':
    engine = 'xlrd'
elif ext == '.xlsx':
    engine = 'openpyxl'
else:
    print(f"Unsupported file format: {ext}. Use .xls or .xlsx")
    sys.exit(1)

output_dir = os.path.dirname(os.path.abspath(file_path))


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def clean_numeric_allow_dash(val):
    """Return the value as number, empty for NaN or dash, 0 only if explicitly 0 in input"""
    if pd.isna(val):
        return ''
    val_str = str(val).strip()
    if val_str.replace(' ', '') == '-' or val_str.strip() == '-':
        return ''
    try:
        num = float(val_str.replace(',', '').strip())
        return int(num) if num == int(num) else num
    except (ValueError, TypeError):
        return ''


def clean_numeric_aum(val):
    """For AUM data - return 0.00 for NaN, round to 2 decimals"""
    if pd.isna(val):
        return 0.00
    try:
        return round(float(val), 2)
    except (ValueError, TypeError):
        try:
            return round(float(str(val).replace(',', '')), 2)
        except (ValueError, TypeError):
            return 0.00


def clean_int_aum(val):
    """For AUM data - return 0 for NaN, integer value"""
    if pd.isna(val):
        return 0
    try:
        return int(float(str(val).replace(',', '').replace('##', '')))
    except (ValueError, TypeError):
        return 0


# ============================================================
# MF AUM NORMALIZATION
# ============================================================

def parse_mf_aum(df):
    """Parse MF AUM sheet and return normalized DataFrame"""
    df.columns = ['Sr', 'Scheme_Name', 'No_of_Schemes', 'No_of_Folios',
                  'Funds_Mobilized', 'Repurchase_Redemption', 'Net_Inflow_Outflow',
                  'Net_AUM', 'Average_Net_AUM', 'No_of_segregated_portfolios',
                  'Net_AUM_in_segregated_portfolio']

    # Skip header rows (rows 0, 1, 2)
    df = df.iloc[3:].reset_index(drop=True)

    scheme_type_map = {'A': 'Open ended Schemes', 'B': 'Close Ended Schemes', 'C': 'Interval Schemes'}
    sub_scheme_roman = ['I', 'II', 'III', 'IV', 'V']

    results = []
    current_scheme_type = None
    current_sub_scheme = None

    for idx, row in df.iterrows():
        sr = str(row['Sr']).strip() if pd.notna(row['Sr']) else ''
        scheme_name = str(row['Scheme_Name']).strip() if pd.notna(row['Scheme_Name']) else ''
        no_of_schemes = row['No_of_Schemes']

        if sr == '' and scheme_name == '' and pd.isna(no_of_schemes):
            continue

        if 'Sub Total' in scheme_name or 'Total A' in scheme_name or \
           'Total B' in scheme_name or 'Total C' in scheme_name or \
           'Grand Total' in scheme_name:
            continue

        if '**' in scheme_name and 'Data in respect' in scheme_name:
            continue
        if '##' in scheme_name and 'Include NFOs' in scheme_name:
            continue

        if sr in scheme_type_map:
            current_scheme_type = scheme_type_map[sr]
            current_sub_scheme = None
            continue

        if sr in sub_scheme_roman and pd.isna(no_of_schemes):
            current_sub_scheme = scheme_name
            continue

        if pd.notna(no_of_schemes):
            fund_name = scheme_name

            if 'Fund of Funds Scheme (Domestic)' in scheme_name:
                current_scheme_type = 'Fund of Funds Scheme (Domestic)'
                current_sub_scheme = ''
                fund_name = ''
                no_of_schemes_val = str(row['No_of_Schemes']).replace('##', '').strip()
                try:
                    no_of_schemes_val = int(float(no_of_schemes_val))
                except (ValueError, TypeError):
                    no_of_schemes_val = row['No_of_Schemes']
            else:
                no_of_schemes_val = row['No_of_Schemes']

            if sr in sub_scheme_roman and pd.notna(no_of_schemes):
                current_sub_scheme = scheme_name
                fund_name = ''

            if sr == 'III' and pd.notna(no_of_schemes) and current_scheme_type in ['Close Ended Schemes', 'Interval Schemes']:
                current_sub_scheme = scheme_name
                fund_name = ''

            fund_name = fund_name.strip() if fund_name else ''

            results.append({
                'Scheme_Category': current_scheme_type if current_scheme_type else '',
                'Sub_Scheme_Category': current_sub_scheme if current_sub_scheme else '',
                'Fund_Name': fund_name,
                'No_of_Schemes': clean_int_aum(no_of_schemes_val),
                'No_of_Folios': clean_int_aum(row['No_of_Folios']),
                'Funds_Mobilized': clean_numeric_aum(row['Funds_Mobilized']),
                'Repurchase_Redemption': clean_numeric_aum(row['Repurchase_Redemption']),
                'Net_Inflow_Outflow': clean_numeric_aum(row['Net_Inflow_Outflow']),
                'Net_AUM': clean_numeric_aum(row['Net_AUM']),
                'Average_Net_AUM': clean_numeric_aum(row['Average_Net_AUM']),
                'No_of_segregated_portfolios': clean_numeric_aum(row['No_of_segregated_portfolios']),
                'Net_AUM_in_segregated_portfolio': clean_numeric_aum(row['Net_AUM_in_segregated_portfolio'])
            })

    output_df = pd.DataFrame(results)
    output_df['Scheme_Category'] = output_df['Scheme_Category'].fillna('')
    output_df['Sub_Scheme_Category'] = output_df['Sub_Scheme_Category'].fillna('')
    output_df['Fund_Name'] = output_df['Fund_Name'].fillna('')
    return output_df


# ============================================================
# NSR NORMALIZATION
# ============================================================

def detect_nsr_format(df):
    """Detect if NSR file is Sep-style (has *NEW SCHEMES LAUNCHED section)"""
    for idx, row in df.iterrows():
        for col in df.columns:
            val = str(row[col]).strip() if pd.notna(row[col]) else ''
            if val.startswith('*NEW SCHEMES LAUNCHED') or val.startswith('*New Schemes Launched'):
                return 'sep_style'
    return 'mar_apr_style'


def parse_nsr_mar_apr_style(df):
    """Parse Mar/Apr/Dec style NSR files"""
    num_cols = df.shape[1]

    # Detect if header indicates Open-ended or Close-ended only
    scheme_type = 'both'
    for idx in range(min(6, len(df))):
        for col in range(num_cols):
            val = str(df.iloc[idx, col]).strip() if pd.notna(df.iloc[idx, col]) else ''
            if 'Open-ended' in val or 'Open ended' in val or 'Open Ended' in val:
                if 'Name of the Scheme' in val:
                    scheme_type = 'open_only'
            elif 'Close-ended' in val or 'Close ended' in val or 'Close Ended' in val:
                if 'Name of the Scheme' in val:
                    scheme_type = 'close_only'

    has_closed_cols = num_cols >= 6

    if has_closed_cols:
        df.columns = ['Category', 'Scheme_Name', 'Open_No_of_Schemes', 'Open_Funds_Mobilized',
                      'Closed_No_of_Schemes', 'Closed_Funds_Mobilized'][:num_cols]
    else:
        df.columns = ['Category', 'Scheme_Name', 'No_of_Schemes', 'No_Funds_Mobilized'][:num_cols]

    df = df.iloc[3:].reset_index(drop=True)

    results = []
    current_category = None
    current_fund_name = None

    for idx, row in df.iterrows():
        cat_val = str(row['Category']).strip() if pd.notna(row['Category']) else ''
        scheme_name = str(row['Scheme_Name']).strip() if pd.notna(row['Scheme_Name']) else ''

        if cat_val == '' and scheme_name == '':
            continue

        if 'Sub total' in cat_val or 'Sub Total' in cat_val or 'Subtotal' in cat_val or 'Grand Total' in cat_val or 'Total' in cat_val:
            continue

        is_category_header = False
        if cat_val and cat_val[0].isalpha() and len(cat_val) > 2 and cat_val[1] == '.':
            current_category = cat_val[2:].strip()
            current_fund_name = None
            is_category_header = True

        if is_category_header:
            continue

        if has_closed_cols:
            open_val = str(row['Open_No_of_Schemes']).strip() if pd.notna(row['Open_No_of_Schemes']) else ''
        else:
            open_val = str(row['No_of_Schemes']).strip() if pd.notna(row['No_of_Schemes']) else ''
        if 'No. of Schemes' in open_val or 'No. of Scheme' in open_val:
            continue

        if 'Name of the Scheme' in scheme_name:
            continue

        if cat_val != '':
            current_fund_name = cat_val

        if scheme_name != '':
            if has_closed_cols:
                open_schemes = clean_numeric_allow_dash(row['Open_No_of_Schemes'])
                open_funds = clean_numeric_allow_dash(row['Open_Funds_Mobilized'])
                closed_schemes = clean_numeric_allow_dash(row['Closed_No_of_Schemes'])
                closed_funds = clean_numeric_allow_dash(row['Closed_Funds_Mobilized'])
            else:
                num_schemes = clean_numeric_allow_dash(row['No_of_Schemes'])
                funds_mobilized = clean_numeric_allow_dash(row['No_Funds_Mobilized'])

                if scheme_type == 'open_only':
                    open_schemes = num_schemes
                    open_funds = funds_mobilized
                    closed_schemes = ''
                    closed_funds = ''
                elif scheme_type == 'close_only':
                    open_schemes = ''
                    open_funds = ''
                    closed_schemes = num_schemes
                    closed_funds = funds_mobilized
                else:
                    open_schemes = num_schemes
                    open_funds = funds_mobilized
                    closed_schemes = ''
                    closed_funds = ''

            results.append({
                'Scheme_Category': current_category if current_category else '',
                'Scheme_Name': scheme_name,
                'Fund_name': current_fund_name if current_fund_name else '',
                'Open_no_of_schemes': open_schemes,
                'Open_Funds_mobilized': open_funds,
                'Closed_no_of_schemes': closed_schemes,
                'Closed_Funds_mobilized': closed_funds,
            })

    return results


def parse_nsr_sep_style(df):
    """Parse Sep-style NSR files with top data section and bottom *NEW SCHEMES LAUNCHED section"""
    num_cols = df.shape[1]

    # Find the row where *NEW SCHEMES LAUNCHED starts
    new_schemes_row = None
    for idx, row in df.iterrows():
        for col in df.columns:
            val = str(row[col]).strip() if pd.notna(row[col]) else ''
            if val.startswith('*NEW SCHEMES LAUNCHED') or val.startswith('*New Schemes Launched'):
                new_schemes_row = idx
                break
        if new_schemes_row is not None:
            break

    # Parse the top section (fund-level data)
    fund_data = {}
    current_category = None

    for idx in range(0, new_schemes_row if new_schemes_row else len(df)):
        row = df.iloc[idx]
        col0 = str(row[0]).strip() if pd.notna(row[0]) else ''

        if col0 == '':
            continue
        if 'NEW SCHEMES' in col0.upper() or 'Rs.' in col0:
            continue
        if col0 in ['Open End', 'Close End', 'Total', 'No. of Schemes', 'Funds mobilized']:
            continue
        if 'Subtotal' in col0 or 'Sub Total' in col0 or 'Sub total' in col0 or 'Total' in col0:
            continue

        if col0 and col0[0].isalpha() and len(col0) > 2 and col0[1] == '.':
            current_category = col0[2:].strip()
            continue

        if col0 != '' and current_category:
            fund_name = col0
            open_schemes = clean_numeric_allow_dash(row[1]) if num_cols > 1 else ''
            open_funds = clean_numeric_allow_dash(row[2]) if num_cols > 2 else ''
            closed_schemes = clean_numeric_allow_dash(row[3]) if num_cols > 3 else ''
            closed_funds = clean_numeric_allow_dash(row[4]) if num_cols > 4 else ''

            fund_data[(current_category, fund_name)] = {
                'open_schemes': open_schemes,
                'open_funds': open_funds,
                'closed_schemes': closed_schemes,
                'closed_funds': closed_funds
            }

    # Parse the bottom section (*NEW SCHEMES LAUNCHED)
    results = []
    current_category = None
    current_fund_name = None

    for idx in range(new_schemes_row, len(df)):
        row = df.iloc[idx]
        col0 = str(row[0]).strip() if pd.notna(row[0]) else ''
        col1 = str(row[1]).strip() if pd.notna(row[1]) else ''

        if '*NEW SCHEMES LAUNCHED' in col0 or 'NEW SCHEMES LAUNCHED' in col0.upper():
            continue
        if 'Open End Scheme' in col0 or 'Close End Scheme' in col0:
            continue
        if col0 == '' and col1 == '':
            continue

        if col0 and col0[0].isalpha() and len(col0) > 2 and col0[1] == '.':
            current_category = col0[2:].strip()
            current_fund_name = None
            continue

        if col0 != '':
            current_fund_name = col0
            if col1 != '':
                data = fund_data.get((current_category, current_fund_name), {})
                results.append({
                    'Scheme_Category': current_category if current_category else '',
                    'Scheme_Name': col1,
                    'Fund_name': current_fund_name if current_fund_name else '',
                    'Open_no_of_schemes': data.get('open_schemes', ''),
                    'Open_Funds_mobilized': data.get('open_funds', ''),
                    'Closed_no_of_schemes': data.get('closed_schemes', ''),
                    'Closed_Funds_mobilized': data.get('closed_funds', ''),
                })
        elif col1 != '':
            data = fund_data.get((current_category, current_fund_name), {})
            results.append({
                'Scheme_Category': current_category if current_category else '',
                'Scheme_Name': col1,
                'Fund_name': current_fund_name if current_fund_name else '',
                'Open_no_of_schemes': data.get('open_schemes', ''),
                'Open_Funds_mobilized': data.get('open_funds', ''),
                'Closed_no_of_schemes': data.get('closed_schemes', ''),
                'Closed_Funds_mobilized': data.get('closed_funds', ''),
            })

    return results


def parse_nsr(df):
    """Parse NSR data and return normalized DataFrame"""
    file_format = detect_nsr_format(df)

    if file_format == 'sep_style':
        results = parse_nsr_sep_style(df)
    else:
        results = parse_nsr_mar_apr_style(df)

    output_df = pd.DataFrame(results)

    for col in ['Scheme_Category', 'Scheme_Name', 'Fund_name', 'Open_no_of_schemes',
                'Open_Funds_mobilized', 'Closed_no_of_schemes', 'Closed_Funds_mobilized']:
        if col not in output_df.columns:
            output_df[col] = ''

    output_df = output_df[['Scheme_Category', 'Scheme_Name', 'Fund_name', 'Open_no_of_schemes',
                            'Open_Funds_mobilized', 'Closed_no_of_schemes', 'Closed_Funds_mobilized']]
    return output_df


# ============================================================
# MAIN LOGIC - Detect single sheet vs multi-sheet
# ============================================================

xls = pd.ExcelFile(file_path, engine=engine)
sheet_names = xls.sheet_names

# Detect if file has multiple sheets (combined file like amsep2025repo.xls)
# or a single sheet (individual files like MF_AUM_AMFI.xls or Sample_NSR_input_*.xls)

aum_sheet = None
nsr_sheet = None

if len(sheet_names) > 1:
    # Multi-sheet file: identify which sheet is AUM and which is NSR
    for sheet in sheet_names:
        df_check = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=10)
        is_aum = False
        is_nsr = False
        for idx, row in df_check.iterrows():
            for col in df_check.columns:
                val = str(row[col]).strip() if pd.notna(row[col]) else ''
                if 'Monthly Report' in val or 'Net Assets Under Management' in val:
                    is_aum = True
                if 'NEW SCHEMES' in val.upper() or 'New Schemes' in val:
                    is_nsr = True
        if is_aum:
            aum_sheet = sheet
        if is_nsr:
            nsr_sheet = sheet

    if aum_sheet:
        print(f"\n{'='*60}")
        print(f"Processing AUM data from sheet: '{aum_sheet}'")
        print(f"{'='*60}")
        df_aum = pd.read_excel(xls, sheet_name=aum_sheet, header=None)
        aum_df = parse_mf_aum(df_aum)
        aum_output_path = os.path.join(output_dir, 'Normalized_MF_AUM.csv')
        aum_df.to_csv(aum_output_path, index=False)
        print(f"Saved to: {aum_output_path}")
        print(f"Total rows: {len(aum_df)}")
        print(aum_df.head(5).to_string())

    if nsr_sheet:
        print(f"\n{'='*60}")
        print(f"Processing NSR data from sheet: '{nsr_sheet}'")
        print(f"{'='*60}")
        df_nsr = pd.read_excel(xls, sheet_name=nsr_sheet, header=None)
        nsr_df = parse_nsr(df_nsr)
        nsr_output_path = os.path.join(output_dir, 'Normalized_NSR.csv')
        nsr_df.to_csv(nsr_output_path, index=False)
        print(f"Saved to: {nsr_output_path}")
        print(f"Total rows: {len(nsr_df)}")
        print(nsr_df.to_string())

    if not aum_sheet and not nsr_sheet:
        print("Could not identify AUM or NSR sheets in the file.")
        sys.exit(1)

else:
    # Single-sheet file: detect if it's AUM or NSR based on content
    df = pd.read_excel(xls, sheet_name=sheet_names[0], header=None)

    is_aum = False
    for idx in range(min(5, len(df))):
        for col in df.columns:
            val = str(df.iloc[idx, col]).strip() if pd.notna(df.iloc[idx, col]) else ''
            if 'Monthly Report' in val or 'Net Assets Under Management' in val:
                is_aum = True
                break

    if is_aum:
        print("Detected: MF AUM data")
        aum_df = parse_mf_aum(df)
        aum_output_path = os.path.join(output_dir, 'Normalized_MF_AUM.csv')
        aum_df.to_csv(aum_output_path, index=False)
        print(f"Saved to: {aum_output_path}")
        print(f"Total rows: {len(aum_df)}")
        print(aum_df.to_string())
    else:
        print("Detected: NSR data")
        nsr_df = parse_nsr(df)
        nsr_output_path = os.path.join(output_dir, 'Normalized_NSR.csv')
        nsr_df.to_csv(nsr_output_path, index=False)
        print(f"Saved to: {nsr_output_path}")
        print(f"Total rows: {len(nsr_df)}")
        print(nsr_df.to_string())
