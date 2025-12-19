"""
SAP Export to HeavyBid Import Transformation Script

This script transforms SAP transaction exports into HeavyBid import format with 3 tabs:
1. Actuals Report - Main import data
2. Actual BoE - Notes for each BidItem/Activity
3. Resource File - Resource definitions

Version: 2.0 - Standalone (no Excel dependencies)
All WBS operations data is embedded in wbs_operations_mapper.py

Usage:
    python sap_to_heavybid.py
    
    The script will guide you through selecting:
    - Your SAP export file (via file picker)
    - Output folder (via folder picker)
    
    The output file will be automatically named <Order>_actuals.xlsx in the selected folder.
    
Requirements:
    - sap_to_heavybid.py (this file)
    - reference_data.py (operations and cost elements dictionary)
    - pandas, openpyxl libraries
"""

import pandas as pd
import numpy as np
from datetime import datetime
from collections import defaultdict
import sys
import os
import time
import tkinter as tk
from tkinter import filedialog

# Import the reference data (no external Excel file needed!)
from reference_data import build_operations_map, build_cost_elements_map


# Cost Element to Resource Code Abbreviation Mapping
COST_ELEMENT_TO_ABBREV = {
    # AFUDC
    5590030: 'AFUDC-Bo',
    5590031: 'AFUDC-Eq',
    
    # Contracts / Overhead
    5091100: 'Meals Ex',
    5091140: 'Reimburs',
    5490000: 'Contract',
    5490003: 'Engr/Dsg',
    5490015: 'Environm',
    
    # Labor Allocation (overhead)
    'Labor Alloc.': 'Labor OH',
    
    # Labor Cost Elements (660xxxx)
    6603001: 'CONSTR',   # Construction
    6603004: 'ACQLIT',   # Acquisition - Misc
    6603005: 'ANLYST',   # Analyst Services
    6603006: 'DRFT',     # Design Drafting Svcs
    6603023: 'ENGSVC',   # Engineering Services
    6603024: 'ENVSVC',   # Environmental Svcs
    6603027: 'ENVPLN',   # Environ Pln & Permit
    6603048: 'PLANSV',   # Planning Services
    6603058: 'TECHSV',   # Technical Services
    6603059: 'LNDENG',   # Land Survey & Engine
    6603082: 'MO-OT',    # Maint & Oper OT Svcs
    6603083: 'MO',       # Maintain & Oper Svc
    6603150: 'ADM-OT',   # Admin Svcs-OT
    6603195: 'CORRSN',   # Corrosion Service
    6603227: 'LNDRTS',   # Land Rights - Misc
    6603823: 'BIOCUL',   # Manage L&EM
    6608158: 'XCON02',   # Contrctr - Consult
    6608160: 'XCON04',   # Contrctr - Engineer
}

# Cost Element to Job Cost Code Mapping
COST_ELEMENT_TO_JOB_COST = {
    5590030: 5590030,  # AFUDC-Borrowed
    5590031: 5590031,  # AFUDC-Equity
    5091100: 5091100,  # Meals Expense
    5091140: 5091140,  # Reimbursed Mileage E
    5490000: 5490000,  # Contracts
    5490003: 5490003,  # Engr/Dsgn & EPC
    5490015: 5490015,  # Environment Contract
    'Labor Alloc.': 'Labor Alloc.',
    # Labor codes map to themselves
    6603001: 6603001,
    6603004: 6603004,
    6603005: 6603005,
    6603006: 6603006,
    6603023: 6603023,
    6603024: 6603024,
    6603027: 6603027,
    6603048: 6603048,
    6603058: 6603058,
    6603059: 6603059,
    6603082: 6603082,
    6603083: 6603083,
    6603150: 6603150,
    6603195: 6603195,
    6603227: 6603227,
    6603823: 6603823,
    6608158: 6608158,
    6608160: 6608160,
}

# Cost Element to Cost Type Mapping
COST_ELEMENT_TO_COST_TYPE = {
    5590030: 'AFUDC',
    5590031: 'AFUDC',
    5091100: 'Contracts',
    5091140: 'Contracts',
    5490000: 'Contracts',
    5490003: 'Contracts',
    5490015: 'Contracts',
    'Labor Alloc.': 'Labor Alloc.',
}

# Default cost type for labor (660xxxx)
LABOR_COST_TYPE = 'Labor'

# Cost Type to Description Prefix Mapping (for Resource File)
COST_TYPE_TO_PREFIX = {
    'AFUDC': 'Actls. - AFUDC - ',
    'Contracts': 'Actls. - Cont. - ',
    'Labor': 'Actls. - Labr. - ',
    'Labor Alloc.': 'Actls. - L.OH. - ',
}
# Default prefix for unknown cost types
DEFAULT_PREFIX = 'Actls. - Other. - '


def read_sap_export(filepath):
    """Read and clean SAP export file"""
    df = pd.read_excel(filepath)
    
    # Remove rows where Order is null (header/summary rows)
    df_clean = df[df['Order'].notna()].copy()
    
    print(f"Loaded {len(df_clean)} rows from SAP export")
    print(f"Order number: {df_clean['Order'].unique()[0]:.0f}")
    
    return df_clean


def normalize_cost_element(cost_element):
    """Normalize cost element to integer for consistent lookups"""
    if pd.isna(cost_element):
        return None
    try:
        # Convert to int (handles float like 5001237.0 -> 5001237)
        return int(float(cost_element))
    except (ValueError, TypeError):
        return None


def generate_resource_code(cost_element, partner_cctr, cost_element_name, cost_elements_map=None):
    """Generate resource code from cost element and partner center"""
    
    # Normalize cost element to int
    ce_int = normalize_cost_element(cost_element)
    
    # Try embedded cost elements map first
    abbrev = None
    if cost_elements_map and ce_int:
        ce_data = cost_elements_map.get(ce_int)
        if ce_data:
            # Use Cost Element Text to derive abbreviation
            # For contract items, use first meaningful word (e.g., "Consulting Services" -> "Consult")
            text = ce_data.get('Cost Element Text', '')
            if text:
                # Split and process words
                words = text.split()
                if words:
                    first_word = words[0].lower()
                    # Map common first words to abbreviations
                    word_mapping = {
                        'consulting': 'Consult',
                        'consult': 'Consult',
                        'engineering': 'Engr',
                        'engineer': 'Engr',
                        'environmental': 'Environ',
                        'environment': 'Environ',
                        'construction': 'Constr',
                        'contract': 'Contract',
                        'meals': 'Meals',
                        'reimbursed': 'Reimburs',
                    }
                    
                    # Check if first word matches a known pattern
                    if first_word in word_mapping:
                        abbrev = word_mapping[first_word]
                    else:
                        # For other words, use first 6-8 chars, properly capitalized
                        # Remove common suffixes like "Services", "Svc", etc.
                        clean_word = first_word
                        for suffix in ['services', 'service', 'svc', 'svcs']:
                            if clean_word.endswith(suffix):
                                clean_word = clean_word[:-len(suffix)]
                                break
                        
                        if clean_word:
                            abbrev = clean_word[:8].capitalize()
                        else:
                            # Fallback: use first word as-is, capitalized
                            abbrev = words[0][:8].capitalize()
    
    # Fall back to hardcoded mapping
    if abbrev is None and ce_int:
        abbrev = COST_ELEMENT_TO_ABBREV.get(ce_int)
    
    # If still no abbreviation, try to create from cost element name
    if abbrev is None:
        # Try to create abbreviation from cost element name
        # For unknown codes, create a simplified abbreviation
        name_parts = str(cost_element_name).upper().split()
        if len(name_parts) >= 2:
            abbrev = ''.join(word[:3] for word in name_parts[:2])
        else:
            abbrev = str(cost_element_name).upper()[:6]
    
    # Build resource code
    if pd.notna(partner_cctr) and partner_cctr > 0:
        # Labor resource with partner center
        resource_code = f"6{abbrev}{int(partner_cctr)}"
    else:
        # Header/contract resource without partner center
        resource_code = f"6{abbrev}"
    
    return resource_code


def aggregate_actuals(df_export, operations_map, cost_elements_map=None):
    """
    Aggregate SAP export data by Operation, Cost Element, and Partner-CCtr
    Returns a DataFrame ready for Actuals Report
    """
    
    # First, calculate AFUDC totals from Operation 1.0 before filtering it out
    afudc_borrowed_total = df_export[
        (df_export['Operation'] == 1.0) & 
        (df_export['Cost Element'] == 5590030.0)
    ]['Val.in rep.cur.'].sum()
    
    afudc_equity_total = df_export[
        (df_export['Operation'] == 1.0) & 
        (df_export['Cost Element'] == 5590031.0)
    ]['Val.in rep.cur.'].sum()
    
    # Calculate Labor OH (overhead) totals per operation (601xxxx cost elements)
    overhead_by_operation = df_export[
        df_export['Cost Element'].astype(str).str.startswith('6010')
    ].groupby('Operation')['Val.in rep.cur.'].sum().to_dict()
    
    # Filter out Operation 1.0 (AFUDC - handled separately)
    df_filtered = df_export[df_export['Operation'] != 1.0].copy()
    
    # Filter out overhead cost elements (601xxxx) - these don't appear in the output
    df_filtered = df_filtered[~df_filtered['Cost Element'].astype(str).str.startswith('6010')].copy()
    
    # Group by Operation, Cost Element, and Partner-CCtr (treat NaN Partner-CCtr as a separate group)
    df_filtered['Partner-CCtr_str'] = df_filtered['Partner-CCtr'].fillna(0).astype(str)
    
    grouped = df_filtered.groupby(['Operation', 'Cost Element', 'Partner-CCtr_str', 'Cost element name']).agg({
        'Total quantity': 'sum',
        'Val.in rep.cur.': 'sum'
    }).reset_index()
    
    # Convert Partner-CCtr back to numeric
    grouped['Partner-CCtr'] = pd.to_numeric(grouped['Partner-CCtr_str'], errors='coerce')
    grouped.drop('Partner-CCtr_str', axis=1, inplace=True)
    
    # Generate resource codes (using embedded cost elements map if available)
    # Create a closure to capture cost_elements_map
    def make_resource_code(row):
        return generate_resource_code(
            row['Cost Element'], 
            row['Partner-CCtr'], 
            row['Cost element name'], 
            cost_elements_map
        )
    
    grouped['Resource'] = grouped.apply(make_resource_code, axis=1)
    
    # Map operations to BidItems and Activities
    grouped['BidItem'] = grouped['Operation'].astype(int)
    grouped['Activity'] = grouped['Operation'].apply(
        lambda op: operations_map.get(int(op), {}).get('activity', f'XXXX-{int(op)}A')
    )
    
    # Add other required columns
    grouped['Quantity'] = grouped['Total quantity']
    grouped['Units'] = 'HR'  # Default unit
    
    # Calculate Unit Price = Val.in rep.cur. / Quantity (handle divide by zero)
    grouped['Unit Price'] = grouped.apply(
        lambda row: row['Val.in rep.cur.'] / row['Quantity'] if row['Quantity'] != 0 else row['Val.in rep.cur.'],
        axis=1
    )
    
    grouped['Tax/OT %'] = 100  # Should be 100, not 1
    grouped['Crew Code'] = np.nan  # Blank as requested
    grouped['Pieces'] = 1
    grouped['Currency'] = np.nan
    grouped['EOE %'] = np.nan
    grouped['Rent Percent'] = np.nan
    grouped['Escalation Percent'] = np.nan
    grouped['Hours Adjustment'] = np.nan
    grouped['Supp. Desc'] = grouped['Cost Element']  # Contains Cost Element code
    grouped['MH/Unit'] = np.nan
    grouped['Material Factor Type'] = np.nan
    grouped['Material Factor'] = np.nan
    grouped['Description'] = grouped['Cost element name']
    
    # Determine Cost Type (using embedded cost elements map if available)
    # Create a closure to capture cost_elements_map
    def make_get_cost_type(cost_elements_map):
        def get_cost_type(cost_element):
            # Normalize to int for lookup
            ce_int = normalize_cost_element(cost_element)
            
            # Try embedded cost elements map first
            if cost_elements_map and ce_int:
                ce_data = cost_elements_map.get(ce_int)
                if ce_data:
                    # Use Level 1 Group or Grouping to determine Cost Type
                    level1_group = ce_data.get('Level 1 Group', '')
                    grouping = ce_data.get('Grouping', '')
                    
                    # Map to Cost Type
                    if level1_group == 'Contract' or grouping == 'Contract':
                        return 'Contracts'
                    elif level1_group == 'Labor' or grouping == 'Labor':
                        return LABOR_COST_TYPE
                    elif level1_group == 'OverHeads' or grouping == 'OverHeads':
                        return 'Labor Alloc.'
                    elif level1_group == 'Materials' or grouping == 'Materials':
                        return 'Other'
            
            # Fall back to hardcoded mapping
            if ce_int and ce_int in COST_ELEMENT_TO_COST_TYPE:
                return COST_ELEMENT_TO_COST_TYPE[ce_int]
            elif ce_int and str(ce_int).startswith('660'):
                return LABOR_COST_TYPE
            elif ce_int and str(ce_int).startswith('50'):
                # Cost elements starting with 50 are typically Contracts
                return 'Contracts'
            else:
                return 'Other'
        return get_cost_type
    
    grouped['Cost Type'] = grouped['Cost Element'].apply(make_get_cost_type(cost_elements_map))
    
    # Set quantity to 1.0 for all non-labor rows (placeholder)
    # AND set Unit Price to the total value for non-labor rows
    non_labor_mask = grouped['Cost Type'] != LABOR_COST_TYPE
    grouped.loc[non_labor_mask, 'Unit Price'] = grouped.loc[non_labor_mask, 'Val.in rep.cur.']
    grouped.loc[non_labor_mask, 'Quantity'] = 1.0
    grouped.loc[non_labor_mask, 'Units'] = 'LS'
    
    # Add Labor Overhead rows for each BidItem/Activity combination
    # NOTE: Labor OH is added to ALL BidItems, even those without Labor!
    labor_oh_rows = []
    
    for (biditem, activity) in grouped[['BidItem', 'Activity']].drop_duplicates().values:
        # Get the overhead value for this operation from the pre-calculated dict
        # Convert biditem to float to match dict keys (Operation is float in df_export)
        operation_float = float(biditem)
        overhead_value = overhead_by_operation.get(operation_float, 0.0)
        
        # Only add Labor OH row if there's actual overhead value (skip zero values and NaN)
        # Use epsilon check for floating-point precision
        if pd.notna(overhead_value) and abs(overhead_value) > 1e-10:
            labor_oh_row = {
                'BidItem': biditem,
                'Activity': activity,
                'Resource': '6Labor OH',
                'Quantity': 1.0,
                'Units': 'LS',
                'Unit Price': overhead_value,  # Use calculated overhead value
                'Tax/OT %': 100,
                'Crew Code': np.nan,
                'Pieces': 1,
                'Currency': np.nan,
                'EOE %': np.nan,
                'Rent Percent': np.nan,
                'Escalation Percent': np.nan,
                'Hours Adjustment': np.nan,
                'Supp. Desc': np.nan,  # Labor Alloc. rows have NaN for Supp. Desc
                'MH/Unit': np.nan,
                'Material Factor Type': np.nan,
                'Material Factor': np.nan,
                'Description': 'Labor Alloc.',
                'Cost Type': 'Labor Alloc.'
            }
            labor_oh_rows.append(labor_oh_row)
    
    if labor_oh_rows:
        labor_oh_df = pd.DataFrame(labor_oh_rows)
        grouped = pd.concat([grouped, labor_oh_df], ignore_index=True)
    
    # Add AFUDC rows ONLY if AFUDC data actually exists in the SAP export
    # Check if totals are non-zero (using epsilon for floating-point comparison)
    afudc_rows = []
    has_afudc_borrowed = pd.notna(afudc_borrowed_total) and abs(afudc_borrowed_total) > 1e-10
    has_afudc_equity = pd.notna(afudc_equity_total) and abs(afudc_equity_total) > 1e-10
    
    if (has_afudc_borrowed or has_afudc_equity) and 1010 in grouped['BidItem'].values:
        # Get the base activity for BidItem 1010
        activity_base = operations_map.get(1010, {}).get('activity', '0101-1010A')
        
        # AFUDC activity changes the operation number's last digit from 0 to 1
        # E.g., 0101-1010A becomes 0101-1011A
        if activity_base[-2] == '0':
            afudc_activity = activity_base[:-2] + '1' + activity_base[-1]
        else:
            afudc_activity = activity_base
        
        # Only add AFUDC-Borrowed if it has a value
        if has_afudc_borrowed:
            afudc_rows.append({
                'BidItem': 1010,
                'Activity': afudc_activity,
                'Resource': '6AFUDC-Bo',
                'Quantity': 1.0,
                'Units': 'LS',
                'Unit Price': afudc_borrowed_total,  # Use calculated AFUDC value
                'Tax/OT %': 100,
                'Crew Code': np.nan,
                'Pieces': 1,
                'Currency': np.nan,
                'EOE %': np.nan,
                'Rent Percent': np.nan,
                'Escalation Percent': np.nan,
                'Hours Adjustment': np.nan,
                'Supp. Desc': 5590030.0,  # Cost Element for AFUDC-Borrowed
                'MH/Unit': np.nan,
                'Material Factor Type': np.nan,
                'Material Factor': np.nan,
                'Description': 'AFUDC-Borrowed',
                'Cost Type': 'AFUDC'
            })
        
        # Only add AFUDC-Equity if it has a value
        if has_afudc_equity:
            afudc_rows.append({
                'BidItem': 1010,
                'Activity': afudc_activity,
                'Resource': '6AFUDC-Eq',
                'Quantity': 1.0,
                'Units': 'LS',
                'Unit Price': afudc_equity_total,  # Use calculated AFUDC value
                'Tax/OT %': 100,
                'Crew Code': np.nan,
                'Pieces': 1,
                'Currency': np.nan,
                'EOE %': np.nan,
                'Rent Percent': np.nan,
                'Escalation Percent': np.nan,
                'Hours Adjustment': np.nan,
                'Supp. Desc': 5590031.0,  # Cost Element for AFUDC-Equity
                'MH/Unit': np.nan,
                'Material Factor Type': np.nan,
                'Material Factor': np.nan,
                'Description': 'AFUDC-Equity',
                'Cost Type': 'AFUDC'
            })
    
    if afudc_rows:
        afudc_df = pd.DataFrame(afudc_rows)
        grouped = pd.concat([grouped, afudc_df], ignore_index=True)
    
    # Sort by BidItem and Activity
    grouped = grouped.sort_values(['BidItem', 'Activity', 'Cost Type', 'Resource'])
    
    # Select and order columns for Actuals Report
    actuals_columns = [
        'BidItem', 'Activity', 'Resource', 'Quantity', 'Units', 'Unit Price', 'Tax/OT %',
        'Crew Code', 'Pieces', 'Currency', 'EOE %', 'Rent Percent', 'Escalation Percent',
        'Hours Adjustment', 'Supp. Desc', 'MH/Unit', 'Material Factor Type', 
        'Material Factor', 'Description', 'Cost Type'
    ]
    
    result = grouped[actuals_columns].copy()
    
    return result


def create_resource_file(df_actuals):
    """Generate Resource File from actuals data"""
    
    # Get unique resources with Cost Type to determine prefix
    # Use subset=['Resource'] to ensure only one row per unique Resource code
    unique_resources = df_actuals[['Resource', 'Description', 'Cost Type']].drop_duplicates(subset=['Resource'], keep='first')
    
    resource_data = []
    
    for _, row in unique_resources.iterrows():
        resource_code = row['Resource']
        cost_type = row['Cost Type']
        
        # For labor resources, use resource code (without "6" prefix) as description
        # For other resources, use the Description from df_actuals
        if cost_type == 'Labor':
            # Remove "6" prefix from resource code
            description = resource_code[1:] if resource_code.startswith('6') else resource_code
        else:
            description = row['Description']
        
        # Get prefix based on Cost Type
        prefix = COST_TYPE_TO_PREFIX.get(cost_type, DEFAULT_PREFIX)
        
        # Apply prefix to description
        prefixed_description = f"{prefix}{description}"
        
        # Extract cost element or type
        is_header = resource_code in ['6AFUDC-Bo', '6AFUDC-Eq', '6Labor OH', '6Meals Ex', 
                                       '6Reimburs', '6Engr/Dsg', '6Environm', '6Contract']
        
        resource_data.append({
            'Local Resource Code': resource_code,
            'Description': prefixed_description,
            'Unit': np.nan,
            'Cost': np.nan,
            'Non-Tax?(Y/N)': np.nan,
            'Job Cost Code 1': np.nan,
            'Job Cost Code 2': np.nan,
            'Job Cost Description': np.nan,
            'Joint Venture Material Type': np.nan,
            'MH/Unit': np.nan,
            'Header Type? (Y/N)': np.nan,
            'Quote Folder': np.nan,
            'Schedule Code': np.nan
        })
    
    df_resource = pd.DataFrame(resource_data)
    
    return df_resource


def create_boe_notes(df_actuals):
    """Generate BoE Notes tab from actuals data - only for activities with Labor rows"""
    
    # Group by BidItem and Activity
    boe_data = []
    
    # Cross-platform date formatting (Windows doesn't support %-m/%-d)
    now = datetime.now()
    current_date = f"{now.month}/{now.day}/{now.strftime('%y')}"
    
    for (biditem, activity), group in df_actuals.groupby(['BidItem', 'Activity']):
        # Only include activities that have Labor rows (exclude AFUDC, Contracts-only, Labor Alloc.-only)
        has_labor = (group['Cost Type'] == 'Labor').any()
        
        if not has_labor:
            continue  # Skip activities without Labor rows
        
        # Build notes string
        notes_lines = [f"{current_date}: "]
        
        # Get labor resources (exclude Labor OH)
        labor_resources = group[
            (group['Cost Type'] == 'Labor') & 
            (~group['Resource'].str.contains('Labor OH', regex=False))
        ]
        
        for _, resource_row in labor_resources.iterrows():
            resource_code = resource_row['Resource']
            # Remove the "6" prefix from resource code for notes
            resource_code_display = resource_code[1:] if resource_code.startswith('6') else resource_code
            quantity = resource_row['Quantity']
            description = resource_row['Description']
            
            # Format quantity: show as integer if it's a whole number, otherwise show with decimals
            if quantity == int(quantity):
                qty_str = str(int(quantity))
            else:
                qty_str = str(quantity)
            
            # Include projection text with 0 instead of ___
            notes_lines.append(f"{resource_code_display}: {qty_str} MH Actuals to date, Projected an additional 0 MH for the remainder of the Activity")
        
        notes = "\n".join(notes_lines)
        
        boe_data.append({
            'BidItem': biditem,
            'Activity': activity,
            'Notes': notes
        })
    
    df_boe = pd.DataFrame(boe_data)
    
    return df_boe


def get_order_from_export(filepath):
    """Extract Order number from SAP export file"""
    df = pd.read_excel(filepath)
    df_clean = df[df['Order'].notna()].copy()
    if len(df_clean) == 0:
        raise ValueError("No valid Order found in SAP export")
    order_num = int(df_clean['Order'].unique()[0])
    return order_num


def generate_output_filename(order_num, output_dir):
    """Generate output filename from Order number with collision-safe suffix in specified directory"""
    base_filename = f"{order_num}_actuals.xlsx"
    output_path = os.path.join(output_dir, base_filename)
    
    # If file exists, add timestamp suffix
    if os.path.exists(output_path):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_filename = f"{order_num}_actuals_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, base_filename)
    
    return output_path


def select_input_file():
    """Open file picker to select SAP export file"""
    # Hide the root tkinter window
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # Open file dialog
    filepath = filedialog.askopenfilename(
        title="Select SAP Export File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    
    root.destroy()
    
    if not filepath:
        print("Canceled - no file selected. Exiting.")
        sys.exit(0)
    
    return filepath


def select_output_folder():
    """Open folder picker to select output directory"""
    # Hide the root tkinter window
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # Open directory dialog
    folderpath = filedialog.askdirectory(
        title="Select Output Folder"
    )
    
    root.destroy()
    
    if not folderpath:
        print("Canceled - no folder selected. Exiting.")
        sys.exit(0)
    
    return folderpath


def transform_sap_to_heavybid(input_file, output_file):
    """Main transformation function"""
    
    banner_width = 80
    title = "SAP EXPORT TO HEAVYBID TRANSFORMATION v2.0"
    padding = (banner_width - len(title)) // 2
    
    print("=" * banner_width)
    print(" " * padding + title)
    print("=" * banner_width)
    
    # Load operations map from embedded data
    print("\nLoading WBS operations map...")
    operations_map = build_operations_map()
    print(f"Loaded {len(operations_map)} operation mappings")
    
    # Load cost elements map from embedded data
    print("\nLoading cost elements map...")
    cost_elements_map = build_cost_elements_map()
    print(f"Loaded {len(cost_elements_map)} cost element mappings")
    
    # Read SAP export
    print(f"\nReading SAP export: {input_file}")
    df_export = read_sap_export(input_file)
    
    # Transform to actuals report
    print("\nAggregating actuals...")
    df_actuals = aggregate_actuals(df_export, operations_map, cost_elements_map)
    print(f"Generated {len(df_actuals)} actuals rows")
    
    # Create resource file
    print("\nCreating resource file...")
    df_resource = create_resource_file(df_actuals)
    print(f"Generated {len(df_resource)} resource definitions")
    
    # Create BoE notes
    print("\nCreating BoE notes...")
    df_boe = create_boe_notes(df_actuals)
    print(f"Generated {len(df_boe)} BoE note entries")
    
    # Write to Excel with 3 tabs
    print(f"\nWriting output to: {output_file}")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_actuals.to_excel(writer, sheet_name='Actuals Report', index=False)
        df_boe.to_excel(writer, sheet_name='Actual BoE', index=False)
        df_resource.to_excel(writer, sheet_name='Resource File', index=False)
    
    # Display transformation complete banner
    banner_width = 80
    title = "TRANSFORMATION COMPLETE"
    padding = (banner_width - len(title)) // 2
    
    print("\n" + "=" * banner_width)
    print(" " * padding + title)
    print("=" * banner_width)
    print(f"\nOutput file created: {output_file}")
    print(f"  - Actuals Report: {len(df_actuals)} rows")
    print(f"  - Actual BoE: {len(df_boe)} rows")
    print(f"  - Resource File: {len(df_resource)} rows")


if __name__ == '__main__':
    # Display welcome banner
    banner_width = 80
    title = "SAP Actuals to HeavyBid"
    padding = (banner_width - len(title)) // 2
    
    print("=" * banner_width)
    print(" " * padding + title)
    print("=" * banner_width)
    print("\nThis tool will transform your SAP export into HeavyBid import format.")
    print("You'll be prompted to select your SAP export file and output folder.")
    print("The output file will contain 3 sheets: Actuals Report, Actual BoE, and Resource File.")
    print("\nYou can cancel at any time by closing the file picker dialogs.\n")
    
    # Wait 3 seconds before opening file picker
    print("Opening file picker in 3 seconds...")
    time.sleep(3)
    
    # Step 1: Select input file
    print("Step 1: Select SAP export file...")
    input_file = select_input_file()
    print(f"✓ Selected: {input_file}\n")
    
    # Extract Order number from the selected file
    print("Extracting Order number from export file...")
    try:
        order_num = get_order_from_export(input_file)
        print(f"✓ Order number: {order_num}\n")
    except Exception as e:
        print(f"Error extracting Order number: {e}")
        sys.exit(1)
    
    # Step 2: Select output folder
    print("Step 2: Select output folder...")
    output_folder = select_output_folder()
    print(f"✓ Selected folder: {output_folder}\n")
    
    # Generate output filename in the selected folder
    output_file = generate_output_filename(order_num, output_folder)
    print(f"Output file will be saved as: {os.path.basename(output_file)}")
    print(f"Full path: {output_file}\n")
    
    # Run transformation
    transform_sap_to_heavybid(input_file, output_file)
