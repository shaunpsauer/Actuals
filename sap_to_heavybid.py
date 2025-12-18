"""
SAP Export to HeavyBid Import Transformation Script

This script transforms SAP transaction exports into HeavyBid import format with 3 tabs:
1. Actuals Report - Main import data
2. Actual BoE - Notes for each BidItem/Activity
3. Resource File - Resource definitions

Version: 2.0 - Standalone (no Excel dependencies)
All WBS operations data is embedded in wbs_operations_mapper.py

Usage:
    python sap_to_heavybid.py input_export.xlsx output_file.xlsx
    
Requirements:
    - sap_to_heavybid.py (this file)
    - wbs_operations_mapper.py (operations dictionary)
    - pandas, openpyxl libraries
"""

import pandas as pd
import numpy as np
from datetime import datetime
from collections import defaultdict
import sys
import os
import tkinter as tk
from tkinter import filedialog

# Import the operations mapper (no external Excel file needed!)
from wbs_operations_mapper import build_operations_map


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


def read_sap_export(filepath):
    """Read and clean SAP export file"""
    df = pd.read_excel(filepath)
    
    # Remove rows where Order is null (header/summary rows)
    df_clean = df[df['Order'].notna()].copy()
    
    print(f"Loaded {len(df_clean)} rows from SAP export")
    print(f"Order number: {df_clean['Order'].unique()[0]:.0f}")
    
    return df_clean


def generate_resource_code(cost_element, partner_cctr, cost_element_name):
    """Generate resource code from cost element and partner center"""
    
    # Get abbreviation
    abbrev = COST_ELEMENT_TO_ABBREV.get(cost_element)
    
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


def aggregate_actuals(df_export, operations_map):
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
    
    # Generate resource codes
    grouped['Resource'] = grouped.apply(
        lambda row: generate_resource_code(row['Cost Element'], row['Partner-CCtr'], row['Cost element name']),
        axis=1
    )
    
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
    
    # Determine Cost Type
    def get_cost_type(cost_element):
        if cost_element in COST_ELEMENT_TO_COST_TYPE:
            return COST_ELEMENT_TO_COST_TYPE[cost_element]
        elif str(cost_element).startswith('660'):
            return LABOR_COST_TYPE
        else:
            return 'Other'
    
    grouped['Cost Type'] = grouped['Cost Element'].apply(get_cost_type)
    
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
        overhead_value = overhead_by_operation.get(biditem, 0.0)
        
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
    
    # Add AFUDC rows ONLY for BidItem 1010
    afudc_rows = []
    if 1010 in grouped['BidItem'].values:
        # Get the base activity for BidItem 1010
        activity_base = operations_map.get(1010, {}).get('activity', '0101-1010A')
        
        # AFUDC activity changes the operation number's last digit from 0 to 1
        # E.g., 0101-1010A becomes 0101-1011A
        if activity_base[-2] == '0':
            afudc_activity = activity_base[:-2] + '1' + activity_base[-1]
        else:
            afudc_activity = activity_base
        
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
    
    # Get unique resources
    unique_resources = df_actuals[['Resource', 'Description']].drop_duplicates()
    
    resource_data = []
    
    for _, row in unique_resources.iterrows():
        resource_code = row['Resource']
        description = row['Description']
        
        # Extract cost element or type
        is_header = resource_code in ['6AFUDC-Bo', '6AFUDC-Eq', '6Labor OH', '6Meals Ex', 
                                       '6Reimburs', '6Engr/Dsg', '6Environm', '6Contract']
        
        resource_data.append({
            'Local Resource Code': resource_code,
            'Description': description,
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


def generate_output_filename(input_file, order_num):
    """Generate output filename from Order number with collision-safe suffix"""
    base_filename = f"{order_num}_actuals.xlsx"
    output_path = os.path.join(os.getcwd(), base_filename)
    
    # If file exists, add timestamp suffix
    if os.path.exists(output_path):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_filename = f"{order_num}_actuals_{timestamp}.xlsx"
        output_path = os.path.join(os.getcwd(), base_filename)
    
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
        print("No file selected. Exiting.")
        sys.exit(0)
    
    return filepath


def transform_sap_to_heavybid(input_file, output_file):
    """Main transformation function"""
    
    print("=" * 100)
    print("SAP EXPORT TO HEAVYBID TRANSFORMATION v2.0")
    print("=" * 100)
    
    # Load operations map from embedded data (no Excel file needed!)
    print("\nLoading WBS operations map...")
    operations_map = build_operations_map()
    print(f"Loaded {len(operations_map)} operation mappings")
    
    # Read SAP export
    print(f"\nReading SAP export: {input_file}")
    df_export = read_sap_export(input_file)
    
    # Transform to actuals report
    print("\nAggregating actuals...")
    df_actuals = aggregate_actuals(df_export, operations_map)
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
    
    print("\n" + "=" * 100)
    print("TRANSFORMATION COMPLETE!")
    print("=" * 100)
    print(f"\nOutput file created: {output_file}")
    print(f"  - Actuals Report: {len(df_actuals)} rows")
    print(f"  - Actual BoE: {len(df_boe)} rows")
    print(f"  - Resource File: {len(df_resource)} rows")


if __name__ == '__main__':
    if len(sys.argv) == 1:
        # No arguments: open file picker and auto-generate output filename
        print("No arguments provided. Opening file picker...")
        input_file = select_input_file()
        
        # Extract Order number and generate output filename
        print(f"\nExtracting Order number from: {input_file}")
        try:
            order_num = get_order_from_export(input_file)
            output_file = generate_output_filename(input_file, order_num)
            print(f"Order number: {order_num}")
            print(f"Output file: {output_file}")
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)
        
        transform_sap_to_heavybid(input_file, output_file)
        
    elif len(sys.argv) == 3:
        # Two arguments: existing CLI behavior
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        transform_sap_to_heavybid(input_file, output_file)
        
    else:
        # Invalid argument count
        print("Usage: python sap_to_heavybid.py [<input_export.xlsx> <output_file.xlsx>]")
        print("\nIf no arguments provided, a file picker will open to select the SAP export.")
        print("Output will be automatically named <Order>_actuals.xlsx in the current directory.")
        sys.exit(1)
