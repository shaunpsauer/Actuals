"""
Gas Transmission WBS Operations Dictionary - Standalone Version

This file contains all the WBS operation mappings extracted from the
Gas Transmission WBS and Operations Dictionary. No external Excel file needed!

Last Updated: December 2024
Based on: Gas_Transmission_WBS_and_Operations_Dictionary_REFERENCE.xlsx
"""

# Complete WBS Operations Mapping
# Maps Operation codes to Activity codes with Level 2 and Level 3 WBS hierarchy
OPERATIONS_MAP = {
    1010: {'activity': '0101-1010A', 'l2': '01', 'l3': '01'},
    1020: {'activity': '0101-1020A', 'l2': '01', 'l3': '01'},
    1030: {'activity': '0101-1030A', 'l2': '01', 'l3': '01'},
    1040: {'activity': '0101-1040A', 'l2': '01', 'l3': '01'},
    1100: {'activity': '0102-1100A', 'l2': '01', 'l3': '02'},
    1110: {'activity': '0102-1110A', 'l2': '01', 'l3': '02'},
    1120: {'activity': '0102-1120A', 'l2': '01', 'l3': '02'},
    1130: {'activity': '0102-1130A', 'l2': '01', 'l3': '02'},
    1140: {'activity': '0102-1140A', 'l2': '01', 'l3': '02'},
    1190: {'activity': '0102-1190A', 'l2': '01', 'l3': '02'},
    2010: {'activity': '0201-2010A', 'l2': '02', 'l3': '01'},
    2110: {'activity': '0202-2110A', 'l2': '02', 'l3': '02'},
    2210: {'activity': '0203-2210A', 'l2': '02', 'l3': '03'},
    3010: {'activity': '0301-3010A', 'l2': '03', 'l3': '01'},
    3020: {'activity': '0301-3020A', 'l2': '03', 'l3': '01'},
    3030: {'activity': '0301-3030A', 'l2': '03', 'l3': '01'},
    3100: {'activity': '0302-3100A', 'l2': '03', 'l3': '02'},
    3110: {'activity': '0302-3110A', 'l2': '03', 'l3': '02'},
    3150: {'activity': '0302-3150A', 'l2': '03', 'l3': '02'},
    3210: {'activity': '0303-3210A', 'l2': '03', 'l3': '03'},
    4010: {'activity': '0401-4010A', 'l2': '04', 'l3': '01'},
    4030: {'activity': '0401-4030A', 'l2': '04', 'l3': '01'},
    4040: {'activity': '0401-4040A', 'l2': '04', 'l3': '01'},
    4050: {'activity': '0401-4050A', 'l2': '04', 'l3': '01'},
    4060: {'activity': '0401-4060A', 'l2': '04', 'l3': '01'},
    4070: {'activity': '0401-4070A', 'l2': '04', 'l3': '01'},
    4110: {'activity': '0402-4110A', 'l2': '04', 'l3': '02'},
    4200: {'activity': '0403-4200A', 'l2': '04', 'l3': '03'},
    4210: {'activity': '0403-4210A', 'l2': '04', 'l3': '03'},
    4220: {'activity': '0403-4220A', 'l2': '04', 'l3': '03'},
    5010: {'activity': '0501-5010A', 'l2': '05', 'l3': '01'},
    5020: {'activity': '0501-5020A', 'l2': '05', 'l3': '01'},
    5030: {'activity': '0502-5030A', 'l2': '05', 'l3': '02'},
    5040: {'activity': '0503-5040A', 'l2': '05', 'l3': '03'},
    5050: {'activity': '0503-5050A', 'l2': '05', 'l3': '03'},
    5060: {'activity': '0503-5060A', 'l2': '05', 'l3': '03'},
    5070: {'activity': '0503-5070A', 'l2': '05', 'l3': '03'},
    5080: {'activity': '0503-5080A', 'l2': '05', 'l3': '03'},
    5085: {'activity': '0503-5085A', 'l2': '05', 'l3': '03'},
    5090: {'activity': '0503-5090A', 'l2': '05', 'l3': '03'},
    6000: {'activity': '0504-6000A', 'l2': '05', 'l3': '04'},
    6050: {'activity': '0504-6050A', 'l2': '05', 'l3': '04'},
    6100: {'activity': '0504-6100A', 'l2': '05', 'l3': '04'},
    6200: {'activity': '0504-6200A', 'l2': '05', 'l3': '04'},
    6300: {'activity': '0504-6300A', 'l2': '05', 'l3': '04'},
    6400: {'activity': '0504-6400A', 'l2': '05', 'l3': '04'},
    6500: {'activity': '0504-6500A', 'l2': '05', 'l3': '04'},
    6600: {'activity': '0504-6600A', 'l2': '05', 'l3': '04'},
    6700: {'activity': '0504-6700A', 'l2': '05', 'l3': '04'},
    6800: {'activity': '0504-6800A', 'l2': '05', 'l3': '04'},
    6900: {'activity': '0504-6900A', 'l2': '05', 'l3': '04'},
    7000: {'activity': '0504-7000A', 'l2': '05', 'l3': '04'},
    7100: {'activity': '0504-7100A', 'l2': '05', 'l3': '04'},
    7200: {'activity': '0504-7200A', 'l2': '05', 'l3': '04'},
    7300: {'activity': '0504-7300A', 'l2': '05', 'l3': '04'},
    7400: {'activity': '0504-7400A', 'l2': '05', 'l3': '04'},
    7500: {'activity': '0504-7500A', 'l2': '05', 'l3': '04'},
    7600: {'activity': '0504-7600A', 'l2': '05', 'l3': '04'},
    7700: {'activity': '0504-7700A', 'l2': '05', 'l3': '04'},
    7800: {'activity': '0505-7800A', 'l2': '05', 'l3': '05'},
    7900: {'activity': '0505-7900A', 'l2': '05', 'l3': '05'},
    8000: {'activity': '0505-8000A', 'l2': '05', 'l3': '05'},
    8100: {'activity': '0505-8100A', 'l2': '05', 'l3': '05'},
    8200: {'activity': '0506-8200A', 'l2': '05', 'l3': '06'},
    8300: {'activity': '0506-8300A', 'l2': '05', 'l3': '06'},
    8400: {'activity': '0507-8400A', 'l2': '05', 'l3': '07'},
    8500: {'activity': '0507-8500A', 'l2': '05', 'l3': '07'},
    8600: {'activity': '0507-8600A', 'l2': '05', 'l3': '07'},
    8700: {'activity': '0508-8700A', 'l2': '05', 'l3': '08'},
    8800: {'activity': '0508-8800A', 'l2': '05', 'l3': '08'},
    9010: {'activity': '0601-9010A', 'l2': '06', 'l3': '01'},
    9110: {'activity': '0602-9110A', 'l2': '06', 'l3': '02'},
    9120: {'activity': '0602-9120A', 'l2': '06', 'l3': '02'},
    9130: {'activity': '0602-9130A', 'l2': '06', 'l3': '02'},
}


def build_operations_map():
    """
    Return the complete operations map.
    This function is kept for backward compatibility.
    """
    return OPERATIONS_MAP


def get_activity_code(operation_code):
    """
    Get the activity code for a given operation code.
    
    Args:
        operation_code (int): The operation code (e.g., 1010, 4030)
        
    Returns:
        str: The activity code (e.g., '0101-1010A', '0401-4030A')
        None: If operation code not found
    """
    op_data = OPERATIONS_MAP.get(operation_code)
    if op_data:
        return op_data['activity']
    return None


def get_all_operations():
    """
    Get a list of all valid operation codes.
    
    Returns:
        list: Sorted list of all operation codes
    """
    return sorted(OPERATIONS_MAP.keys())


def is_valid_operation(operation_code):
    """
    Check if an operation code is valid.
    
    Args:
        operation_code (int): The operation code to check
        
    Returns:
        bool: True if valid, False otherwise
    """
    return operation_code in OPERATIONS_MAP


# Module info
__version__ = '2.0.0'
__description__ = 'Standalone WBS Operations Mapper - No Excel file required'
__total_operations__ = len(OPERATIONS_MAP)


if __name__ == '__main__':
    # Test the module
    print(f"WBS Operations Mapper v{__version__}")
    print(f"Total operations loaded: {__total_operations__}")
    print("\nSample mappings:")
    for op in [1010, 4030, 7000, 8300]:
        print(f"  Operation {op} -> Activity {get_activity_code(op)}")
