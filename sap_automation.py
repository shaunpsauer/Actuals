"""
SAP GUI Automation Module

This module provides functions to automate SAP GUI interactions for exporting data.
Uses Windows COM automation (pywin32) to interact with SAP GUI scripting interface.

Requirements:
    - pywin32 (pip install pywin32)
    - SAP GUI client installed and configured
    - User must be logged into SAP (or module can launch SAP Logon)
"""

import win32com.client
import time
import os
import subprocess
import sys
from pathlib import Path


def check_sap_connection():
    """
    Check if SAP GUI is running and user is logged in.
    
    Returns:
        tuple: (is_connected: bool, session: object or None, error_message: str or None)
    """
    try:
        # Try to get SAP GUI application
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        
        # Check if there are any connections
        if application.Connections.Count == 0:
            return False, None, "SAP GUI is running but no connections found. Please log in to SAP."
        
        # Get first connection
        connection = application.Connections(0)
        
        # Check if there are any sessions
        if connection.Sessions.Count == 0:
            return False, None, "SAP connection exists but no active sessions. Please log in to SAP."
        
        # Get first session
        session = connection.Sessions(0)
        
        # Try to access session info to verify it's active
        try:
            session_info = session.Info
            return True, session, None
        except:
            return False, None, "SAP session exists but appears inactive. Please ensure you are logged in."
            
    except Exception as e:
        error_msg = str(e)
        if "GetObject" in error_msg or "SAPGUI" in error_msg:
            return False, None, "SAP GUI is not running. Please launch SAP Logon and log in."
        return False, None, f"Error checking SAP connection: {error_msg}"


def launch_sap_gui():
    """
    Launch SAP Logon application.
    
    Returns:
        bool: True if launch was successful, False otherwise
    """
    try:
        # Common SAP Logon paths
        sap_paths = [
            r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe",
            r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
            os.path.expanduser(r"~\AppData\Local\SAP\SAP GUI\saplogon.exe"),
        ]
        
        # Try to find and launch SAP Logon
        for path in sap_paths:
            if os.path.exists(path):
                subprocess.Popen([path])
                return True
        
        # If not found in common locations, try to find it
        try:
            # Try using Windows search
            result = subprocess.run(
                ["where", "saplogon.exe"],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0 and result.stdout.strip():
                sap_path = result.stdout.strip().split('\n')[0]
                subprocess.Popen([sap_path])
                return True
        except:
            pass
        
        return False
        
    except Exception as e:
        print(f"Error launching SAP GUI: {e}")
        return False


def wait_for_sap_login(timeout=300):
    """
    Wait for user to log into SAP.
    
    Args:
        timeout: Maximum time to wait in seconds (default: 5 minutes)
    
    Returns:
        tuple: (success: bool, session: object or None, error_message: str or None)
    """
    print("Waiting for SAP login...")
    print("Please log in to SAP when prompted. Press Enter here when you have successfully logged in.")
    
    start_time = time.time()
    check_interval = 2  # Check every 2 seconds
    
    while time.time() - start_time < timeout:
        is_connected, session, error_msg = check_sap_connection()
        
        if is_connected:
            print("✓ SAP connection detected!")
            return True, session, None
        
        time.sleep(check_interval)
    
    return False, None, "Timeout waiting for SAP login. Please ensure you are logged in and try again."


def get_sap_session():
    """
    Get the active SAP session via COM automation.
    
    Returns:
        tuple: (success: bool, session: object or None, error_message: str or None)
    """
    is_connected, session, error_msg = check_sap_connection()
    
    if is_connected:
        return True, session, None
    else:
        return False, None, error_msg


def prompt_sap_parameters():
    """
    Prompt user for SAP export parameters.
    
    Returns:
        dict: Dictionary with keys: order_num, controlling_area, date_from, date_to
    """
    print("\n" + "=" * 80)
    print("SAP Export Parameters")
    print("=" * 80)
    
    # Order number
    order_num = input("Enter Order number: ").strip()
    while not order_num:
        print("Order number is required.")
        order_num = input("Enter Order number: ").strip()
    
    # Controlling Area (default: ORDFIN)
    controlling_area = input("Enter Controlling Area [ORDFIN]: ").strip()
    if not controlling_area:
        controlling_area = "ORDFIN"
    
    # Date From (default: 01/01/2010)
    date_from = input("Enter Date From [01/01/2010]: ").strip()
    if not date_from:
        date_from = "01/01/2010"
    
    # Date To (default: 12/22/2025)
    date_to = input("Enter Date To [12/22/2025]: ").strip()
    if not date_to:
        date_to = "12/22/2025"
    
    return {
        'order_num': order_num,
        'controlling_area': controlling_area,
        'date_from': date_from,
        'date_to': date_to
    }


def navigate_to_transaction(session, transaction_code):
    """
    Navigate to a specific SAP transaction.
    
    Args:
        session: SAP session object
        transaction_code: Transaction code (e.g., "KOB1")
    
    Returns:
        tuple: (success: bool, error_message: str or None)
    """
    # Method 1: Try using StartTransaction (preferred method)
    try:
        session.StartTransaction(transaction_code)
        time.sleep(1.5)  # Wait for transaction to load
        return True, None
    except:
        pass
    
    # Method 2: Alternative - use transaction input field
    try:
        # Clear any existing input in the transaction field
        okcd_field = session.FindById("wnd[0]/tbar[0]/okcd")
        okcd_field.Text = transaction_code
        session.FindById("wnd[0]").SendVKey(0)  # Press Enter
        time.sleep(1.5)  # Wait for transaction to load
        return True, None
    except Exception as e2:
        return False, f"Could not navigate to transaction {transaction_code}. Error: {str(e2)}. Please ensure you are logged into SAP and try again."


def execute_sap_export(order_num, controlling_area, date_from, date_to, output_path=None, transaction_code="KOB1"):
    """
    Execute the SAP export workflow based on the VBA script.
    
    Args:
        order_num: Order number (e.g., "74066927")
        controlling_area: Controlling Area (e.g., "ORDFIN")
        date_from: Start date in MM/DD/YYYY format (e.g., "01/01/2010")
        date_to: End date in MM/DD/YYYY format (e.g., "12/22/2025")
        output_path: Optional path to save the exported file. If None, uses temp directory.
        transaction_code: SAP transaction code (default: "KOB1")
    
    Returns:
        tuple: (success: bool, file_path: str or None, error_message: str or None)
    """
    try:
        # Get SAP session
        success, session, error_msg = get_sap_session()
        if not success:
            return False, None, error_msg
        
        print("Connected to SAP. Navigating to transaction...")
        
        # Navigate to the transaction
        success, error_msg = navigate_to_transaction(session, transaction_code)
        if not success:
            return False, None, error_msg
        
        print(f"✓ Navigated to transaction {transaction_code}")
        time.sleep(1)  # Wait for screen to fully load
        
        print("Executing export workflow...")
        
        # Step 1: Resize working pane (from VBA line 16)
        try:
            session.FindById("wnd[0]").ResizeWorkingPane(143, 23, False)
        except Exception as e:
            print(f"Warning: Could not resize working pane: {e}")
        
        # Step 2: Set Order number (from VBA line 17)
        try:
            session.FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = order_num
        except Exception as e:
            return False, None, f"Could not set Order number field: {e}. Please ensure you are on the correct SAP transaction."
        
        # Step 3: Set Controlling Area (from VBA line 18)
        try:
            session.FindById("wnd[0]/usr/ctxtKOAGR").Text = controlling_area
        except Exception as e:
            return False, None, f"Could not set Controlling Area field: {e}"
        
        # Step 4: Set Date From field (simplified - skip calendar picker that was a mistake in VBA)
        try:
            date_low_field = session.FindById("wnd[0]/usr/ctxtR_BUDAT-LOW")
            date_low_field.Text = date_from
        except Exception as e:
            return False, None, f"Could not set Date From field: {e}"
        
        # Step 5: Set Date To field (simplified - skip focus/caret position that was a mistake in VBA)
        try:
            date_high_field = session.FindById("wnd[0]/usr/ctxtR_BUDAT-HIGH")
            date_high_field.Text = date_to
        except Exception as e:
            return False, None, f"Could not set Date To field: {e}"
        
        # Step 6: Execute report (button 8) (from VBA line 27)
        print("Executing report...")
        try:
            session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
        except Exception as e:
            return False, None, f"Could not execute report: {e}. Please check that all fields are filled correctly."
        
        # Wait for report to load - poll for grid to be available
        print("Waiting for report to load...")
        grid = None
        max_wait_time = 30  # Maximum wait time in seconds
        check_interval = 0.5  # Check every 0.5 seconds
        elapsed_time = 0
        
        while elapsed_time < max_wait_time:
            try:
                grid = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
                # If we can access the grid, it's loaded
                break
            except:
                time.sleep(check_interval)
                elapsed_time += check_interval
                if int(elapsed_time) % 2 == 0:  # Print every 2 seconds
                    print(f"  Still waiting... ({int(elapsed_time)}s)")
        
        if grid is None:
            return False, None, f"Report did not load within {max_wait_time} seconds. Please check if the report executed successfully or if there's an error message in SAP."
        
        print("✓ Report loaded successfully")
        
        # Step 7: Prepare grid for export (from VBA lines 28-30)
        # These operations are optional - try each one individually, but don't fail if they don't work
        print("Preparing grid for export...")
        try:
            grid.SetCurrentCell(4, "WRBTR")
        except Exception as e:
            print(f"  Warning: Could not set current cell: {e}")
        
        try:
            grid.FirstVisibleColumn = "UOB_TXT"
        except Exception as e:
            print(f"  Warning: Could not set first visible column: {e}")
        
        try:
            grid.SelectedRows = "4"
        except Exception as e:
            print(f"  Warning: Could not set selected rows: {e}")
        
        print("✓ Grid prepared (some operations may have been skipped)")
        
        # Step 8: Export to Excel - First export (from VBA lines 31-35)
        print("Exporting to Excel (first export)...")
        try:
            grid.ContextMenu()
            grid.SelectContextMenuItem("&XXL")
            
            # Handle Excel save dialog
            time.sleep(1)
            try:
                # Button 0 is typically "OK" or "Save"
                session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
                # Button 12 might be "Save" or similar
                session.FindById("wnd[1]/tbar[0]/btn[12]").Press()
            except Exception as e:
                print(f"Warning: Could not handle first export dialog automatically: {e}")
                print("You may need to manually handle the save dialog.")
            
            # Close export dialog
            try:
                time.sleep(0.5)
                session.FindById("wnd[1]").Close()
            except:
                pass
        except Exception as e:
            return False, None, f"Error during first Excel export: {e}"
        
        time.sleep(1)
        
        # Step 9: Export to Excel - Second export (from VBA lines 36-39)
        # The VBA script does this twice, so we'll do the same
        print("Exporting to Excel (second export)...")
        try:
            grid.ContextMenu()
            grid.SelectContextMenuItem("&XXL")
            
            # Handle Excel save dialog - second time
            time.sleep(1)
            try:
                session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
                session.FindById("wnd[1]/tbar[0]/btn[0]").Press()  # Press twice as in VBA
            except Exception as e:
                print(f"Warning: Could not handle second export dialog automatically: {e}")
                print("You may need to manually handle the save dialog.")
        except Exception as e:
            return False, None, f"Error during second Excel export: {e}"
        
        # Determine output file path
        if output_path is None:
            # Use temp directory with order number
            temp_dir = os.path.join(os.environ.get('TEMP', os.getcwd()), 'sap_exports')
            os.makedirs(temp_dir, exist_ok=True)
            output_path = os.path.join(temp_dir, f"sap_export_{order_num}.xlsx")
        
        # Note: The actual file path will depend on where SAP saves it
        # The user may need to specify the save location in the Excel dialog
        # For now, we'll return a suggested path
        print(f"\n✓ Export workflow completed!")
        print(f"Note: If a save dialog appeared, please note where you saved the file.")
        print(f"Suggested output path: {output_path}")
        
        return True, output_path, None
        
    except Exception as e:
        error_msg = f"Error during SAP export: {str(e)}"
        import traceback
        print(f"Detailed error: {traceback.format_exc()}")
        return False, None, error_msg


def handle_sap_export_workflow():
    """
    Complete workflow for SAP export with user interaction.
    Handles connection checking, login prompting, parameter collection, and export.
    
    Returns:
        tuple: (success: bool, file_path: str or None, error_message: str or None)
    """
    # Check SAP connection
    print("\nChecking SAP connection...")
    is_connected, session, error_msg = check_sap_connection()
    
    if not is_connected:
        print(f"SAP is not connected: {error_msg}")
        response = input("Would you like to launch SAP Logon? (y/n): ").strip().lower()
        
        if response == 'y':
            print("Launching SAP Logon...")
            if launch_sap_gui():
                print("SAP Logon launched. Please log in when prompted.")
                success, session, error_msg = wait_for_sap_login()
                if not success:
                    return False, None, error_msg
            else:
                return False, None, "Could not launch SAP Logon. Please launch it manually and try again."
        else:
            return False, None, "SAP connection required for automated export."
    
    # Get parameters
    params = prompt_sap_parameters()
    
    # Execute export (KOB1 is the transaction code)
    success, file_path, error_msg = execute_sap_export(
        params['order_num'],
        params['controlling_area'],
        params['date_from'],
        params['date_to'],
        transaction_code="KOB1"
    )
    
    return success, file_path, error_msg

