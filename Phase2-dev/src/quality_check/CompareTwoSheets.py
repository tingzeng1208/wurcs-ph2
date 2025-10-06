import openpyxl
from openpyxl import load_workbook
from typing import Optional, Tuple, List, Dict, Any
import os
from datetime import datetime




def compare_worksheets(workbook1_path: str, worksheet1_name: str, 
                      workbook2_path: str, worksheet2_name: str,
                      max_row: Optional[int] = None, max_col: Optional[int] = None,
                      case_sensitive: bool = False) -> Dict[str, Any]:
    """
    Compare two worksheets cell by cell and return differences.
    
    Args:
        workbook1_path (str): Path to the first Excel workbook
        worksheet1_name (str): Name of worksheet in first workbook
        workbook2_path (str): Path to the second Excel workbook  
        worksheet2_name (str): Name of worksheet in second workbook
        max_row (int, optional): Maximum row to compare (if None, compares all rows)
        max_col (int, optional): Maximum column to compare (if None, compares all columns)
        case_sensitive (bool): Whether comparison should be case sensitive (default: False)
        
    Returns:
        Dict containing comparison results and statistics
    """
    
    # Load workbooks
    try:
        wb1 = load_workbook(workbook1_path, data_only=True)
        wb2 = load_workbook(workbook2_path, data_only=True)
    except Exception as e:
        print(f"Error loading workbooks: {e}")
        return {"error": str(e)}
    
    # Get worksheets
    try:
        ws1 = wb1[worksheet1_name]
        ws2 = wb2[worksheet2_name]
    except KeyError as e:
        print(f"Error accessing worksheet: {e}")
        return {"error": f"Worksheet not found: {e}"}
    
    # Determine comparison range
    if max_row is None:
        max_row = max(ws1.max_row, ws2.max_row)
    if max_col is None:
        max_col = max(ws1.max_column, ws2.max_column)
    
    differences = []
    total_cells_compared = 0
    identical_cells = 0
    different_cells = 0
    
    print(f"Comparing worksheets:")
    print(f"  Workbook 1: {os.path.basename(workbook1_path)} -> {worksheet1_name}")
    print(f"  Workbook 2: {os.path.basename(workbook2_path)} -> {worksheet2_name}")
    print(f"  Range: A1 to {openpyxl.utils.get_column_letter(max_col)}{max_row}")
    print(f"  Case sensitive: {case_sensitive}")
    print("-" * 80)
    
    # Compare cell by cell
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            total_cells_compared += 1
            
            # Get cell values
            cell1 = ws1.cell(row=row, column=col)
            cell2 = ws2.cell(row=row, column=col)
            
            value1 = cell1.value
            value2 = cell2.value
            
            # Convert to string for comparison (handle None values)
            str1 = str(value1) if value1 is not None else ""
            str2 = str(value2) if value2 is not None else ""

            # Apply case sensitivity setting
            if not case_sensitive:
                str1 = str1.lower()
                str2 = str2.lower()

            # Special rule for cell (1,1)
            if row == 1 and col == 1:
                import re
                def extract_special_parts(s):
                    # Match: first two letters, year (4 digits), and "Run"
                    m = re.match(r"([a-zA-Z]{2}).*?(\d{4}).*?(run)", s)
                    if m:
                        return m.group(1), m.group(2), m.group(3)
                    return None
                parts1 = extract_special_parts(str1)
                parts2 = extract_special_parts(str2)
                if parts1 and parts2 and parts1 == parts2:
                    identical_cells += 1
                    continue  # Considered identical, skip to next cell

            # Treat as identical if only difference is a single leading apostrophe
            def strip_leading_apostrophe(s):
                return s[1:] if s.startswith("'") else s

            if str1 != str2:
                # Check for leading apostrophe difference
                if strip_leading_apostrophe(str1) == str2 or str1 == strip_leading_apostrophe(str2):
                    identical_cells += 1
                    continue  # Considered identical, skip to next cell

                different_cells += 1
                cell_address = f"{openpyxl.utils.get_column_letter(col)}{row}"

                difference = {
                    "cell": cell_address,
                    "workbook1_value": value1,
                    "workbook2_value": value2,
                    "workbook1_display": str(value1) if value1 is not None else "None",
                    "workbook2_display": str(value2) if value2 is not None else "None"
                }
                differences.append(difference)
                
                # Print difference immediately
                print(f"Cell {cell_address:6}: '{difference['workbook1_display']:20}' != '{difference['workbook2_display']:20}'")
            else:
                identical_cells += 1
    
    # Print summary
    print("-" * 80)
    print(f"Comparison Summary:")
    print(f"  Total cells compared: {total_cells_compared:,}")
    print(f"  Identical cells: {identical_cells:,}")
    print(f"  Different cells: {different_cells:,}")
    print(f"  Match percentage: {(identical_cells/total_cells_compared)*100:.2f}%")
    
    # Write results to ComparisonResult.txt in the same folder as workbook1
    try:
        workbook1_dir = os.path.dirname(os.path.abspath(workbook1_path))
        output_file = os.path.join(workbook1_dir, "ComparisonResult.txt")
        
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(f"Excel Worksheet Comparison Report\n")
            f.write(f"Generated: {datetime.now()}\n")
            f.write(f"="*80 + "\n\n")
            
            f.write(f"Source Files:\n")
            f.write(f"  Workbook 1: {workbook1_path} -> {worksheet1_name}\n")
            f.write(f"  Workbook 2: {workbook2_path} -> {worksheet2_name}\n\n")
            
            f.write(f"Comparison Settings:\n")
            f.write(f"  Max Row: {max_row}\n")
            f.write(f"  Max Col: {max_col}\n")
            f.write(f"  Case Sensitive: {case_sensitive}\n\n")
            
            f.write(f"Summary:\n")
            f.write(f"  Total cells compared: {total_cells_compared:,}\n")
            f.write(f"  Identical cells: {identical_cells:,}\n")
            f.write(f"  Different cells: {different_cells:,}\n")
            f.write(f"  Match percentage: {(identical_cells/total_cells_compared)*100:.2f}%\n\n")
            
            if differences:
                f.write(f"Differences Found ({len(differences)}):\n")
                f.write(f"{'Cell':8} {'Workbook 1':30} {'Workbook 2':30}\n")
                f.write(f"{'-'*8} {'-'*30} {'-'*30}\n")
                
                for diff in differences:
                    wb1_val = str(diff['workbook1_display'])[:29]
                    wb2_val = str(diff['workbook2_display'])[:29]
                    f.write(f"{diff['cell']:8} {wb1_val:30} {wb2_val:30}\n")
            else:
                f.write("No differences found - worksheets are identical!\n")
        
        print(f"Comparison results saved to: {output_file}")
        
    except Exception as e:
        print(f"Error saving comparison results: {e}")
    
    # Return results
    return {
        "success": True,
        "workbook1_path": workbook1_path,
        "worksheet1_name": worksheet1_name,
        "workbook2_path": workbook2_path,
        "worksheet2_name": worksheet2_name,
        "total_cells_compared": total_cells_compared,
        "identical_cells": identical_cells,
        "different_cells": different_cells,
        "match_percentage": (identical_cells/total_cells_compared)*100,
        "differences": differences,
        "case_sensitive": case_sensitive,
        "max_row": max_row,
        "max_col": max_col,
        "output_file": output_file if 'output_file' in locals() else None
    }


def compare_worksheets_detailed(workbook1_path: str, worksheet1_name: str,
                               workbook2_path: str, worksheet2_name: str,
                               output_file: Optional[str] = None,
                               **kwargs) -> Dict[str, Any]:
    """
    Compare worksheets and optionally save detailed results to a file.
    
    Args:
        workbook1_path (str): Path to the first Excel workbook
        worksheet1_name (str): Name of worksheet in first workbook
        workbook2_path (str): Path to the second Excel workbook
        worksheet2_name (str): Name of worksheet in second workbook
        output_file (str, optional): Path to save detailed comparison results
        **kwargs: Additional arguments passed to compare_worksheets()
        
    Returns:
        Dict containing comparison results
    """
    
    result = compare_worksheets(workbook1_path, worksheet1_name, 
                               workbook2_path, worksheet2_name, **kwargs)
    
    if result.get("success") and output_file:
        try:
            with open(output_file, 'a', encoding='utf-8') as f:
                f.write(f"Excel Worksheet Comparison Report\n")
                f.write(f"Generated: {datetime.now()}\n")
                f.write(f"="*80 + "\n\n")
                
                f.write(f"Source Files:\n")
                f.write(f"  Workbook 1: {result['workbook1_path']} -> {result['worksheet1_name']}\n")
                f.write(f"  Workbook 2: {result['workbook2_path']} -> {result['worksheet2_name']}\n\n")
                
                f.write(f"Comparison Settings:\n")
                f.write(f"  Max Row: {result['max_row']}\n")
                f.write(f"  Max Col: {result['max_col']}\n")
                f.write(f"  Case Sensitive: {result['case_sensitive']}\n\n")
                
                f.write(f"Summary:\n")
                f.write(f"  Total cells compared: {result['total_cells_compared']:,}\n")
                f.write(f"  Identical cells: {result['identical_cells']:,}\n")
                f.write(f"  Different cells: {result['different_cells']:,}\n")
                f.write(f"  Match percentage: {result['match_percentage']:.2f}%\n\n")
                
                if result['differences']:
                    f.write(f"Differences Found ({len(result['differences'])}):\n")
                    f.write(f"{'Cell':8} {'Workbook 1':25} {'Workbook 2':25}\n")
                    f.write(f"{'-'*8} {'-'*25} {'-'*25}\n")
                    
                    for diff in result['differences']:
                        f.write(f"{diff['cell']:8} {str(diff['workbook1_display'])[:24]:25} {str(diff['workbook2_display'])[:24]:25}\n")
                else:
                    f.write("No differences found - worksheets are identical!\n")
                    
            print(f"\nDetailed comparison report saved to: {output_file}")
            
        except Exception as e:
            print(f"Error saving detailed report: {e}")
    
    return result




# Example usage and test function
if __name__ == "__main__":
    # Example usage
    print("Excel Worksheet Comparison Tool")
    print("=" * 50)

    reports_dir = os.path.join(os.path.dirname(__file__), "../..", "reports")
    source_wb1 = os.path.abspath(os.path.join(reports_dir, "CN2023.xlsx"))
    target_wb2 = os.path.abspath(os.path.join(reports_dir, "CN-2023_report.xlsx"))

    # Remove ComparisonResult.txt if it exists
    comparison_result_path = os.path.join(reports_dir, "ComparisonResult.txt")
    if os.path.exists(comparison_result_path):
        os.remove(comparison_result_path)

    sheetnames = ["INDEX", "A1P1", "A1P2A", "A1P2B"]

    for sheetname in sheetnames:
        print(f"\nComparing worksheet: {sheetname}")

        # Check if files exist
        if os.path.exists(source_wb1) and os.path.exists(target_wb2):
            # Compare the actual worksheets
            result = compare_worksheets(
                workbook1_path=source_wb1,
                worksheet1_name=sheetname, 
                workbook2_path=target_wb2,
                worksheet2_name=sheetname,
                case_sensitive=False
            )
            print(f"\nComparison completed. Success: {result.get('success', False)}")
        else:
            print("One or both Excel files not found:")
            print(f"  File 1 exists: {os.path.exists(source_wb1)} - {source_wb1}")
            print(f"  File 2 exists: {os.path.exists(target_wb2)} - {target_wb2}")

            # Example with placeholder paths
            print("\nExample usage:")
            result = compare_worksheets(
                workbook1_path="path/to/workbook1.xlsx",
                worksheet1_name="Sheet1", 
                workbook2_path="path/to/workbook2.xlsx",
                worksheet2_name="Sheet1",
                case_sensitive=False
            )
    # Example 2: Compare with detailed output
    # result = compare_worksheets_detailed(
    #     workbook1_path="path/to/workbook1.xlsx",
    #     worksheet1_name="INDEX",
    #     workbook2_path="path/to/workbook2.xlsx", 
    #     worksheet2_name="INDEX",
    #     output_file="comparison_report.txt",
    #     max_row=100,
    #     max_col=10,
    #     case_sensitive=False
    # )
    
    # Example 3: Compare multiple worksheet pairs
    # comparisons = [
    #     ("wb1.xlsx", "INDEX", "wb2.xlsx", "INDEX"),
    #     ("wb1.xlsx", "A1P1", "wb2.xlsx", "A1P1"),
    # ]
    # results = compare_multiple_worksheets(comparisons, case_sensitive=False)
    
    print("Functions defined. Use compare_worksheets(), compare_worksheets_detailed(), or compare_multiple_worksheets()")
