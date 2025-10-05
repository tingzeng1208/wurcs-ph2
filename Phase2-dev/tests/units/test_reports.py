import os
import pytest
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Add project root to Python path
import sys
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from core.Create_reports_new import CreateReports
from quality_check.CompareTwoSheets import compare_worksheets

@pytest.fixture(scope="module")
def report_generator():
    """Fixture to initialize CreateReports and prepare data once per module."""
    print("Setting up report generator...")
    # Initialize with a specific year
    generator = CreateReports(year=2023)
    
    # Initialize data needed for all reports
    generator.initialize_report_format()
    
    # Get the first railroad for testing
    df_rr = generator.o_db.get_class1_rail_list()
    if df_rr.empty:
        pytest.fail("No railroad data found for testing.")
    
    test_railroad = df_rr.iloc[0]
    generator.current_df = test_railroad
    
    # Initialize data specific to this railroad
    generator.initialize_report_for_a_railroad(test_railroad.rr_id)
    
    yield generator
    print("\nReport generator teardown.")

def test_A1P2A_worksheet(report_generator):
    """Test the generation of the A1P2A worksheet."""
    short_name = report_generator.current_df.short_name
    year = report_generator.current_year
    
    # Generate the report
    report_generator.create_a_report(short_name)
    
    # Define paths for comparison
    reports_dir = os.path.abspath(os.path.join(parent_dir, "..", "reports"))
    generated_file = os.path.join(reports_dir, f"{short_name}-{year}_report.xlsx")
    reference_file = os.path.join(reports_dir, f"{short_name}{year}.xlsx") # e.g., CN2023.xlsx

    assert os.path.exists(generated_file), f"Generated file not found: {generated_file}"
    assert os.path.exists(reference_file), f"Reference file not found: {reference_file}"

    # Compare the specific worksheet
    comparison_result = compare_worksheets(reference_file, "A1P2A", generated_file, "A1P2A")
    
    assert comparison_result.get("different_cells", -1) == 0, f"Worksheet A1P2A has differences. Check ComparisonResult.txt in {reports_dir}."

