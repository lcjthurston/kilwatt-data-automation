"""
Test script to verify the date mapping logic is working correctly.
Column B should be today's date, Column C should be the start date from input.
"""

import pandas as pd
from datetime import date, datetime
from transformer import transform_to_master_format


def test_date_mapping():
    """Test that date mapping works correctly."""
    print("=== Testing Date Mapping Logic ===\n")
    
    # Create sample input data that mimics HDA matrix table format
    sample_data = {
        'MatrixDescription': ['HOUSTON High Load Factor', 'NORTH Low Load Factor'],
        'Price': [75.50, 80.25],
        'TermCode': [12, 24],
        'StartDate': ['2024-03-01', '2024-06-01'],  # These should go to Column C
        'Product': ['Fixed Price', 'Fixed Price']
    }
    
    input_df = pd.DataFrame(sample_data)
    print("Input data:")
    print(input_df)
    print()
    
    # Transform to master format
    transformed_df = transform_to_master_format(input_df)
    
    if transformed_df.empty:
        print("❌ Transformation failed - no data returned")
        return False
    
    print("Transformed data:")
    print(transformed_df[['Price_Date', 'Date', 'Zone', 'Load', 'Term']].head())
    print()
    
    # Verify date mapping
    today = date.today()
    
    print("=== Date Mapping Verification ===")
    
    # Check Price_Date (Column B)
    price_dates = transformed_df['Price_Date'].unique()
    print(f"Column B (Price_Date) values: {price_dates}")
    
    if len(price_dates) == 1 and price_dates[0] == today:
        print("✅ Column B (Price_Date) correctly set to today's date")
    else:
        print(f"❌ Column B (Price_Date) should be {today}, but got {price_dates}")
        return False
    
    # Check Date (Column C)
    dates = transformed_df['Date'].tolist()
    expected_dates = [date(2024, 3, 1), date(2024, 6, 1)]
    print(f"Column C (Date) values: {dates}")
    print(f"Expected dates: {expected_dates}")
    
    if dates == expected_dates:
        print("✅ Column C (Date) correctly set to start dates from input")
    else:
        print(f"❌ Column C (Date) mapping incorrect")
        return False
    
    print("\n=== Summary ===")
    print("✅ Date mapping logic is working correctly!")
    print(f"   • Column B (Price_Date): Today's date ({today})")
    print(f"   • Column C (Date): Start dates from input file")
    
    return True


def test_column_positions():
    """Test that the columns are in the correct positions for the master table."""
    print("\n=== Testing Column Positions ===\n")
    
    # Create sample data
    sample_data = {
        'MatrixDescription': ['HOUSTON High Load Factor'],
        'Price': [75.50],
        'TermCode': [12],
        'StartDate': ['2024-03-01'],
        'Product': ['Fixed Price']
    }
    
    input_df = pd.DataFrame(sample_data)
    transformed_df = transform_to_master_format(input_df)
    
    if transformed_df.empty:
        print("❌ No data to test")
        return False
    
    # Expected master table column order
    expected_columns = [
        'ID', 'Price_Date', 'Date', 'Zone', 'Load', 'REP1', 'Term', 'Min_MWh', 'Max_MWh',
        'Daily_No_Ruc', 'RUC_Nodal', 'Daily', 'Com_Disc', 'HOA_Disc', 'Broker_Fee', 'Meter_Fee', 'Max_Meters'
    ]
    
    actual_columns = list(transformed_df.columns)
    
    print("Expected column order:")
    for i, col in enumerate(expected_columns):
        excel_col = chr(65 + i)  # A, B, C, etc.
        print(f"   {excel_col}: {col}")
    
    print("\nActual column order:")
    for i, col in enumerate(actual_columns):
        excel_col = chr(65 + i)  # A, B, C, etc.
        print(f"   {excel_col}: {col}")
    
    if actual_columns == expected_columns:
        print("\n✅ Column positions are correct!")
        print("   • Column A: ID")
        print("   • Column B: Price_Date (today's date)")
        print("   • Column C: Date (start date from input)")
        return True
    else:
        print("\n❌ Column positions are incorrect!")
        return False


def show_sample_output():
    """Show what the actual output should look like."""
    print("\n=== Sample Output Preview ===\n")
    
    # Create sample data
    sample_data = {
        'MatrixDescription': ['HOUSTON High Load Factor', 'NORTH Low Load Factor'],
        'Price': [75.50, 80.25],
        'TermCode': [12, 24],
        'StartDate': ['2024-03-01', '2024-06-01'],
        'Product': ['Fixed Price', 'Fixed Price']
    }
    
    input_df = pd.DataFrame(sample_data)
    transformed_df = transform_to_master_format(input_df)
    
    if not transformed_df.empty:
        print("Sample transformed output (first few columns):")
        preview_cols = ['ID', 'Price_Date', 'Date', 'Zone', 'Load', 'REP1', 'Term']
        print(transformed_df[preview_cols])
        
        print(f"\nThis should result in:")
        print(f"   • Column B: {date.today()} (today's date)")
        print(f"   • Column C: 2024-03-01, 2024-06-01 (start dates from input)")


if __name__ == "__main__":
    print("Date Mapping Test for Master Table Transformation")
    print("=" * 60)
    
    success = test_date_mapping()
    
    if success:
        test_column_positions()
        show_sample_output()
    
    print(f"\n{'✅ All tests passed!' if success else '❌ Tests failed!'}")
