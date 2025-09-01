"""
Verification script to confirm the date mapping fix is working correctly
in both transformer.py and excel_processor.py
"""

import pandas as pd
from datetime import date, datetime
from pathlib import Path


def test_transformer_date_mapping():
    """Test the transformer.py date mapping."""
    print("=== Testing transformer.py Date Mapping ===\n")
    
    from transformer import transform_to_master_format
    
    # Create sample input data
    sample_data = {
        'MatrixDescription': ['HOUSTON High Load Factor', 'NORTH Low Load Factor'],
        'Price': [75.50, 80.25],
        'TermCode': [12, 24],
        'StartDate': ['2024-03-01', '2024-06-01'],
        'Product': ['Fixed Price', 'Fixed Price']
    }
    
    input_df = pd.DataFrame(sample_data)
    print("Input data:")
    print(input_df)
    print()
    
    # Transform
    result_df = transform_to_master_format(input_df)
    
    if result_df.empty:
        print("❌ Transformation failed")
        return False
    
    print("Transformed result:")
    print(result_df[['Price_Date', 'Date', 'Zone', 'Load', 'Term']].head())
    print()
    
    # Verify dates
    today = date.today()
    price_dates = result_df['Price_Date'].unique()
    dates = result_df['Date'].tolist()
    
    print("Date verification:")
    print(f"  Column B (Price_Date): {price_dates}")
    print(f"  Column C (Date): {dates}")
    
    # Check Column B (Price_Date) - should be today
    if len(price_dates) == 1 and price_dates[0] == today:
        print("  ✅ Column B correctly set to today's date")
    else:
        print(f"  ❌ Column B should be {today}")
        return False
    
    # Check Column C (Date) - should be start dates from input
    expected_dates = [date(2024, 3, 1), date(2024, 6, 1)]
    if dates == expected_dates:
        print("  ✅ Column C correctly set to start dates from input")
    else:
        print(f"  ❌ Column C should be {expected_dates}")
        return False
    
    return True


def test_excel_processor_date_mapping():
    """Test the excel_processor.py date mapping."""
    print("\n=== Testing excel_processor.py Date Mapping ===\n")
    
    try:
        from excel_processor import hda_matrix_to_master_cols
    except ImportError as e:
        print(f"❌ Could not import from excel_processor: {e}")
        return False
    
    # Create sample input data
    sample_data = {
        'MatrixDescription': ['HOUSTON High Load Factor', 'SOUTH Medium Load Factor'],
        'Price': [72.30, 78.90],
        'TermCode': [36, 48],
        'StartDate': ['2024-09-01', '2024-12-01'],
        'Product': ['Fixed Price', 'Fixed Price']
    }
    
    input_df = pd.DataFrame(sample_data)
    print("Input data:")
    print(input_df)
    print()
    
    # Transform
    result_df = hda_matrix_to_master_cols(input_df)
    
    if result_df.empty:
        print("❌ Transformation failed")
        return False
    
    print("Transformed result:")
    print(result_df[['Price_Date', 'Date', 'Zone', 'Load', 'Term']].head())
    print()
    
    # Verify dates
    today = date.today()
    price_dates = result_df['Price_Date'].unique()
    dates = result_df['Date'].tolist()
    
    print("Date verification:")
    print(f"  Column B (Price_Date): {price_dates}")
    print(f"  Column C (Date): {dates}")
    
    # Check Column B (Price_Date) - should be today
    if len(price_dates) == 1 and price_dates[0] == today:
        print("  ✅ Column B correctly set to today's date")
    else:
        print(f"  ❌ Column B should be {today}")
        return False
    
    # Check Column C (Date) - should be start dates from input
    expected_dates = [date(2024, 9, 1), date(2024, 12, 1)]
    if dates == expected_dates:
        print("  ✅ Column C correctly set to start dates from input")
    else:
        print(f"  ❌ Column C should be {expected_dates}")
        return False
    
    return True


def test_column_mapping_consistency():
    """Test that both modules produce consistent column mapping."""
    print("\n=== Testing Column Mapping Consistency ===\n")
    
    # Same input data for both
    sample_data = {
        'MatrixDescription': ['HOUSTON High Load Factor'],
        'Price': [75.50],
        'TermCode': [12],
        'StartDate': ['2024-03-01'],
        'Product': ['Fixed Price']
    }
    
    input_df = pd.DataFrame(sample_data)
    
    # Test transformer.py
    try:
        from transformer import transform_to_master_format
        transformer_result = transform_to_master_format(input_df)
        transformer_cols = list(transformer_result.columns)
    except Exception as e:
        print(f"❌ Transformer test failed: {e}")
        return False
    
    # Test excel_processor.py
    try:
        from excel_processor import hda_matrix_to_master_cols
        processor_result = hda_matrix_to_master_cols(input_df)
        processor_cols = list(processor_result.columns)
    except Exception as e:
        print(f"❌ Excel processor test failed: {e}")
        return False
    
    print("Column order comparison:")
    print(f"  transformer.py:     {transformer_cols}")
    print(f"  excel_processor.py: {processor_cols}")
    
    if transformer_cols == processor_cols:
        print("  ✅ Column orders match between modules")
        
        # Check specific date columns
        if transformer_cols[1] == 'Price_Date' and transformer_cols[2] == 'Date':
            print("  ✅ Date columns in correct positions (B=Price_Date, C=Date)")
            return True
        else:
            print("  ❌ Date columns not in expected positions")
            return False
    else:
        print("  ❌ Column orders differ between modules")
        return False


def show_expected_output():
    """Show what the expected output should look like in Excel."""
    print("\n=== Expected Excel Output Format ===\n")
    
    today = date.today()
    
    print("When data is appended to Master-Table.xlsx, it should look like:")
    print()
    print("| A  | B          | C          | D       | E    | F      | G    | ... |")
    print("|----|------------|------------|---------|------|--------|------|-----|")
    print("| ID | Price_Date | Date       | Zone    | Load | REP1   | Term | ... |")
    print(f"| 1  | {today}  | 2024-03-01 | HOUSTON | HIGH | HUDSON | 12   | ... |")
    print(f"| 2  | {today}  | 2024-06-01 | NORTH   | LOW  | HUDSON | 24   | ... |")
    print()
    print("Key points:")
    print(f"  • Column B (Price_Date): Always today's date ({today})")
    print("  • Column C (Date): Start date from input file (varies per row)")
    print("  • This ensures proper date tracking for pricing data")


if __name__ == "__main__":
    print("Date Mapping Fix Verification")
    print("=" * 50)
    
    # Run all tests
    test1_passed = test_transformer_date_mapping()
    test2_passed = test_excel_processor_date_mapping()
    test3_passed = test_column_mapping_consistency()
    
    all_passed = test1_passed and test2_passed and test3_passed
    
    if all_passed:
        print("\n🎉 All tests passed! Date mapping fix is working correctly.")
        show_expected_output()
    else:
        print("\n❌ Some tests failed. Please check the issues above.")
    
    print(f"\nSummary:")
    print(f"  transformer.py:     {'✅' if test1_passed else '❌'}")
    print(f"  excel_processor.py: {'✅' if test2_passed else '❌'}")
    print(f"  Consistency:        {'✅' if test3_passed else '❌'}")
