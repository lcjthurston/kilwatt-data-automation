"""
Demonstration script showing how to use the backup functionality
before modifying the Master-Table.xlsx file.
"""

from pathlib import Path
from transformer import create_master_table_backup
import pandas as pd


def demonstrate_backup_workflow():
    """
    Demonstrate the complete backup workflow before appending data.
    """
    print("=== Master Table Backup Demonstration ===\n")
    
    # Define the master table path
    master_table_path = Path("2-copy-reformat/Master-Table.xlsx")
    
    print(f"1. Checking if master table exists: {master_table_path}")
    if not master_table_path.exists():
        print(f"   ❌ Master table not found at {master_table_path}")
        return False
    else:
        print(f"   ✅ Master table found")
    
    print(f"\n2. Creating backup before any modifications...")
    backup_path = create_master_table_backup(master_table_path)
    
    if backup_path is None:
        print("   ❌ Backup creation failed!")
        return False
    else:
        print(f"   ✅ Backup created successfully")
        print(f"   📁 Backup location: {backup_path}")
    
    print(f"\n3. Verifying backup file exists...")
    if backup_path.exists():
        print(f"   ✅ Backup file verified")
        
        # Get file sizes for comparison
        original_size = master_table_path.stat().st_size
        backup_size = backup_path.stat().st_size
        
        print(f"   📊 Original file size: {original_size:,} bytes")
        print(f"   📊 Backup file size: {backup_size:,} bytes")
        
        if original_size == backup_size:
            print(f"   ✅ File sizes match - backup is complete")
        else:
            print(f"   ⚠️  File sizes differ - backup may be incomplete")
    else:
        print(f"   ❌ Backup file not found!")
        return False
    
    print(f"\n4. Listing all backups in backup directory...")
    backup_dir = master_table_path.parent / "backups"
    if backup_dir.exists():
        backup_files = list(backup_dir.glob("Master-Table_backup_*.xlsx"))
        print(f"   📁 Found {len(backup_files)} backup files:")
        for i, backup_file in enumerate(sorted(backup_files), 1):
            file_size = backup_file.stat().st_size
            print(f"      {i}. {backup_file.name} ({file_size:,} bytes)")
    else:
        print(f"   📁 No backup directory found")
    
    print(f"\n✅ Backup workflow completed successfully!")
    print(f"💡 You can now safely modify the master table knowing you have a backup.")
    
    return True


def demonstrate_safe_append_pattern():
    """
    Demonstrate the recommended pattern for safely appending data.
    """
    print("\n=== Safe Append Pattern Demonstration ===\n")
    
    master_table_path = Path("2-copy-reformat/Master-Table.xlsx")
    
    # Step 1: Create backup
    print("1. Creating backup before modifications...")
    backup_path = create_master_table_backup(master_table_path)
    
    if backup_path is None:
        print("   ❌ Cannot proceed without backup!")
        return False
    
    print(f"   ✅ Backup created: {backup_path.name}")
    
    # Step 2: Prepare sample data (this would be your actual data)
    print("\n2. Preparing sample data to append...")
    sample_data = {
        'Product': ['Fixed Price', 'Fixed Price'],
        'Term': [12, 24],
        'Start Month': ['2024-01-01', '2024-02-01'],
        'Zone': ['HOUSTON', 'NORTH'],
        'Load Factor': ['HIGH', 'MED'],
        'Price': [75.5, 80.2]
    }
    df = pd.DataFrame(sample_data)
    print(f"   📊 Sample data prepared: {len(df)} rows")
    
    # Step 3: Transform data (using transformer functions)
    print("\n3. Transforming data to master format...")
    from transformer import transform_to_master_format
    
    try:
        transformed_df = transform_to_master_format(df)
        print(f"   ✅ Data transformed: {len(transformed_df)} rows ready for append")
        
        if not transformed_df.empty:
            print(f"   📋 Transformed columns: {list(transformed_df.columns)}")
        else:
            print(f"   ⚠️  No data after transformation (may be filtered out)")
    
    except Exception as e:
        print(f"   ❌ Transformation failed: {e}")
        return False
    
    # Step 4: Simulate append (don't actually modify the file in demo)
    print("\n4. Ready to append data to master table...")
    print(f"   💾 Would append {len(transformed_df)} rows to {master_table_path}")
    print(f"   🔒 Master table is protected by backup: {backup_path.name}")
    print(f"   ✅ Safe to proceed with actual append operation")
    
    return True


def show_backup_best_practices():
    """
    Display best practices for using the backup functionality.
    """
    print("\n=== Backup Best Practices ===\n")
    
    practices = [
        "🔄 Always create a backup before modifying the master table",
        "📅 Backups are timestamped for easy identification",
        "📁 Backups are stored in a separate 'backups' subdirectory",
        "🔍 Verify backup creation was successful before proceeding",
        "🧹 Periodically clean up old backups to save disk space",
        "💾 Keep at least the last few backups for recovery options",
        "⚡ Backup creation is fast - don't skip it to save time",
        "🛡️ Backups protect against data corruption and user errors"
    ]
    
    for practice in practices:
        print(f"   {practice}")
    
    print(f"\n💡 The backup function is integrated into all append operations")
    print(f"   in both excel_processor.py and transformer.py")


if __name__ == "__main__":
    print("Master Table Backup Functionality Demo")
    print("=" * 50)
    
    # Run demonstrations
    success = demonstrate_backup_workflow()
    
    if success:
        demonstrate_safe_append_pattern()
        show_backup_best_practices()
    
    print(f"\n🎉 Demo completed!")
    print(f"📖 Check the 'backups' directory to see your backup files.")
