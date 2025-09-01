# Master Table Backup Functionality

## Overview

Before running the transformer.py file or any operations that modify the Master-Table.xlsx file, a timestamped backup copy is automatically created to protect against data loss.

## Features

### Automatic Backup Creation
- **Timestamped backups**: Each backup includes a timestamp in the format `YYYYMMDD_HHMMSS`
- **Organized storage**: Backups are stored in a dedicated `backups` subdirectory
- **File integrity**: Backup files are exact copies using `shutil.copy2()` which preserves metadata
- **Error handling**: Graceful handling of backup failures with appropriate warnings

### Backup Location
```
2-copy-reformat/
├── Master-Table.xlsx          # Original master table
└── backups/                   # Backup directory
    ├── Master-Table_backup_20250831_001408.xlsx
    ├── Master-Table_backup_20250831_001505.xlsx
    └── Master-Table_backup_20250831_001511.xlsx
```

## Implementation

### In excel_processor.py
The backup functionality has been integrated into all functions that modify the master table:

1. **`append_filtered_dataframe_to_master()`** - Line 520
2. **`a()` function** - Line 627 (master DataFrame append)
3. **`append_from_template()`** - Line 1267

### In transformer.py
- **`create_master_table_backup()`** - Standalone backup function
- **`safe_append_example()`** - Demonstrates safe append pattern
- **Example usage** - Shows how to use backup functionality

## Usage Examples

### Basic Backup Creation
```python
from pathlib import Path
from transformer import create_master_table_backup

master_path = Path("2-copy-reformat/Master-Table.xlsx")
backup_path = create_master_table_backup(master_path)

if backup_path:
    print(f"Backup created: {backup_path}")
else:
    print("Backup failed")
```

### Safe Append Pattern
```python
def safe_append_data(data_df, master_path):
    # 1. Create backup first
    backup_path = create_master_table_backup(master_path)
    if backup_path is None:
        print("Warning: Could not create backup, proceeding anyway...")
    
    # 2. Transform data
    transformed_data = transform_to_master_format(data_df)
    
    # 3. Append to master table
    # Your append logic here...
```

## Backup Function Details

### Function Signature
```python
def create_master_table_backup(master_path: Path) -> Optional[Path]:
```

### Parameters
- **master_path**: Path to the master table file to backup

### Returns
- **Path**: Path to the created backup file if successful
- **None**: If backup creation failed

### Error Handling
- Checks if master table exists before attempting backup
- Creates backup directory if it doesn't exist
- Handles file copy errors gracefully
- Provides informative error messages

## File Naming Convention

Backup files follow this naming pattern:
```
{original_stem}_backup_{timestamp}{original_extension}
```

Example:
- Original: `Master-Table.xlsx`
- Backup: `Master-Table_backup_20250831_001408.xlsx`

Where:
- `20250831` = Date (YYYYMMDD)
- `001408` = Time (HHMMSS)

## Best Practices

### 1. Always Backup Before Modifications
```python
# ✅ Good - Create backup first
backup_path = create_master_table_backup(master_path)
# Then modify the file...

# ❌ Bad - Modify without backup
# Directly modify master table
```

### 2. Verify Backup Success
```python
backup_path = create_master_table_backup(master_path)
if backup_path is None:
    print("Backup failed - consider stopping operation")
    return
```

### 3. Periodic Cleanup
- Backups accumulate over time
- Consider keeping only the last 10-20 backups
- Clean up old backups periodically to save disk space

### 4. Check Backup Integrity
```python
if backup_path and backup_path.exists():
    original_size = master_path.stat().st_size
    backup_size = backup_path.stat().st_size
    if original_size == backup_size:
        print("Backup verified")
```

## Integration Status

### ✅ Integrated Functions
- `append_filtered_dataframe_to_master()` in excel_processor.py
- `a()` function in excel_processor.py  
- `append_from_template()` in excel_processor.py
- All transformer.py examples

### 🔄 Automatic Execution
The backup functionality runs automatically when:
- Processing SharePoint files with `excel_processor.py`
- Appending filtered data to master table
- Using template-based append operations
- Running transformer.py examples

## Testing

### Run Backup Demo
```bash
python backup_demo.py
```

### Run Transformer Examples
```bash
python transformer.py
```

### Verify Backup Directory
Check `2-copy-reformat/backups/` for timestamped backup files.

## Recovery

If you need to restore from a backup:

1. **Identify the backup**: Look in `2-copy-reformat/backups/`
2. **Choose the right timestamp**: Select the backup from before the issue occurred
3. **Restore**: Copy the backup file over the current master table
4. **Verify**: Check that the restored file has the expected data

```bash
# Example recovery command
copy "2-copy-reformat\backups\Master-Table_backup_20250831_001408.xlsx" "2-copy-reformat\Master-Table.xlsx"
```

## Summary

The backup functionality provides:
- ✅ **Automatic protection** before any master table modifications
- ✅ **Timestamped organization** for easy identification
- ✅ **Error handling** with graceful degradation
- ✅ **Integration** into all append operations
- ✅ **Easy recovery** when needed

This ensures that your Master-Table.xlsx file is always protected before any modifications are made.
