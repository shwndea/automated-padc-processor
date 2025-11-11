import pandas as pd
from pathlib import Path


def find_most_recent_attendance_file():
    """Find the most recent PrintMonthlyAttendanceSummaryTotals file in Downloads"""
    downloads_dir = Path("C:\\Users\\Shawn\\Downloads")
    
    if not downloads_dir.exists():
        return None
    
    # Find all attendance summary files
    attendance_files = list(downloads_dir.glob("PrintMonthlyAttendanceSummaryTotals_*.xlsx"))
    
    if not attendance_files:
        return None
    
    # Get the most recent file by modification time
    most_recent = max(attendance_files, key=lambda f: f.stat().st_mtime)
    return str(most_recent)


def find_rows_containing_month_number(student_data, month_number_to_find):
    """Find all rows that contain a specific month number in Column C."""
    matching_row_numbers = []
    
    for row_index, cell_value in enumerate(student_data.iloc[:, 2], start=1):
        if pd.isna(cell_value):
            continue
        
        try:
            if int(cell_value) == month_number_to_find:
                matching_row_numbers.append(row_index)
        except ValueError:
            continue
    
    return matching_row_numbers


def check_available_months():
    """
    Checks which months (1-12) are actually available in Column C of the Excel file.
    Only returns months that have at least one row of data.
    """
    
    print("=" * 60)
    print("📅 CHECKING AVAILABLE MONTHS IN DATA")
    print("=" * 60)
    
    # =================================================================
    # Get input file from user or auto-detect most recent
    # =================================================================
    input_attendance_file = input("\n📂 Enter the full path to the Excel file (or press Enter to auto-detect most recent): ").strip()
    
    if not input_attendance_file:
        # Auto-detect most recent file
        input_attendance_file = find_most_recent_attendance_file()
        
        if input_attendance_file:
            print(f"   ✅ Auto-detected most recent file:")
            print(f"      {input_attendance_file}")
        else:
            print("   ❌ No attendance files found in Downloads folder!")
            return None, None, None
    
    print(f"\n📊 Loading data from:")
    print(f"   {input_attendance_file}")
    
    student_attendance_data = pd.read_excel(input_attendance_file, header=None)
    
    print(f"   ✅ Loaded {len(student_attendance_data)} rows")
    
    # =================================================================
    # Check each month from 1 to 12
    # =================================================================
    print("\n🔍 Scanning Column C for month numbers 1-12...")
    print("-" * 60)
    
    available_months = []
    unavailable_months = []
    month_row_details = {}
    
    for month_number in range(1, 13):
        rows_with_this_month = find_rows_containing_month_number(student_attendance_data, month_number)
        
        if len(rows_with_this_month) > 0:
            available_months.append(month_number)
            month_row_details[month_number] = rows_with_this_month
            print(f"  ✅ Month {month_number:2d}: Found in {len(rows_with_this_month):3d} rows - {rows_with_this_month[:5]}{'...' if len(rows_with_this_month) > 5 else ''}")
        else:
            unavailable_months.append(month_number)
            print(f"  ❌ Month {month_number:2d}: NOT FOUND (0 rows)")
    
    # =================================================================
    # Summary
    # =================================================================
    print("\n" + "=" * 60)
    print("📊 SUMMARY")
    print("=" * 60)
    
    print(f"\n✅ AVAILABLE MONTHS ({len(available_months)} months):")
    print(f"   {available_months}")
    
    if unavailable_months:
        print(f"\n❌ UNAVAILABLE MONTHS ({len(unavailable_months)} months):")
        print(f"   {unavailable_months}")
        print(f"\n⚠️  WARNING: These months should NOT be processed!")
        print(f"   The consolidation script should skip these months entirely.")
    
    print("\n" + "=" * 60)
    print("💡 RECOMMENDATION")
    print("=" * 60)
    print("\nThe consolidation script should:")
    print(f"  1. Only loop through: {available_months}")
    print(f"  2. Skip months: {unavailable_months}")
    print(f"  3. This prevents adding fake '0' values for non-existent months")
    
    # =================================================================
    # Detailed breakdown by month
    # =================================================================
    print("\n" + "=" * 60)
    print("📋 DETAILED MONTH BREAKDOWN")
    print("=" * 60)
    
    for month_number in available_months:
        rows = month_row_details[month_number]
        print(f"\n📅 Month {month_number}:")
        print(f"   Total rows: {len(rows)}")
        print(f"   Row numbers: {rows}")
        
        # Show sample data from first row
        if rows:
            first_row = rows[0]
            print(f"\n   Sample from Row {first_row}:")
            print(f"     Column B (Program): {student_attendance_data.iloc[first_row - 1, 1]}")
            print(f"     Column C (Month):   {student_attendance_data.iloc[first_row - 1, 2]}")
            print(f"     Column E (Age):     {student_attendance_data.iloc[first_row - 1, 4]}")
            print(f"     Column AJ (Value):  {student_attendance_data.iloc[first_row - 1, 35]}")
    
    return available_months, unavailable_months, month_row_details


if __name__ == "__main__":
    available, unavailable, details = check_available_months()
    
    print("\n" + "=" * 60)
    print("✅ Check complete!")
    print("=" * 60)
