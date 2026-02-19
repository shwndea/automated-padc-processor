"""
Run ADA Audit process with configuration from JSON file
"""
import json
import pandas as pd
import openpyxl
from pathlib import Path

# Import functions from the main script
from ADA_Audit_25_26_IMPROVED import (
    find_rows_containing_program_name,
    find_rows_containing_month_number,
    find_program_boundary_rows,
    extract_student_attendance_data,
    write_all_attendance_data_to_excel_efficiently
)

# Load configuration
config_path = r"C:\Users\Shawn\Desktop\GCC_AI\automated-padc-processor\boundary_settings\VA_M1_M2.json"
with open(config_path, 'r') as f:
    config = json.load(f)

print("=" * 60)
print("🎓 ADA AUDIT - RUNNING WITH CONFIGURATION")
print("=" * 60)
print(f"Configuration: {config['name']}")
print(f"Description: {config.get('description', 'N/A')}")
print("=" * 60)

# Configuration values
location = input("📍 Enter Location (e.g., TK-8, Elementary, Middle, High) [default: TK-8]: ").strip() or "TK-8"
school_year = input("📅 Enter School Year (e.g., 2025-2026) [default: 2025-2026]: ").strip() or "2025-2026"
school_name = input("🏫 Enter School Name (e.g., CCCS) [default: CCCS]: ").strip() or "CCCS"

# Ask for school type
print("\n🎯 School Grade Configuration:")
school_type = input("   Is this a TK-12 or K-12 school? (Enter 'TK' or 'K') [default: TK]: ").strip().upper() or "TK"
if school_type not in ['TK', 'K']:
    school_type = 'TK'
    print(f"   Invalid input. Using default: TK-12")

use_tk_12 = (school_type == 'TK')

print(f"\n✅ Configuration:")
print(f"   Location: {location}")
print(f"   School Year: {school_year}")
print(f"   School Name: {school_name}")
print(f"   School Type: {school_type}-12")
print("=" * 60)

# File paths
input_file = input("\n📂 Enter input attendance file path: ").strip().strip('"').strip("'")
if not input_file:
    input_file = r"C:\Users\Shawn\Downloads\PrintMonthlyAttendanceSummaryTotals_20251021_143005_82100f5.xlsx"
    print(f"   Using default: {input_file}")

output_file = input("📂 Enter output audit file path: ").strip().strip('"').strip("'")
if not output_file:
    output_file = r"C:\Users\Shawn\Downloads\2025-2026_I4C_ADA_Reconciliation.xlsx"
    print(f"   Using default: {output_file}")

worksheet_name = input("📄 Enter worksheet name [default: Template- Apportionment Summary]: ").strip() or "Template- Apportionment Summary"

# Load the program boundaries from config
program_boundaries = config['program_boundaries']
program_name_mappings = config['program_mappings']

print("\n" + "=" * 60)
print("📊 LOADING DATA")
print("=" * 60)

# Load student data
print(f"Loading attendance data from: {input_file}")
student_attendance_data = pd.read_excel(input_file, header=None)
print(f"✅ Loaded {len(student_attendance_data)} rows of data")

# Consolidation rules
program_consolidation_rules = {
    "Prog_C": ["Prog_C", "Prog_C_CM", "Prog_C_SYC"],
    "Prog_C_TK": ["Prog_C_TK"],
    "Prog_N": ["Prog_N", "Prog_N_CM", "Prog_N_SYC"],
    "Prog_N_TK": ["Prog_N_TK"],
    "Prog_J": ["Prog_J"],
    "Prog_J_TK": ["Prog_J_TK"],
    "Prog_K": ["Prog_K"],
    "Prog_K_TK": ["Prog_K_TK"],
}

print("\n" + "=" * 60)
print("🔍 PROCESSING ATTENDANCE DATA")
print("=" * 60)

# Find month occurrences
print("Finding month occurrences...")
monthly_attendance_by_program = {}
for month_number in range(1, 13):
    rows_with_this_month = find_rows_containing_month_number(student_attendance_data, month_number)
    monthly_attendance_by_program[month_number] = rows_with_this_month
    print(f"  Month {month_number}: Found in {len(rows_with_this_month)} rows")

# Extract attendance data
print("\nExtracting attendance data...")
raw_attendance_data = extract_student_attendance_data(
    monthly_attendance_by_program,
    program_boundaries,
    student_attendance_data
)
print(f"✅ Extracted {len(raw_attendance_data)} raw attendance data points")

# Consolidate data
print("\n🔄 Consolidating sub-location data with parent programs...")
print("   Program C Total = Main Program C + McClellan (CM) + Sac Youth Center (SYC)")
print("   Program N Total = Main Program N + McClellan (CM) + Sac Youth Center (SYC)")

consolidated_attendance_data = {}

for parent_program, child_programs in program_consolidation_rules.items():
    for month in range(1, 13):
        for age_group in ["TK-3", "4-6", "7-8", "9-12"]:
            field_pattern = f"{parent_program}_Month_{month}_{age_group}: "
            total_value = 0
            
            for child_program in child_programs:
                child_field_pattern = f"{child_program}_Month_{month}_{age_group}: "
                child_value = raw_attendance_data.get(child_field_pattern, 0)
                
                if child_value and not pd.isna(child_value) and child_value != 0:
                    total_value += child_value
            
            consolidated_attendance_data[field_pattern] = total_value

print(f"✅ Consolidated {len(consolidated_attendance_data)} attendance data points")

# Write to Excel
print("\n" + "=" * 60)
print("💾 WRITING TO EXCEL")
print("=" * 60)
school_type_name = "TK-12" if use_tk_12 else "K-12"
print(f"Using {school_type_name} cell mapping...")

write_all_attendance_data_to_excel_efficiently(
    consolidated_attendance_data,
    output_file,
    worksheet_name,
    use_tk_12
)

print("\n" + "=" * 60)
print("🎉 AUDIT COMPLETED SUCCESSFULLY!")
print("=" * 60)
print(f"📊 Results saved to: {output_file}")
print(f"📍 Configuration: {location}, {school_year}, {school_name}, {school_type}-12")
print(f"📝 Used configuration: {config['name']}")
print("\n💡 Note: McClellan (CM) and Sac Youth Center (SYC) totals have been")
print("   automatically added to their respective parent program totals.")
print("=" * 60)
