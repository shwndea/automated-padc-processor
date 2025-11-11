"""
Test script to display all program attendance values
"""
import pandas as pd
from ADA_Audit_25_26_IMPROVED import (
    find_rows_containing_program_name,
    find_rows_containing_month_number,
    find_program_boundary_rows,
    extract_student_attendance_data
)

# File path
input_attendance_file = (
    "C:\\Users\\Shawn\\Downloads\\PrintMonthlyAttendanceSummaryTotals_20251021_143005_82100f5.xlsx"
)

# Program name mappings
program_name_mappings = {
    "Program C Charter Resident": "Prog_C",
    "Program C Charter Resident -  Transitional Kindergarten(TK)": "Prog_C_TK",
    "Program C Charter Resident -  McClellan(CM)": "Prog_C_CM",
    "Program C Charter Resident -  Sac Youth Center(SYC)": "Prog_C_SYC",
    "Program N Non-Resident Charter": "Prog_N", 
    "Program N Non-Resident Charter -  Transitional Kindergarten(TK)": "Prog_N_TK",
    "Program N Non-Resident Charter -  McClellan(CM)": "Prog_N_CM",
    "Program N Non-Resident Charter -  Sac Youth Center(SYC)": "Prog_N_SYC",
    "Program J Indep Study Charter Resident": "Prog_J",
    "Program J Indep Study Charter Non-Resident -  Transitional Kindergarten(TK)": "Prog_J_TK",
    "Program K Indep Study Charter Non-Resident": "Prog_K",
    "Program K Indep Study Charter Non-Resident -  Transitional Kindergarten(TK)": "Prog_K_TK",
}

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

print("=" * 80)
print("📊 ATTENDANCE DATA DISPLAY TEST")
print("=" * 80)

# Load data
print("\n📂 Loading attendance data...")
student_attendance_data = pd.read_excel(input_attendance_file, header=None)

# Find program boundaries
print("🔍 Finding program boundaries...")
program_boundaries = {}
for short_code in program_name_mappings.values():
    program_boundaries[short_code] = {"start": None, "stop": None}

for full_program_name, short_code in program_name_mappings.items():
    matching_rows = find_rows_containing_program_name(student_attendance_data, full_program_name)
    start_row, end_row = find_program_boundary_rows(matching_rows)
    program_boundaries[short_code]["start"] = start_row
    program_boundaries[short_code]["stop"] = end_row

# Adjust boundaries
prog_C_tk_start = program_boundaries["Prog_C_TK"]["start"]
prog_N_start = program_boundaries["Prog_N"]["start"]

if prog_C_tk_start is not None and prog_N_start is not None:
    program_boundaries["Prog_C"]["stop"] = prog_C_tk_start - 1

if prog_N_start is not None:
    program_boundaries["Prog_C_TK"]["stop"] = prog_N_start - 1

prog_N_tk_start = program_boundaries["Prog_N_TK"]["start"]
if prog_N_tk_start is not None:
    program_boundaries["Prog_N"]["stop"] = prog_N_tk_start - 1

programs_to_adjust = ["Prog_N_TK", "Prog_J", "Prog_K"]
for i in range(len(programs_to_adjust) - 1):
    current_program = programs_to_adjust[i]
    next_program = programs_to_adjust[i + 1]
    
    current_start = program_boundaries[current_program]["start"]
    next_start = program_boundaries[next_program]["start"]
    
    if current_start is not None and next_start is not None:
        program_boundaries[current_program]["stop"] = next_start - 1

# Find month occurrences
print("📅 Finding month occurrences...")
monthly_attendance_by_program = {}
for month_number in range(1, 13):
    rows_with_this_month = find_rows_containing_month_number(student_attendance_data, month_number)
    monthly_attendance_by_program[month_number] = rows_with_this_month

# Extract raw data
print("📈 Extracting raw attendance data...")
raw_attendance_data = extract_student_attendance_data(
    monthly_attendance_by_program, 
    program_boundaries, 
    student_attendance_data
)

# Consolidate data
print("🔄 Consolidating data...")
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

# Display all values
print("\n" + "=" * 80)
print("📋 ALL CONSOLIDATED ATTENDANCE VALUES")
print("=" * 80)

# Group by program and display
programs = ["Prog_C", "Prog_C_TK", "Prog_N", "Prog_N_TK", "Prog_J", "Prog_J_TK", "Prog_K", "Prog_K_TK"]
age_groups = ["TK-3", "4-6", "7-8", "9-12"]

for program in programs:
    print(f"\n{'='*80}")
    print(f"🎯 {program}")
    print(f"{'='*80}")
    
    for month in range(1, 13):
        print(f"\n  📅 Month {month}:")
        for age_group in age_groups:
            field_name = f"{program}_Month_{month}_{age_group}: "
            value = consolidated_attendance_data.get(field_name, 0)
            
            # Highlight the specific example requested
            if field_name == "Prog_J_Month_11_7-8: ":
                print(f"    ⭐ {field_name} = {value} ⭐")
            else:
                print(f"    {field_name} = {value}")

# Display summary statistics
print("\n" + "=" * 80)
print("📊 SUMMARY STATISTICS")
print("=" * 80)

for program in programs:
    total = 0
    non_zero_count = 0
    
    for month in range(1, 13):
        for age_group in age_groups:
            field_name = f"{program}_Month_{month}_{age_group}: "
            value = consolidated_attendance_data.get(field_name, 0)
            if value and not pd.isna(value) and value != 0:
                total += value
                non_zero_count += 1
    
    print(f"\n{program}:")
    print(f"  Total attendance: {total}")
    print(f"  Non-zero entries: {non_zero_count}/48")
    if non_zero_count > 0:
        print(f"  Average (non-zero): {total/non_zero_count:.2f}")

print("\n" + "=" * 80)
print("✅ Test complete!")
print("=" * 80)
