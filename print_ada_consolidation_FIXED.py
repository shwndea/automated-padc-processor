import pandas as pd
from pathlib import Path


# =============================================================================
# UTILITY FUNCTIONS FOR FILE DETECTION
# =============================================================================

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


# =============================================================================
# UTILITY FUNCTIONS FOR SEARCHING AND PROCESSING DATA
# =============================================================================

def find_rows_containing_program_name(student_data, program_name_to_find):
    """Find all rows that contain a specific program name."""
    matching_row_numbers = []
    
    for row_index, cell_value in enumerate(student_data.iloc[:, 1], start=1):
        if cell_value == program_name_to_find:
            matching_row_numbers.append(row_index)
    
    return matching_row_numbers


def find_rows_containing_month_number(student_data, month_number_to_find):
    """Find all rows that contain a specific month number."""
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


def find_program_boundary_rows(list_of_row_numbers):
    """Find where a program's data starts and ends."""
    if not list_of_row_numbers:
        return None, None
    
    first_row = min(list_of_row_numbers)
    last_row = max(list_of_row_numbers)
    
    return first_row, last_row


def extract_student_attendance_data(monthly_attendance_by_program, program_boundary_info, student_data):
    """Extract attendance data for each program and month combination."""
    attendance_data_dictionary = {}
    
    for month_number, rows_with_this_month in monthly_attendance_by_program.items():
        for current_row_number in rows_with_this_month:
            for program_code, boundary_info in program_boundary_info.items():
                program_start_row = boundary_info["start"]
                program_end_row = boundary_info["stop"]
                
                if program_start_row is not None and program_end_row is not None:
                    if program_start_row <= current_row_number <= program_end_row:
                        age_group = student_data.iloc[current_row_number - 1, 4]
                        month_value = student_data.iloc[current_row_number - 1, 2]
                        attendance_value = student_data.iloc[current_row_number - 1, 35]
                        
                        descriptive_field_name = f"{program_code}_Month_{month_value}_{age_group}: "
                        attendance_data_dictionary[descriptive_field_name] = attendance_value
    
    return attendance_data_dictionary


# =============================================================================
# MAIN PROGRAM EXECUTION
# =============================================================================

def print_ada_consolidation_fixed():
    """
    Performs steps 1-9 of the ADA audit process and prints all consolidated values.
    FIXED: Only processes months that actually exist in the data.
    """
    
    # =================================================================
    # STEP 1: Get user input
    # =================================================================
    print("=" * 60)
    print("🎓 ADA AUDIT CONFIGURATION")
    print("=" * 60)
    
    location = input("📍 Enter Location (e.g., TK-8, Elementary, Middle, High): ").strip()
    if not location:
        location = "TK-8"
        print(f"   Using default: {location}")
    
    school_year = input("📅 Enter School Year (e.g., 2025-2026, 2024-2025): ").strip()
    if not school_year:
        school_year = "2025-2026"
        print(f"   Using default: {school_year}")
    
    school_name = input("🏫 Enter School Name (e.g., CCCS, Lincoln Elementary): ").strip()
    if not school_name:
        school_name = "CCCS"
        print(f"   Using default: {school_name}")
    
    print(f"\n✅ Configuration:")
    print(f"   Location: {location}")
    print(f"   School Year: {school_year}")
    print(f"   School Name: {school_name}")
    print("=" * 60)
    
    # =================================================================
    # STEP 2: Define file paths and program information
    # =================================================================
    input_attendance_file = input("\n📂 Enter the full path to the Excel input file (or press Enter to auto-detect most recent): ").strip()
    
    if not input_attendance_file:
        # Auto-detect most recent file
        input_attendance_file = find_most_recent_attendance_file()
        
        if input_attendance_file:
            print(f"   ✅ Auto-detected most recent file:")
            print(f"      {Path(input_attendance_file).name}")
        else:
            print("   ❌ No attendance files found in Downloads folder!")
            print("   Please enter the full path to your attendance file.")
            return
    
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
    
    # =================================================================
    # STEP 3: Load the attendance data from Excel
    # =================================================================
    print("\n📊 Loading student attendance data from Excel...")
    student_attendance_data = pd.read_excel(input_attendance_file, header=None)
    
    # =================================================================
    # STEP 4: Find where each program's data starts and ends
    # =================================================================
    print("🔍 Locating program boundaries in the data...")
    
    program_boundaries = {}
    for short_code in program_name_mappings.values():
        program_boundaries[short_code] = {"start": None, "stop": None}
    
    for full_program_name, short_code in program_name_mappings.items():
        matching_rows = find_rows_containing_program_name(student_attendance_data, full_program_name)
        start_row, end_row = find_program_boundary_rows(matching_rows)
        program_boundaries[short_code]["start"] = start_row
        program_boundaries[short_code]["stop"] = end_row
    
    # =================================================================
    # STEP 5: Adjust boundaries to prevent overlaps
    # =================================================================
    print("🔧 Adjusting program boundaries to prevent overlaps...")

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
    
    # =================================================================
    # STEP 6: Display boundaries and allow user verification
    # =================================================================
    print("\n📍 Program boundary information:")
    for program_code, boundaries in program_boundaries.items():
        start = boundaries.get("start", "Not found")
        stop = boundaries.get("stop", "Not found") 
        print(f"  {program_code}: Start Row {start}, End Row {stop}")
    
    # Allow user to verify and correct boundaries if needed
    for program_code in program_boundaries.keys():
        user_response = input(
            f"\n❓ Are the boundaries for {program_code} correct? (yes/no): "
        ).lower().strip()
        
        if user_response == "no":
            while True:
                user_input = input(
                    f"📝 Enter new start and stop indices for {program_code} separated by a comma (e.g., 'start, stop'): "
                )
                new_indices = user_input.split(",")
                
                # Validate user input and update program_boundaries
                if len(new_indices) == 2:
                    start = (
                        int(new_indices[0].strip())
                        if new_indices[0].strip().lower() != "none"
                        else None
                    )
                    stop = (
                        int(new_indices[1].strip())
                        if new_indices[1].strip().lower() != "none"
                        else None
                    )
                    program_boundaries[program_code]["start"] = start
                    program_boundaries[program_code]["stop"] = stop
                    print(f"✅ Updated {program_code}: Start Row {start}, End Row {stop}")
                    break  # Exit the loop if input is valid
                else:
                    print(
                        "❌ Invalid input. Please enter valid integer start and stop indices separated by a comma."
                    )
    
    # Print final boundaries after user edits
    print("\n📍 Final program boundary information:")
    for program_code, boundaries in program_boundaries.items():
        start = boundaries.get("start", "Not found")
        stop = boundaries.get("stop", "Not found") 
        print(f"  {program_code}: Start Row {start}, End Row {stop}")
    
    # =================================================================
    # STEP 7: Find ONLY AVAILABLE month occurrences in the data
    # =================================================================
    print("\n📅 Finding month occurrences in attendance data...")
    
    monthly_attendance_by_program = {}
    available_months = []
    
    for month_number in range(1, 13):
        rows_with_this_month = find_rows_containing_month_number(student_attendance_data, month_number)
        
        if len(rows_with_this_month) > 0:
            monthly_attendance_by_program[month_number] = rows_with_this_month
            available_months.append(month_number)
            print(f"  ✅ Month {month_number}: Found in {len(rows_with_this_month)} rows")
        else:
            print(f"  ❌ Month {month_number}: NOT FOUND - Will be skipped in consolidation")
    
    print(f"\n📊 Summary: {len(available_months)} months available: {available_months}")
    
    # =================================================================
    # STEP 8: Extract all raw attendance data
    # =================================================================
    print("\n📈 Extracting attendance data for all programs and months...")
    
    raw_attendance_data = extract_student_attendance_data(
        monthly_attendance_by_program, 
        program_boundaries, 
        student_attendance_data
    )
    
    print(f"✅ Extracted {len(raw_attendance_data)} raw attendance data points")
    
    # =================================================================
    # STEP 9: Consolidate and PRINT all values (ONLY FOR AVAILABLE MONTHS)
    # =================================================================
    print("\n" + "=" * 80)
    print("🔄 CONSOLIDATED ATTENDANCE DATA (ONLY AVAILABLE MONTHS)")
    print("=" * 80)
    
    # FIXED: Only generate keys for months that actually exist in the data
    all_keys = []
    for parent_program in program_consolidation_rules.keys():
        for month in available_months:  # ✅ ONLY loop through available months
            for age_group in ["TK-3", "4-6", "7-8", "9-12"]:
                field_pattern = f"{parent_program}_Month_{month}_{age_group}: "
                all_keys.append(field_pattern)
    
    # Process and print each value
    for field_pattern in all_keys:
        # Extract parent program from field pattern
        parent_program = field_pattern.split('_Month_')[0]
        
        # Get consolidation rules for this program
        child_programs = program_consolidation_rules.get(parent_program, [parent_program])
        
        # Extract month and age group
        parts = field_pattern.replace(": ", "").split("_")
        month_idx = parts.index("Month") + 1
        month = parts[month_idx]
        age_group = "_".join(parts[month_idx + 1:])
        
        # Sum up values from all child programs
        total_value = 0
        component_strings = []
        
        for child_program in child_programs:
            child_field_pattern = f"{child_program}_Month_{month}_{age_group}: "
            child_value = raw_attendance_data.get(child_field_pattern, 0)
            
            if child_value and not pd.isna(child_value) and child_value != 0:
                total_value += child_value
                component_strings.append(f"{child_program}: {child_value}")
        
        # Print in the requested format
        if total_value > 0:
            components = " + ".join(component_strings) if component_strings else "0"
            print(f"{field_pattern} = {components} = {total_value}")
    
    print("=" * 80)
    print("✅ Consolidation complete!")
    print(f"📍 Configuration: {location}, {school_year}, {school_name}")
    print(f"📅 Processed months: {available_months}")


# =============================================================================
# RUN THE PROGRAM
# =============================================================================

if __name__ == "__main__":
    print_ada_consolidation_fixed()
