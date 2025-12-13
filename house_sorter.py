import pandas as pd
import random
from pathlib import Path
import sys


def read_student_data(input_file):
    """Read student data from Excel file"""
    try:
        df = pd.read_excel(input_file)
        return df
    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found!")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)


def validate_data(df):
    """Validate the input data has required columns"""
    required_columns = ['Student Name', 'House']
    
    for col in required_columns:
        if col not in df.columns:
            print(f"Error: Column '{col}' not found in the Excel file!")
            print(f"Required columns: {required_columns}")
            sys.exit(1)
    
    return True


def sort_students(df):
    """Sort students into houses, respecting existing assignments"""
    
    # Define house names
    houses = ['saphine', 'topaz', 'agat', 'ruby']
    
    # Normalize house names in the dataframe (handle case and whitespace)
    df['House'] = df['House'].apply(lambda x: str(x).strip().lower() if pd.notna(x) else None)
    
    # Separate assigned and unassigned students
    assigned = df[df['House'].isin(houses)].copy()
    unassigned = df[~df['House'].isin(houses)].copy()
    
    # Count current house populations
    house_counts = {house: 0 for house in houses}
    for house in houses:
        house_counts[house] = len(assigned[assigned['House'] == house])
    
    print("\n=== Current House Distribution ===")
    for house, count in house_counts.items():
        print(f"{house.capitalize()}: {count} students")
    
    # Calculate total students and target per house
    total_students = len(df)
    unassigned_count = len(unassigned)
    target_per_house = total_students / len(houses)
    
    print(f"\nTotal students: {total_students}")
    print(f"Already assigned: {len(assigned)}")
    print(f"Need to assign: {unassigned_count}")
    print(f"Target per house: {target_per_house:.1f}")
    
    # Calculate how many students each house needs
    house_needs = {}
    for house in houses:
        needed = int(target_per_house) - house_counts[house]
        # Handle rounding - some houses might need one more
        if needed < 0:
            needed = 0
        house_needs[house] = needed
    
    # Adjust for rounding issues
    total_needed = sum(house_needs.values())
    if total_needed < unassigned_count:
        # Distribute extra students to houses with fewer students
        remaining = unassigned_count - total_needed
        sorted_houses = sorted(houses, key=lambda h: house_counts[h])
        for i in range(remaining):
            house_needs[sorted_houses[i % len(houses)]] += 1
    elif total_needed > unassigned_count:
        # Reduce from houses that need the most
        excess = total_needed - unassigned_count
        sorted_houses = sorted(houses, key=lambda h: house_needs[h], reverse=True)
        for i in range(excess):
            if house_needs[sorted_houses[i % len(houses)]] > 0:
                house_needs[sorted_houses[i % len(houses)]] -= 1
    
    print("\n=== Students Needed Per House ===")
    for house, need in house_needs.items():
        print(f"{house.capitalize()}: {need} more students")
    
    # Create list of house assignments for unassigned students
    assignment_pool = []
    for house, need in house_needs.items():
        assignment_pool.extend([house] * need)
    
    # Shuffle for random assignment
    random.shuffle(assignment_pool)
    
    # Assign houses to unassigned students
    for idx, student_idx in enumerate(unassigned.index):
        if idx < len(assignment_pool):
            df.at[student_idx, 'House'] = assignment_pool[idx]
    
    return df


def save_results(df, output_file):
    """Save the sorted students to a new Excel file"""
    try:
        df.to_excel(output_file, index=False)
        print(f"\n✓ Results saved to: {output_file}")
    except Exception as e:
        print(f"Error saving file: {e}")
        sys.exit(1)


def print_final_distribution(df):
    """Print final house distribution"""
    houses = ['saphine', 'topaz', 'agat', 'ruby']
    
    print("\n=== Final House Distribution ===")
    for house in houses:
        count = len(df[df['House'] == house])
        print(f"{house.capitalize()}: {count} students")
    
    print(f"\nTotal: {len(df)} students")


def main():
    """Main function"""
    print("=" * 50)
    print("     HOUSE SORTER")
    print("=" * 50)
    
    # Get input file
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = input("Enter the input Excel file name (e.g., students.xlsx): ").strip()
    
    # Read data
    print(f"\nReading file: {input_file}")
    df = read_student_data(input_file)
    
    # Validate data
    validate_data(df)
    
    # Sort students
    df_sorted = sort_students(df)
    
    # Generate output filename
    input_path = Path(input_file)
    output_file = input_path.stem + "_sorted" + input_path.suffix
    
    # Save results
    save_results(df_sorted, output_file)
    
    # Print final distribution
    print_final_distribution(df_sorted)
    
    print("\n" + "=" * 50)
    print("     COMPLETE!")
    print("=" * 50)


if __name__ == "__main__":
    main()
