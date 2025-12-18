# House Sorter

A Python program that assigns students to four houses (sapphire, topaz, agate, ruby) while maintaining existing assignments and balancing the distribution.

## Features

- Reads student data from Excel files
- Preserves existing house assignments from previous years
- Randomly assigns unassigned students to balance house populations
- Generates a new Excel file with complete assignments
- Shows before/after distribution statistics

## Installation

1. Make sure you have Python 3 installed
2. Install required packages:
   ```bash
   pip3 install -r requirements.txt
   ```

## Excel File Format

Your input Excel file should have these columns:

| Student Name | House |
|--------------|-------|
| Alice Johnson | sapphire |
| Bob Smith | topaz |
| Carol Williams | |
| David Brown | agate |

- **Student Name**: The student's full name
- **House**: Either one of the four house names (sapphire, topaz, agate, ruby) or empty/blank if not yet assigned

**Note**: Leave the House column empty or blank for students who need to be assigned.

## Usage

### Option 1: Run with file as argument
```bash
python3 house_sorter.py your_students.xlsx
```

### Option 2: Run and enter filename when prompted
```bash
python3 house_sorter.py
```
Then type the filename when asked.

## Output

The program will:
1. Display current house distribution
2. Show how many students need to be assigned to each house
3. Create a new file named `your_students_sorted.xlsx` with all assignments
4. Display the final balanced distribution

Example output:
```
=== Current House Distribution ===
sapphire: 3 students
Topaz: 2 students
Agate: 2 students
Ruby: 2 students

Total students: 20
Already assigned: 9
Need to assign: 11
Target per house: 5.0

=== Final House Distribution ===
sapphire: 5 students
Topaz: 5 students
Agate: 5 students
Ruby: 5 students

✓ Results saved to: your_students_sorted.xlsx
```

## Example

A sample file `sample_students.xlsx` is included to demonstrate the format. Try it:

```bash
python3 house_sorter.py sample_students.xlsx
```

This will create `sample_students_sorted.xlsx` with all students assigned to houses.

## Notes

- If the total number of students doesn't divide evenly by 4, some houses will have one more student than others
- Assignments are random for unassigned students
- The original file is never modified - a new "_sorted" file is always created
- House names are case-insensitive
