# Grade Processing System

A Python-based grade management system that collects, validates, and processes student grades with support for multiple assessment categories and automatic report generation.

## Features

- **Grade Collection**: Input assignment names, grades, categories (Formative/Summative), and weights
- **Data Validation**: Validates grades (0-100), categories (FA/SA), and weights (positive numbers)
- **Database Storage**: Stores all grades in SQLite database with automatic timestamping
- **Grade Calculation**: 
  - Weighted grade computation
  - Final grade calculation
  - GPA conversion (0-5.0 scale)
  - Pass/Fail determination based on minimum thresholds
- **Report Export**: 
  - Excel reports with formatted sheets (Assignments & Summary)
  - Color-coded results (green for pass ≥50, red for fail <50)
  - CSV export of all grades
- **Archive System**: Automatically archives databases for record-keeping

## Requirements

- Python 3.x
- `openpyxl` (for Excel export)
- `sqlite3` (built-in)
- `csv` (built-in)

## Installation

```bash
pip install openpyxl
```

## Usage

1. Run the program:
```bash
python grade_processing_system.py
```

2. Follow the prompts to enter assignment details:
   - Assignment name
   - Category (FA for Formative or SA for Summative)
   - Grade (0-100)
   - Weight (positive number)

3. Add multiple assignments by confirming "yes" when prompted

4. After all assignments are entered, the system will:
   - Calculate weighted grades for each assessment
   - Compute final grade and GPA
   - Determine pass/fail status
   - Generate Excel and CSV reports
   - Archive the database

## Output Files

- `grades_report_[TIMESTAMP].xlsx` - Formatted Excel report with assignments and summary
- `grades_[TIMESTAMP].csv` - CSV export of all grades
- `archives/grades_[TIMESTAMP].db` - SQLite database backup

## Grade Calculation

- **Weighted Grade**: (Grade / 100) × Weight
- **Final Grade**: (Total Points / Total Possible Points) × 100
- **GPA**: (Final Grade / 100) × 5.0
- **Pass Condition**: FA ≥ 50% of total FA weight AND SA ≥ 50% of total SA weight

## Notes

- Grades below 50 are highlighted in reports and require resubmission
- All timestamps use YYYYMMDD_HHMMSS format
- Database files are automatically archived in the `archives/` folder