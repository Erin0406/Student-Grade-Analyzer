# Student Grades Management System (Excel VBA + MS Access)

## Overview

This workbook is an Excel-based tool built with VBA UserForms to manage, analyze, and visualize student grade data sourced from an MS Access database (Registrar.mdb). It supports importing data, viewing course structures, computing statistics, generating distributions, and producing formal reports.

## System Requirements
- Microsoft Excel (Desktop version with VBA support)
- Macros enabled
- Microsoft Access Database Engine (for .mdb file support)
- Source database file: Registrar.mdb

##How to Start
- Open the Excel workbook.
- Enable macros when prompted.
- Go to the custom ribbon tab: "Students Grades"
- Click "Open Student Grades" to launch the main interface.

## Workflow Rules (Important)
- Nothing works until Import Data is completed.
- Most features depend on specific prior steps (explained below).
- Some outputs overwrite previous results (notably charts and statistics sheets).

## Features

1. Import Data (Mandatory First Step)

Path: Students Grades → Import Data → Continue

- Opens the Import Database File UserForm.
- User selects the Registrar.mdb file.
- Imports all relevant tables into the Excel workbook.
- Enables all other system functions.

Without this step, all other features remain disabled.

2. List Courses

Path: Students Grades → List Courses → Continue

- Displays:
  - Subject codes (e.g., CP212)
  - Subject descriptions
- Shows all courses imported from the database.

Dependency: Requires successful data import only.

3. Display Class Average

Path: Students Grades → Display Class Average

- Opens a dedicated UserForm.
- Allows selection of:
  - Specific course(s)
  - Individual assessment types (Assignments, Exams, etc.)
  - Or all subjects at once

Output:
- Generates a worksheet named: ClassAverage
- Displays computed averages per assessment and/or course

Dependency: Requires data import only.

4. Course Enrollment Analysis

Path: Students Grades → Course Enrollment → Continue

- User inputs a valid course code (e.g., CP212).
- System validates format and retrieves enrolled students.

Output:

- Creates worksheet: Enrollment
- Displays:
  - Student names
  - Student codes
  - Grades for all assessment components
  - All enrolled sections for that course

Dependency: Requires data import only.

5. Create Chart (Statistical Analysis)

Path: Students Grades → Create Chart → Continue

- User selects:
  - Course (must already be processed via Course Enrollment)
  - Assessment table (e.g., A1, A2, FinalExam)

Output:

- Distribution chart for selected dataset
- Worksheet: Statistics
  - Mean
  - Minimum
  - Maximum
  - Additional summary metrics
- Optional worksheet:
  - FinalExam_Distribution (or similar based on selection)

Important Constraints:

- Designed for one chart at a time
- Generating multiple charts without cleanup may cause errors or overwrite outputs
- Manual deletion of previous distribution sheets is recommended before generating new charts

Dependency: Requires Course Enrollment to be completed first.

6. Generate Report (Word Export)

Path: Students Grades → Generate Report → Continue

- Generates a Microsoft Word document report based on the most recent chart.

Includes:

- Chart interpretation
- Statistical summary
- Written analysis and commentary
- References to generated Excel outputs

Dependency: Requires Create Chart to be completed first.

## Dependency Chain (Critical)
Import Data
   ↓
(List Courses / Display Class Average / Course Enrollment)
   ↓
Course Enrollment
   ↓
Create Chart
   ↓
Generate Report


## Known Limitations
- Only one active chart/statistics view is supported at a time.
- Re-running chart generation without cleanup may overwrite or break outputs.
- Requires strict course code formatting (e.g., CP212).
- Performance depends on size of Registrar.mdb.

## Notes for Users
- Always ensure data import is successful before proceeding.
- Follow the dependency order to avoid errors.
- Do not attempt parallel chart generation.
- Recommended to save workbook after each major step.

## Purpose
This tool is designed to streamline academic data handling by combining:

- Database import (MS Access)
- Excel-based analytics
- Automated charting
- Report generation (Word integration)

It reduces manual grade analysis workload and standardizes reporting across courses.
