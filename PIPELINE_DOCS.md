# Student Assignment Processing Pipeline

This document outlines the logic and structure of the `StudentAssignmentProcessor` class implemented in Python. This class processes student responses for course assignments, organizes them into sectors, and manages group assignments. Below is a detailed explanation of the pipeline logic.

## Overview

The main purpose of this pipeline is to:

1. **Merge Responses**: Combine student responses with class data.
2. **Extract Sectors**: Identify valid sectors from student preferences.
3. **Assign Sectors**: Allocate students to sectors based on their priorities while maintaining sector size limits.
4. **Group Students**: Organize students into groups for each assigned sector.
5. **Update Class DataFrames**: Reflect the assignments in the original class DataFrames.
6. **Write to Excel**: Output the processed data to an Excel file.

## Class Structure

### Initialization

The `StudentAssignmentProcessor` class is initialized with paths to response and class list files. It defines several properties, including the sectors available for assignment and maximum limits for sector sizes.

### Methods

#### 1. `merge_responses`
Merges student responses into class DataFrames, ensuring that each class only contains unique entries based on the timestamp.

#### 2. `extract_sectors`
Extracts valid sector preferences from the priority columns (1st, 2nd, and 3rd) in the DataFrame.

#### 3. `assign_sectors`
Allocates students to sectors for Round 1 and Round 2, ensuring that sector sizes do not exceed predefined limits. It handles cases where students may need to be reassigned to balance groups.

#### 4. `group_students_by_sector`
Groups students within each sector based on their assignments, aiming for a balance between maximum and minimum group sizes.

#### 5. `update_class_dfs`
Updates the original class DataFrames with the assigned sectors and groups for each student.

#### 6. `write_to_excel`
Outputs the processed data to an Excel file, creating individual sheets for each class and sector, as well as an overview of group assignments.

### Main Processing Flow

1. **Load Data**: Read the response and class list files from Excel.
2. **Merge Responses**: Combine response data with class lists using `merge_responses`.
3. **Extract Valid Sectors**: Use `extract_sectors` to find valid preferences from students.
4. **Sector Assignment**: Call `assign_sectors` to allocate sectors based on student preferences while respecting size limits.
5. **Group Assignment**: Use `group_students_by_sector` to create balanced groups.
6. **Update DataFrames**: Reflect the assignment results in the original class DataFrames.
7. **Save Results**: Write the final output to an Excel file using `write_to_excel`.

## Example Usage

To run the processing pipeline, create an instance of the `StudentAssignmentProcessor` and call the `process` method:

```python
processor = StudentAssignmentProcessor(response_file='responses.xlsx', classlist_file='classlist.xlsx')
processor.process(min_sector_size=10, output_file='output.xlsx')