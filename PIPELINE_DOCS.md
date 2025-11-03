# Student Assignment Processing Pipeline

This document outlines the logic and structure of the `StudentAssignmentProcessor` class and the Flask server application for processing student responses for course assignments. The application now provides two methods of operation: through a hosted front end or via a Jupyter notebook (calling the `process` method directly).

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
Allocates students to sectors for Round 1 and Round 2, ensuring that sector sizes do not exceed predefined limits. The logic includes:
- **Priority Handling**: Students are assigned to their preferred sectors based on the order of priority (1st, 2nd, 3rd).
- **Sector Size Management**: Each sector can accommodate a maximum number of students (`MAX_PER_SECTOR`). If a sector exceeds this limit, students are reassigned to their next preferred sector.
- **Balancing**: The algorithm aims to maintain a balance between the number of students assigned to each sector, with a maximum allowable difference defined by `MAX_DIFF`.

#### 4. `group_students_by_sector`
Groups students within each sector based on their assignments, aiming for balanced group sizes. The logic includes:
- **Group Size Constraints**: Groups are formed with a maximum size (`max_group_size`) and a minimum size (`min_group_size`).
- **Randomization**: Students are shuffled to avoid bias in grouping, ensuring a fair distribution of students.
- **Dynamic Grouping**: The number of groups is determined based on the total number of students assigned to a sector, creating groups that are as evenly sized as possible.

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

### Through a Jupyter Notebook

To run the processing pipeline directly in a notebook, create an instance of the `StudentAssignmentProcessor` and call the `process` method:

```python
processor = StudentAssignmentProcessor(response_file='responses.xlsx', classlist_file='classlist.xlsx')
processor.process(min_sector_size=10, output_file='output.xlsx')
```

### Through the Hosted Front End

The application can also be run as a Flask server. Here’s how to use the web interface:

1. **Run the Server**:
   Execute the `server.py` file to start the Flask server.

   ```bash
   python server.py
   ```

2. **Upload Files**:
   Navigate to `http://127.0.0.1:5000/` in your web browser. You will see a form to upload the response and class list files.

3. **Process Files**:
   After uploading both files, the application will process the data and redirect you to a download link for the processed Excel file.

4. **Download Output**:
   Click the provided link to download the processed output.

## Conclusion

This application provides a robust solution for assigning students to sectors based on their preferences, whether through a user-friendly web interface or directly in a programming environment. The design ensures flexibility and ease of use for managing student assignments effectively.
```