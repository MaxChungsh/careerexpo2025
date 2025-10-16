import pandas as pd
from collections import defaultdict, Counter
import random
from openpyxl.styles import PatternFill
import os

class StudentAssignmentProcessor:
    """Processes student responses and assigns sectors for Round 1 and Round 2."""
    
    def __init__(self, response_file, classlist_file):
        """Initialize with input and output file paths."""
        self.response_file = response_file
        self.classlist_file = classlist_file
        self.sectors = [
            'Allied Health (PT, OT, RT)', 'Architecture', 'Arts', 'Banking & Finance',
            'Computer Science', 'Education', 'Engineering', 'Law', 'Medical',
            'Social Science and Social Work'
        ]
        self.sector_dict = {s.lower(): s for s in self.sectors}
        self.classes = ['5A', '5B', '5C', '5D', '6A', '6B', '6C', '6D']
        self.MAX_PER_SECTOR = 48
        self.MAX_DIFF = 10
        self.yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        self.red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        self.grey_fill = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')
    
    def merge_responses(self, response_df, class_dfs):
        """
        Merge response data into class DataFrames and create combined All F5 F6 DataFrame.
        
        Args:
            response_df (pd.DataFrame): Response data.
            class_dfs (dict): Dictionary of class DataFrames (key: class name, value: DataFrame).
        
        Returns:
            tuple: (updated class_dfs, all_df)
        """
        response_df['Class'] = response_df['Class'].str.upper()
        response_df['Class No.'] = pd.to_numeric(response_df['Class No.'], errors='coerce')
        f56_response = response_df[response_df['Class'].str.startswith(('5', '6'))].copy()
        f56_response = f56_response.sort_values(by='Timestamp', ascending=False)
        f56_response = f56_response.drop_duplicates(subset=['Class', 'Class No.'], keep='first')
        
        for cls in self.classes:
            df = class_dfs[cls]
            for idx, row in df.iterrows():
                class_ = row['Class'].upper()
                no = row['Cls No']
                match = f56_response[(f56_response['Class'] == class_) & (f56_response['Class No.'] == no)]
                if not match.empty:
                    m = match.iloc[0]
                    df.at[idx, '1st Priority'] = m['1st Priority']
                    df.at[idx, '2nd Priority'] = m['2nd Priority']
                    df.at[idx, '3rd Priority'] = m['3rd Priority']
                    df.at[idx, 'Particular Course'] = m['Particular Course']
                    df.at[idx, 'Topics and Suggestions'] = m['Topics and Suggestions']
                    df.at[idx, 'Other Majors'] = m['Other Majors']
        
        all_df = pd.concat(class_dfs.values(), ignore_index=True)
        all_df = all_df[[col for col in all_df.columns if not (col.startswith('FT') or col.startswith('AFT'))]]
        
        return class_dfs, all_df
    
    def extract_sectors(self, df):
        """
        Extract valid sectors from priority columns.
        
        Args:
            df (pd.DataFrame): DataFrame with priority columns.
        
        Returns:
            pd.DataFrame: DataFrame with added p1_list, p2_list, p3_list columns.
        """
        def get_sectors(entry):
            if pd.isnull(entry) or entry == '':
                return []
            parts = [p.strip().lower() for p in entry.split(',')]
            matched = []
            for p in parts:
                if 'allied health' in p and '(pt' in p:
                    matched.append(self.sectors[0])
                elif p in self.sector_dict:
                    matched.append(self.sector_dict[p])
            return matched
        
        df['p1_list'] = df['1st Priority'].apply(get_sectors)
        df['p2_list'] = df['2nd Priority'].apply(get_sectors)
        df['p3_list'] = df['3rd Priority'].apply(get_sectors)
        return df
    
    def assign_sectors(self, assign_df, min_sector_size=0):
        """
        Assign students to sectors for Round 1 and Round 2, discarding sectors with fewer than min_sector_size students.
        
        Args:
            assign_df (pd.DataFrame): DataFrame with students to assign.
            min_sector_size (int): Minimum number of students in a sector across both rounds (default 0).
        
        Returns:
            list: List of student dictionaries with assignments.
        """
        students = assign_df.to_dict(orient='records')
        
        for s in students:
            s['round1'] = None
            s['round2'] = None
            s['priority_used_r1'] = None
            s['priority_used_r2'] = None
        
        sector_prefs = defaultdict(list)
        for s in students:
            for p, plist in [(1, s['p1_list']), (2, s['p2_list'])]:
                for sec in plist:
                    sector_prefs[sec].append((s, p))
        
        for sec in self.sectors:
            candidates = sector_prefs[sec]
            if not candidates:
                continue
            candidates.sort(key=lambda x: x[1])
            total = len(candidates)
            target_per_round = total // 2
            r1_count = 0
            assigned = set()
            
            for s, p in candidates:
                student_id = (s['Class'], s['Cls No'])
                if r1_count < target_per_round + (total % 2) and student_id not in assigned:
                    if not s['round1'] or s['priority_used_r1'] > p:
                        s['round1'] = sec
                        s['priority_used_r1'] = p
                        r1_count += 1
                        assigned.add(student_id)
            
            for s, p in candidates:
                student_id = (s['Class'], s['Cls No'])
                if student_id not in assigned and (not s['round2'] or s['priority_used_r2'] > p):
                    if s['round1'] != sec:
                        s['round2'] = sec
                        s['priority_used_r2'] = p
                        assigned.add(student_id)
        
        for s in students:
            student_id = (s['Class'], s['Cls No'])
            if not s['round1']:
                for sec in s['p3_list']:
                    r1_counts = Counter(st['round1'] for st in students if st['round1'])
                    if r1_counts.get(sec, 0) < self.MAX_PER_SECTOR:
                        s['round1'] = sec
                        s['priority_used_r1'] = 3
                        break
            if not s['round2']:
                for sec in s['p3_list']:
                    if sec != s['round1']:
                        r2_counts = Counter(st['round2'] for st in students if st['round2'])
                        if r2_counts.get(sec, 0) < self.MAX_PER_SECTOR:
                            s['round2'] = sec
                            s['priority_used_r2'] = 3
                            break
        
        if min_sector_size > 0:
            for sec in self.sectors:
                r1_stds = [s for s in students if s['round1'] == sec]
                r2_stds = [s for s in students if s['round2'] == sec]
                total = len(r1_stds) + len(r2_stds)
                if total < min_sector_size:
                    for s in r1_stds:
                        for p, plist in [(1, s['p1_list']), (2, s['p2_list']), (3, s['p3_list'])]:
                            for new_sec in plist:
                                if new_sec != sec:
                                    r1_counts = Counter(st['round1'] for st in students if st['round1'])
                                    if r1_counts.get(new_sec, 0) < self.MAX_PER_SECTOR:
                                        s['round1'] = new_sec
                                        s['priority_used_r1'] = p
                                        break
                            if s['round1'] != sec:
                                break
                        if s['round1'] == sec:
                            s['round1'] = None
                            s['priority_used_r1'] = None
                    for s in r2_stds:
                        for p, plist in [(1, s['p1_list']), (2, s['p2_list']), (3, s['p3_list'])]:
                            for new_sec in plist:
                                if new_sec != sec and new_sec != s['round1']:
                                    r2_counts = Counter(st['round2'] for st in students if st['round2'])
                                    if r2_counts.get(new_sec, 0) < self.MAX_PER_SECTOR:
                                        s['round2'] = new_sec
                                        s['priority_used_r2'] = p
                                        break
                            if s['round2'] == sec:
                                s['round2'] = None
                                s['priority_used_r2'] = None
        
        for sec in self.sectors:
            r1_stds = [s for s in students if s['round1'] == sec]
            r2_stds = [s for s in students if s['round2'] == sec]
            n1, n2 = len(r1_stds), len(r2_stds)
            if n1 + n2 < min_sector_size:
                continue
            if abs(n1 - n2) > self.MAX_DIFF:
                excess_round = 'round1' if n1 > n2 else 'round2'
                deficit_round = 'round2' if n1 > n2 else 'round1'
                excess_count = n1 if n1 > n2 else n2
                target = (n1 + n2) // 2
                candidates = [(s, s[f'priority_used_{excess_round}']) for s in students if s[excess_round] == sec]
                candidates.sort(key=lambda x: x[1], reverse=True)
                to_move = excess_count - target
                for s, _ in candidates[:to_move]:
                    for p, plist in [(1, s['p1_list']), (2, s['p2_list'])]:
                        for new_sec in plist:
                            if new_sec != sec and new_sec != s[deficit_round]:
                                counts = Counter(st[deficit_round] for st in students if st[deficit_round])
                                if counts.get(new_sec, 0) < self.MAX_PER_SECTOR:
                                    s[excess_round] = None
                                    s[f'priority_used_{excess_round}'] = None
                                    s[deficit_round] = new_sec
                                    s[f'priority_used_{deficit_round}'] = p
                                    break
                        if s[deficit_round] != sec:
                            break
        
        for round_key, counts in [('round1', Counter(s['round1'] for s in students if s['round1'])), 
                                  ('round2', Counter(s['round2'] for s in students if s['round2']))]:
            over_sectors = [sec for sec in counts if counts[sec] > self.MAX_PER_SECTOR]
            for sec in over_sectors:
                excess = counts[sec] - self.MAX_PER_SECTOR
                candidates = [(s, s[f'priority_used_{round_key}']) for s in students if s[round_key] == sec]
                candidates.sort(key=lambda x: x[1], reverse=True)
                for s, _ in candidates[:excess]:
                    for p, plist in [(1, s['p1_list']), (2, s['p2_list']), (3, s['p3_list'])]:
                        for new_sec in plist:
                            if (round_key == 'round2' and new_sec != s['round1'] or round_key == 'round1') and counts.get(new_sec, 0) < self.MAX_PER_SECTOR:
                                s[round_key] = new_sec
                                s[f'priority_used_{round_key}'] = p
                                counts[sec] -= 1
                                counts[new_sec] = counts.get(new_sec, 0) + 1
                                break
        
        return students
    
    def group_students_by_sector(self, students, max_group_size=16, min_group_size=10):
        """
        Group students within each sector for Round 1 and Round 2, balancing group sizes.
        
        Args:
            students (list): List of student dictionaries with 'round1' and 'round2' assignments.
            max_group_size (int): Maximum number of students per group (default 16).
            min_group_size (int): Minimum number of students per group to aim for (default 10).
        
        Returns:
            list: Updated student dictionaries with 'round1_group' and 'round2_group' assignments.
        """
        if min_group_size <= 0:
            raise ValueError("min_group_size must be greater than 0")
        
        for s in students:
            s['round1_group'] = None
            s['round2_group'] = None
        
        for sec in self.sectors:
            r1_stds = [s for s in students if s['round1'] == sec]
            r2_stds = [s for s in students if s['round2'] == sec]
            
            # Group Round 1 students
            if r1_stds:
                random.shuffle(r1_stds)
                num_students = len(r1_stds)
                # Calculate number of groups to balance sizes
                num_groups = max(1, min((num_students + max_group_size - 1) // max_group_size, 
                                       num_students // min_group_size + 1))
                base_size = num_students // num_groups
                extra = num_students % num_groups
                group_sizes = [base_size + 1 if i < extra else base_size for i in range(num_groups)]
                
                current_idx = 0
                for i, size in enumerate(group_sizes):
                    group_name = f"Group {i + 1}"
                    for s in r1_stds[current_idx:current_idx + size]:
                        s['round1_group'] = group_name
                    current_idx += size
            
            # Group Round 2 students
            if r2_stds:
                random.shuffle(r2_stds)
                num_students = len(r2_stds)
                num_groups = max(1, min((num_students + max_group_size - 1) // max_group_size, 
                                       num_students // min_group_size + 1))
                base_size = num_students // num_groups
                extra = num_students % num_groups
                group_sizes = [base_size + 1 if i < extra else base_size for i in range(num_groups)]
                
                current_idx = 0
                for i, size in enumerate(group_sizes):
                    group_name = f"Group {i + 1}"
                    for s in r2_stds[current_idx:current_idx + size]:
                        s['round2_group'] = group_name
                    current_idx += size
        
        return students

    def update_class_dfs(self, class_dfs, all_df, students):
        """
        Update class DataFrames with sector and group assignments.
        
        Args:
            class_dfs (dict): Dictionary of class DataFrames.
            all_df (pd.DataFrame): Combined All F5 F6 DataFrame.
            students (list): List of student dictionaries with assignments.
        
        Returns:
            tuple: (updated class_dfs, updated all_df)
        """
        student_map = {(s['Class'], s['Cls No']): s for s in students}
        for idx, row in all_df.iterrows():
            student_id = (row['Class'], row['Cls No'])
            if student_id in student_map:
                s = student_map[student_id]
                all_df.at[idx, 'Round 1 Sector'] = s.get('round1', '')
                all_df.at[idx, 'Round 1 Group'] = s.get('round1_group', '')
                all_df.at[idx, 'Round 2 Sector'] = s.get('round2', '')
                all_df.at[idx, 'Round 2 Group'] = s.get('round2_group', '')
        
        for cls in self.classes:
            df = class_dfs[cls]
            for idx, row in df.iterrows():
                student_id = (row['Class'], row['Cls No'])
                if student_id in student_map:
                    s = student_map[student_id]
                    df.at[idx, 'Round 1 Sector'] = s.get('round1', '')
                    df.at[idx, 'Round 1 Group'] = s.get('round1_group', '')
                    df.at[idx, 'Round 2 Sector'] = s.get('round2', '')
                    df.at[idx, 'Round 2 Group'] = s.get('round2_group', '')
        
        return class_dfs, all_df

    def write_to_excel(self, class_dfs, all_df, students, output_file=None):
        """
        Write class, sector, and All F5 F6 DataFrames to Excel with highlighting and an overview sheet.

        Args:
            class_dfs (dict): Dictionary of class DataFrames.
            all_df (pd.DataFrame): Combined All F5 F6 DataFrame.
            students (list): List of student dictionaries with assignments.
            output_file (str, optional): Path to output Excel file.
        """
        if output_file is None:
            output_file = f"output_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Write class sheets
            for cls in self.classes:
                df = class_dfs[cls]
                df.to_excel(writer, sheet_name=cls, index=False)
                worksheet = writer.sheets[cls]
                for idx, row in df.iterrows():
                    if row['1st Priority'] == '':
                        for col in range(1, len(df.columns) + 1):
                            worksheet.cell(row=idx + 2, column=col).fill = self.yellow_fill
            
            # Write sector sheets
            student_map = {(s['Class'], s['Cls No']): s for s in students}
            for sec in self.sectors:
                sheet_df = pd.DataFrame(columns=['Student', 'Group'])
                r1_stds = [s for s in students if s['round1'] == sec]
                r2_stds = [s for s in students if s['round2'] == sec]
                
                row_idx = 0
                if r1_stds:
                    sheet_df.loc[row_idx, 'Student'] = 'Round 1'
                    sheet_df.loc[row_idx, 'Group'] = ''
                    row_idx += 1
                    r1_groups = sorted(set(s['round1_group'] for s in r1_stds if s['round1_group']))
                    for group in r1_groups:
                        group_stds = [s for s in r1_stds if s['round1_group'] == group]
                        group_stds.sort(key=lambda s: (s['Class'], s['Cls No']))
                        for s in group_stds:
                            sheet_df.loc[row_idx, 'Student'] = f"{s['Class']} {s['Cls No']} {s['Eng Name']}"
                            sheet_df.loc[row_idx, 'Group'] = s['round1_group'] or ''
                            row_idx += 1
                
                row_idx = 31
                if r2_stds:
                    sheet_df.loc[row_idx, 'Student'] = 'Round 2'
                    sheet_df.loc[row_idx, 'Group'] = ''
                    row_idx += 1
                    r2_groups = sorted(set(s['round2_group'] for s in r2_stds if s['round2_group']))
                    for group in r2_groups:
                        group_stds = [s for s in r2_stds if s['round2_group'] == group]
                        group_stds.sort(key=lambda s: (s['Class'], s['Cls No']))
                        for s in group_stds:
                            sheet_df.loc[row_idx, 'Student'] = f"{s['Class']} {s['Cls No']} {s['Eng Name']}"
                            sheet_df.loc[row_idx, 'Group'] = s['round2_group'] or ''
                            row_idx += 1
                
                sheet_df.to_excel(writer, sheet_name=sec[:31], index=False)
            
            # Write All F5 F6 sheet
            assign_df = all_df.drop(columns=['p1_list', 'p2_list', 'p3_list'], errors='ignore')
            assign_df.to_excel(writer, sheet_name='All F5 F6', index=False)
            worksheet = writer.sheets['All F5 F6']
            for idx, row in assign_df.iterrows():
                student_id = (row['Class'], row['Cls No'])
                all_prefs = []
                if student_id in student_map:
                    s = student_map[student_id]
                    all_prefs = s['p1_list'] + s['p2_list'] + s['p3_list']
                
                if not row['Round 1 Sector'] or not row['Round 2 Sector'] or \
                   not row['Round 1 Group'] or not row['Round 2 Group']:
                    for col in range(1, len(assign_df.columns) + 1):
                        worksheet.cell(row=idx + 2, column=col).fill = self.yellow_fill
                elif (row['Round 1 Sector'] and row['Round 1 Sector'] not in all_prefs) or \
                     (row['Round 2 Sector'] and row['Round 2 Sector'] not in all_prefs):
                    for col in range(1, len(assign_df.columns) + 1):
                        worksheet.cell(row=idx + 2, column=col).fill = self.red_fill
                elif row['1st Priority'] == '' and row['2nd Priority'] == '' and row['3rd Priority'] == '':
                    for col in range(1, len(assign_df.columns) + 1):
                        worksheet.cell(row=idx + 2, column=col).fill = self.grey_fill
            
            # Write overview sheet
            overview_data = []
            for sec in self.sectors:
                r1_stds = [s for s in students if s['round1'] == sec]
                r2_stds = [s for s in students if s['round2'] == sec]
                r1_groups = Counter(s['round1_group'] for s in r1_stds if s['round1_group'])
                r2_groups = Counter(s['round2_group'] for s in r2_stds if s['round2_group'])
                
                for group, count in r1_groups.items():
                    overview_data.append({'Sector': sec, 'Round': 'Round 1', 'Group': group, 'Students': count})
                for group, count in r2_groups.items():
                    overview_data.append({'Sector': sec, 'Round': 'Round 2', 'Group': group, 'Students': count})
            
            overview_df = pd.DataFrame(overview_data)
            overview_df.to_excel(writer, sheet_name='Overview', index=False)

    def process(self, min_sector_size=0, output_file=None, max_group_size=16, min_group_size=10, 
                response_file=None, classlist_file=None):
        """
        Main method to process responses and assign sectors.

        Args:
            min_sector_size (int): Minimum number of students in a sector across both rounds (default 0).
            output_file (str, optional): Path to output Excel file.
            max_group_size (int): Maximum number of students per group (default 16).
            min_group_size (int): Minimum number of students per group to aim for (default 10).
            response_file (str, optional): Path to the response file. Defaults to self.response_file.
            classlist_file (str, optional): Path to the class list file. Defaults to self.classlist_file.
        """
        output_folder = "output"
        os.makedirs(output_folder, exist_ok=True)

        # Use the provided file paths or default to self attributes
        response_file_path = response_file if response_file else self.response_file
        classlist_file_path = classlist_file if classlist_file else self.classlist_file

        response_df = pd.read_excel(response_file_path, sheet_name="Form Responses 1")
        response_df.columns = [
            'Timestamp', 'Chinese Name', 'English Name', 'Class', 'Class No.', 'Column 12',
            '1st Priority', '2nd Priority', '3rd Priority', 'Particular Course',
            'Topics and Suggestions', 'Other Majors'
        ]
        
        class_dfs = {}
        for cls in self.classes:
            df = pd.read_excel(classlist_file_path, sheet_name=cls)
            df.drop(columns=[col for col in df.columns if col.startswith('AFT')], inplace=True, errors='ignore')
            df.rename(columns={col: 'Chi Name' for col in df.columns if col.startswith('FT')}, inplace=True)
            new_cols = [
                '1st Priority', '2nd Priority', '3rd Priority', 'Particular Course', 
                'Topics and Suggestions', 'Other Majors', 
                'Round 1 Sector', 'Round 1 Group', 'Round 2 Sector', 'Round 2 Group'
            ]
            for col in new_cols:
                df[col] = ''
            class_dfs[cls] = df
        
        class_dfs, all_df = self.merge_responses(response_df, class_dfs)
        all_df = self.extract_sectors(all_df)
        assign_df = all_df[all_df['p1_list'].map(len) > 0].copy()
        students = self.assign_sectors(assign_df, min_sector_size=min_sector_size)
        students = self.group_students_by_sector(students, max_group_size=max_group_size, min_group_size=min_group_size)
        class_dfs, all_df = self.update_class_dfs(class_dfs, all_df, students)
        
        if output_file is None:
            output_file = os.path.join(output_folder, f"output_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        self.write_to_excel(class_dfs, all_df, students, output_file=output_file)
