import pandas as pd
from collections import defaultdict, Counter
import random
import os
import logging
import math
from openpyxl.styles import PatternFill

# --- Configure Logging ---
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - [IN %(funcName)s] - %(message)s'
)

class StudentAssignmentProcessor:
    """
    Processes student responses to assign them to two rounds of sector talks.
    The number of groups per sector is determined by the number of available speakers,
    with a maximum size constraint per group.
    """
    
    def __init__(self, response_file, classlist_file, speaker_file):
        logging.info("Initializing StudentAssignmentProcessor...")
        self.response_file = response_file
        self.classlist_file = classlist_file
        self.speaker_file = speaker_file
        
        # --- Configuration ---
        self.SECTORS = [
            'Allied Health (PT, OT, RT)', 'Architecture', 'Arts', 'Banking & Finance',
            'Computer Science', 'Education', 'Engineering', 'Law', 'Medical',
            'Social Science & Social Work'
        ]
        self.CLASSES = ['5A', '5B', '5C', '5D', '6A', '6B', '6C', '6D']
        self.MAX_PER_SECTOR = 48
        self.MAX_GROUP_SIZE = 25

        # --- File Column/Sheet Names ---
        self.SPEAKER_SECTOR_COL = '請勾選你希望參加的範疇'
        self.RESPONSE_SHEET = 'Form Responses 1'
        self.RESPONSE_COLS = {
            'Timestamp': 'Timestamp', 
            'English Name': 'Eng Name', 
            'Class': 'Class', 
            'Class No.': 'Class No.', 
            'Please prioritize the majors you are interested in [1st priority]': '1st Priority', 
            'Please prioritize the majors you are interested in [2nd priority]': '2nd Priority', 
            'Please prioritize the majors you are interested in [3rd priority]': '3rd Priority'
        }
        
        self.speaker_counts = {}
        self.sector_dict = {s.lower(): s for s in self.SECTORS}
        self.yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        self.red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        
        logging.info("Initialization complete.")

    def _load_speaker_data(self, sheet_name=0):
        logging.info(f"Loading speaker data from '{self.speaker_file}'...")
        try:
            speaker_df = pd.read_excel(self.speaker_file, sheet_name=sheet_name)
            if self.SPEAKER_SECTOR_COL not in speaker_df.columns:
                raise KeyError(f"FATAL: Column '{self.SPEAKER_SECTOR_COL}' not found in speaker file.")
            
            speaker_df[self.SPEAKER_SECTOR_COL] = speaker_df[self.SPEAKER_SECTOR_COL].str.replace('⁠', '').str.strip()
            speaker_counts_series = speaker_df[self.SPEAKER_SECTOR_COL].value_counts()
            
            self.speaker_counts = {sector: int(speaker_counts_series.get(sector, 0)) for sector in self.SECTORS}
            self.speaker_names = defaultdict(list)
            for _, row in speaker_df.iterrows():
                sector = row[self.SPEAKER_SECTOR_COL]
                if sector in self.SECTORS:
                    self.speaker_names[sector].append(row.get('Speaker Name', f"Speaker {len(self.speaker_names[sector]) + 1}"))
            
            logging.info(f"Speaker counts loaded successfully: {self.speaker_counts}")
            logging.info(f"Speaker names loaded successfully: {dict(self.speaker_names)}")

        except FileNotFoundError:
            logging.error(f"FATAL: Speaker file not found at '{self.speaker_file}'")
            raise
        except Exception as e:
            logging.error(f"An error occurred while loading speaker data: {e}")
            raise

    # --- MODIFIED MERGE FUNCTION TO PREVENT DUPLICATE COLUMNS ---
    def _merge_responses_with_classlists(self, response_df, class_dfs):
        logging.info("Merging student responses with class lists...")
        
        response_df['Class'] = response_df['Class'].str.upper().str.strip()
        response_df['Class No.'] = pd.to_numeric(response_df['Class No.'], errors='coerce')
        latest_responses = response_df.sort_values(by='Timestamp', ascending=False).drop_duplicates(subset=['Class', 'Class No.'])
        
        # Drop 'Eng Name' and 'Timestamp' from responses
        latest_responses = latest_responses.drop(columns=['Eng Name', 'Timestamp'], errors='ignore')

        all_class_df = pd.concat(class_dfs.values(), ignore_index=True)
        all_class_df['Class'] = all_class_df['Class'].str.upper().str.strip()

        merged_df = pd.merge(
            left=all_class_df,
            right=latest_responses,
            how='left',
            left_on=['Class', 'Cls No'],
            right_on=['Class', 'Class No.']
        )
        
        # Drop 'Class No.' from merged DataFrame
        if 'Class No.' in merged_df.columns:
            merged_df.drop(columns=['Class No.'], inplace=True)
        
        # Drop columns starting with 'FT' or 'AFT'
        cols_to_drop = [col for col in merged_df.columns if col.startswith('FT') or col.startswith('AFT')]
        merged_df.drop(columns=cols_to_drop, inplace=True, errors='ignore')
        
        # Rename 'Eng Name_x' to 'Eng Name'
        if 'Eng Name_x' in merged_df.columns:
            merged_df.rename(columns={'Eng Name_x': 'Eng Name'}, inplace=True)
        
        for cls in self.CLASSES:
            class_dfs[cls] = merged_df[merged_df['Class'] == cls].copy()
            
        logging.info("Merge complete.")
        return class_dfs, merged_df

    def _extract_sectors(self, df):
        logging.info("Extracting sector preferences...")
        def get_sectors(entry):
            if pd.isnull(entry): return []
            entry = str(entry).replace('⁠', '')
            parts = [p.strip().lower() for p in entry.split(',')]
            matched = []
            for p in parts:
                if 'allied health' in p and '(pt' in p: 
                    matched.append(self.SECTORS[0])
                elif 'social science' in p: 
                    matched.append('Social Science & Social Work')
                elif p in self.sector_dict: 
                    matched.append(self.sector_dict[p])
            return matched

        df['p1_list'] = df['1st Priority'].apply(get_sectors)
        df['p2_list'] = df['2nd Priority'].apply(get_sectors)
        df['p3_list'] = df['3rd Priority'].apply(get_sectors)
        logging.info("Sector extraction complete.")
        return df

    def _assign_sectors(self, assign_df):
        logging.info("Assigning students to sectors...")
        students_df = assign_df.copy()
        students_df['student_id'] = list(zip(students_df['Class'], students_df['Cls No']))
        students = students_df.to_dict('records')

        for s in students:
            s.update({'round1': None, 'round2': None})
        
        sector_prefs = defaultdict(list)
        for s in students:
            for p, plist in [(1, s['p1_list']), (2, s['p2_list'])]:
                for sec in plist:
                    if self.speaker_counts.get(sec, 0) > 0:  # Skip sectors with zero speakers
                        sector_prefs[sec].append((s, p))
        
        for sec in self.SECTORS:
            if self.speaker_counts.get(sec, 0) == 0:
                continue  # Skip sectors with zero speakers
            
            candidates = sector_prefs.get(sec, [])
            if not candidates:
                continue
            
            random.shuffle(candidates)
            candidates.sort(key=lambda x: x[1])
            
            assigned_students = set()
            for s, p in candidates:
                if s['student_id'] in assigned_students:
                    continue
                
                r1_counts = Counter(st.get('round1') for st in students if st.get('round1'))
                r2_counts = Counter(st.get('round2') for st in students if st.get('round2'))

                if r1_counts.get(sec, 0) <= r2_counts.get(sec, 0):
                    if not s['round1'] and r1_counts.get(sec, 0) < self.MAX_PER_SECTOR:
                        s['round1'] = sec
                        assigned_students.add(s['student_id'])
                else:
                    if not s['round2'] and s.get('round1') != sec and r2_counts.get(sec, 0) < self.MAX_PER_SECTOR:
                        s['round2'] = sec
                        assigned_students.add(s['student_id'])

        for s in students:
            if not s['round1'] or not s['round2']:
                all_prefs = [sec for sec in (s['p1_list'] + s['p2_list'] + s['p3_list']) if self.speaker_counts.get(sec, 0) > 0]
                for sec in all_prefs:
                    if not s['round1']:
                        s['round1'] = sec
                    elif not s['round2'] and s['round1'] != sec:
                        s['round2'] = sec
                        break
        logging.info("Sector assignment complete.")
        return students

    def _group_students_by_speaker_availability(self, students):
        logging.info("Grouping students by sector based on speaker counts and group size...")
        for s in students:
            s['round1_group'] = None
            s['round2_group'] = None
        
        for sec in self.SECTORS:
            num_speakers = self.speaker_counts.get(sec, 0)
            if num_speakers == 0:
                continue  # Skip sectors with zero speakers

            # Round 1 grouping
            r1_stds = [s for s in students if s['round1'] == sec]
            if r1_stds:
                num_students_r1 = len(r1_stds)
                groups_needed = math.ceil(num_students_r1 / self.MAX_GROUP_SIZE)
                num_groups_to_form = min(groups_needed, num_speakers)
                
                random.shuffle(r1_stds)
                for i, s in enumerate(r1_stds):
                    group_number = (i % num_groups_to_form) + 1
                    s['round1_group'] = f"Group {group_number}"

            # Round 2 grouping
            r2_stds = [s for s in students if s['round2'] == sec]
            if r2_stds:
                num_students_r2 = len(r2_stds)
                groups_needed = math.ceil(num_students_r2 / self.MAX_GROUP_SIZE)
                num_groups_to_form = min(groups_needed, num_speakers)

                random.shuffle(r2_stds)
                for i, s in enumerate(r2_stds):
                    group_number = (i % num_groups_to_form) + 1
                    s['round2_group'] = f"Group {group_number}"

        # Highlight students who could not be assigned to any sector
        for s in students:
            if not s['round1'] or not s['round2']:
                s['highlight'] = 'red'
            else:
                s['highlight'] = None

        logging.info("Grouping complete.")
        return students

    def _update_dataframes_with_assignments(self, class_dfs, all_df, students):
        logging.info("Updating DataFrames with final assignments...")
        student_map = {(s['Class'], s['Cls No']): s for s in students}
        
        for col in ['Round 1 Sector', 'Round 1 Group', 'Round 2 Sector', 'Round 2 Group']:
            if col not in all_df.columns:
                all_df[col] = None

        all_df['key'] = list(zip(all_df['Class'], all_df['Cls No']))
        all_df['Round 1 Sector'] = all_df['key'].map(lambda x: student_map.get(x, {}).get('round1'))
        all_df['Round 1 Group'] = all_df['key'].map(lambda x: student_map.get(x, {}).get('round1_group'))
        all_df['Round 2 Sector'] = all_df['key'].map(lambda x: student_map.get(x, {}).get('round2'))
        all_df['Round 2 Group'] = all_df['key'].map(lambda x: student_map.get(x, {}).get('round2_group'))
        all_df.drop(columns=['key'], inplace=True)

        for cls in self.CLASSES:
            class_dfs[cls] = all_df[all_df['Class'] == cls].copy()

        logging.info("DataFrame updates complete.")
        return class_dfs, all_df

    # --- REVISED EXCEL WRITING FOR CLEAN OUTPUT ---
    def _write_to_excel(self, class_dfs, all_df, students, output_file):
        logging.info(f"Writing output to Excel file: {output_file}")
        
        # Define the exact columns for the final output
        final_output_columns = [
            'Class', 'Cls No', 'Reg No', 'Eng Name', 'Chi Name', 
            '1st Priority', '2nd Priority', '3rd Priority', 
            'Round 1 Sector', 'Round 1 Group', 
            'Round 2 Sector', 'Round 2 Group'
        ]

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pd.DataFrame(list(self.speaker_counts.items()), columns=['Sector', 'NumberOfSpeakers']).to_excel(
                writer, sheet_name='Speaker Allocation', index=False
            )
            
            # Write individual class sheets
            for cls in self.CLASSES:
                df_full = class_dfs[cls]
                cols_to_write = [col for col in final_output_columns if col in df_full.columns]
                df_to_write = df_full[cols_to_write]
                df_to_write.to_excel(writer, sheet_name=cls, index=False)
            
            # Write individual sector sheets
            for sec in self.SECTORS:
                r1_stds = sorted([s for s in students if s['round1'] == sec], key=lambda x: (x.get('round1_group', ''), x['Class'], x['Cls No']))
                r2_stds = sorted([s for s in students if s['round2'] == sec], key=lambda x: (x.get('round2_group', ''), x['Class'], x['Cls No']))
                
                sector_data = []
                if r1_stds:
                    sector_data.append({'Student': 'Round 1', 'Group': ''})
                    for s in r1_stds:
                        sector_data.append({'Student': f"{s['Class']} {s['Cls No']} {s.get('Eng Name', '')}", 'Group': s.get('round1_group')})
                
                if r2_stds:
                    sector_data.append({'Student': '', 'Group': ''})
                    sector_data.append({'Student': 'Round 2', 'Group': ''})
                    for s in r2_stds:
                        sector_data.append({'Student': f"{s['Class']} {s['Cls No']} {s.get('Eng Name', '')}", 'Group': s.get('round2_group')})
                
                pd.DataFrame(sector_data).to_excel(writer, sheet_name=sec[:31], index=False)
            
            # Write all students sheet
            cols_to_write_all = [col for col in final_output_columns if col in all_df.columns]
            all_df_final = all_df[cols_to_write_all]
            all_df_final.to_excel(writer, sheet_name='All F5 F6', index=False)
            worksheet = writer.sheets['All F5 F6']
            
            for idx, row in all_df.iterrows():
                if row.get('highlight') == 'red':
                    for col in range(1, len(all_df_final.columns) + 1):
                        worksheet.cell(row=idx + 2, column=col).fill = self.red_fill

            # Write overview sheet
            overview_data = []
            for sec in self.SECTORS:
                r1_groups_counter = Counter(s['round1_group'] for s in students if s['round1'] == sec and s['round1_group'])
                r2_groups_counter = Counter(s['round2_group'] for s in students if s['round2'] == sec and s['round2_group'])
                
                for group, count in sorted(r1_groups_counter.items()):
                    overview_data.append({'Sector': sec, 'Round': 'Round 1', 'Group': group, 'Students': count})
                for group, count in sorted(r2_groups_counter.items()):
                    overview_data.append({'Sector': sec, 'Round': 'Round 2', 'Group': group, 'Students': count})
            
            pd.DataFrame(overview_data).to_excel(writer, sheet_name='Overview', index=False)
        logging.info("Excel file writing complete.")

    def process(self, output_file=None, speaker_sheet_name=0):
        logging.info("Starting main processing workflow...")
        os.makedirs("output", exist_ok=True)
        
        self._load_speaker_data(sheet_name=speaker_sheet_name)
        
        response_df = pd.read_excel(self.response_file, sheet_name=self.RESPONSE_SHEET)
        response_df = response_df[list(self.RESPONSE_COLS.keys())].rename(columns=self.RESPONSE_COLS)
        
        # --- REVISED CLASSLIST LOADING TO CLEAN COLUMNS ---
        logging.info("Loading and preparing class list files...")
        class_dfs = {}
        for cls in self.CLASSES:
            df = pd.read_excel(self.classlist_file, sheet_name=cls)
            
            # Find the first 'FT:' column and rename it to 'Chi Name'
            ft_cols = [col for col in df.columns if str(col).startswith('FT:')]
            if ft_cols:
                df.rename(columns={ft_cols[0]: 'Chi Name'}, inplace=True)
            
            # Find all remaining 'FT:' and all 'AFT:' columns to drop
            cols_to_drop = [col for col in df.columns if str(col).startswith('AFT:')]
            ft_cols_remaining = [col for col in df.columns if str(col).startswith('FT:')]
            cols_to_drop.extend(ft_cols_remaining)
            
            df.drop(columns=cols_to_drop, inplace=True, errors='ignore')
            
            df['Class'] = cls
            class_dfs[cls] = df
        
        class_dfs, all_df = self._merge_responses_with_classlists(response_df, class_dfs)
        all_df = self._extract_sectors(all_df)
        
        assign_df = all_df[all_df['p1_list'].map(len) > 0].dropna(subset=['Cls No']).copy()
        
        students = self._assign_sectors(assign_df)
        students = self._group_students_by_speaker_availability(students)
        
        class_dfs, all_df = self._update_dataframes_with_assignments(class_dfs, all_df, students)
        
        if output_file is None:
            output_file = os.path.join("output", f"student_assignments_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        self._write_to_excel(class_dfs, all_df, students, output_file=output_file)
        logging.info(f"Processing complete. Output saved to {output_file}")

if __name__ == '__main__':
    try:
        processor = StudentAssignmentProcessor(
            response_file='input/response_file.xlsx',
            classlist_file='input/classlist_file.xlsx',
            speaker_file='input/Speaker_list.xlsx'
        )
        processor.process()
        
    except FileNotFoundError as e:
        logging.error(f"A required file was not found. Please check the file paths. Details: {e}")
    except KeyError as e:
        logging.error(f"A required column was not found in one of the files. Please check column names. Details: {e}")
    except Exception as e:
        logging.error(f"An unexpected error occurred during processing: {e}")