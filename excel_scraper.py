#!/usr/bin/env python3
"""
Excel to CSV Transformer
Extracts everything right and under the first 'APPLICABILITY OF SCORING CRITERIA' cell in the 'Part VIII Scoring Criteria' sheet for all files in 2023_application/.
"""

import pandas as pd
import openpyxl
import os
import glob

class ExcelApplicabilityExtractor:
    def __init__(self, file_path):
        self.file_path = file_path

    def extract_applicability_block(self, output_csv):
        try:
            df = pd.read_excel(self.file_path, sheet_name='Part VIII Scoring Criteria', header=None)
            start_row, start_col = None, None
            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    cell = str(df.iat[i, j])
                    if 'APPLICABILITY OF SCORING CRITERIA' in cell.upper():
                        start_row, start_col = i, j
                        break
                if start_row is not None:
                    break
            if start_row is None or start_col is None:
                print(f"Could not find 'APPLICABILITY OF SCORING CRITERIA' in {self.file_path}.")
                return False
            # Extract everything right and under the found cell
            sub_df = df.iloc[start_row:, start_col:]
            os.makedirs(os.path.dirname(output_csv), exist_ok=True)
            sub_df.to_csv(output_csv, index=False, header=False)
            print(f"Saved applicability block to: {output_csv}")
            return True
        except Exception as e:
            print(f"Error processing {self.file_path}: {e}")
            return False

if __name__ == "__main__":
    input_dir = "2023_application"
    output_dir = "excel_to_csv_output"
    os.makedirs(output_dir, exist_ok=True)
    excel_files = glob.glob(os.path.join(input_dir, '*.xlsx'))
    print(f"Found {len(excel_files)} Excel files.")
    for file_path in excel_files:
        filename = os.path.splitext(os.path.basename(file_path))[0]
        output_csv = os.path.join(output_dir, f"{filename}_Applicability_Block.csv")
        extractor = ExcelApplicabilityExtractor(file_path)
        extractor.extract_applicability_block(output_csv)