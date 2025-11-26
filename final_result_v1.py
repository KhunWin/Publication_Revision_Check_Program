
import pandas as pd
import numpy as np
import re
import sys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


class ResultGenerator:
    """Handles result file generation and summary statistics."""
    
    def __init__(self, client_df, output_file_path):
        self.client_df = client_df
        self.output_file_path = output_file_path
    
    def save_results(self):
        """Save results to Excel file."""
        try:
            output_path = Path(self.output_file_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            ##comment this to save excel

            # self.client_df.to_csv(output_path, index=False)
            print(f"Formatted client file saved to: {output_path}")
            
            self.client_df.to_excel(self.output_file_path, index=False, engine='openpyxl')
            # print(f"Results saved to: {self.output_file_path}")
            
        except Exception as e:
            print(f"Error saving results: {e}")
            sys.exit(1)
    
    def generate_summary(self):
        """Generate and print summary statistics."""
        total = len(self.client_df)
        verified = len(self.client_df[self.client_df['Result'] == 'Verified'])
        not_found = len(self.client_df[self.client_df['Result'].str.contains('Not found', na=False)])
        needs_check = total - verified - not_found
        
        print("\n" + "="*50)
        print("COMPARISON SUMMARY")
        print("="*50)
        print(f"Total documents processed: {total}")
        print(f"Verified (matching): {verified} ({verified/total*100:.1f}%)")
        print(f"Needs checking (different): {needs_check} ({needs_check/total*100:.1f}%)")
        print(f"Not found: {not_found} ({not_found/total*100:.1f}%)")
        print("="*50 + "\n")
