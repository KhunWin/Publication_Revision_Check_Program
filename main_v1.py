import pandas as pd
import numpy as np
from datetime import datetime
import re
import sys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from pub_v1 import *
from compare_v2 import *
from final_result_v1 import *


class DocumentRevisionTool:
    """Main orchestrator class that coordinates all operations."""
    
    def __init__(self, client_file, home_file, output_file='result_one.xlsx'):
        self.client_file = client_file
        self.home_file = home_file
        self.output_file = output_file
        self.formatted_client_file = 'client_formatted.csv'
    
    def run(self):
        """Execute the complete comparison workflow."""
        print("="*50)
        print("Document Revision Comparison Tool")
        print("="*50 + "\n")
        
        # Step 1: Load data
        print("Step 1: Loading data files...")
        loader = DataLoader(self.client_file, self.home_file)
        client_df, home_df = loader.load_files()
        
        # Step 2: Format client file
        print("\nStep 2: Formatting client file...")
        formatter = ClientFormatter(client_df)
        # client_df = formatter.create_formatted_column()
        client_df = formatter.process()
        formatter.save_formatted_file(self.formatted_client_file)
        
        # Step 3: Process home file
        print("\nStep 3: Processing home file...")
        home_processor = HomeProcessor(home_df)
        home_df = home_processor.remove_duplicates()
        
        # Step 4: Compare documents
        print("\nStep 4: Comparing documents...")
        comparator = RevisionComparator(client_df, home_df)
        result_df = comparator.process_comparisons()
        
        # Step 5: Save results
        print("\nStep 5: Saving results...")
        result_gen = ResultGenerator(result_df, self.output_file)
        result_gen.save_results()
        
        # Step 6: Apply formatting
        print("\nStep 6: Applying cell colors...")
        excel_formatter = ExcelFormatter(self.output_file)
        excel_formatter.apply_colors()
        
        # Step 7: Generate summary
        print("\nStep 7: Generating summary...")
        result_gen.generate_summary()
        
        print("Process completed successfully!")


def main():
    """Main entry point."""
    if len(sys.argv) >= 3:
        client_file = sys.argv[1]
        home_file = sys.argv[2]
        output_file = sys.argv[3] if len(sys.argv) >= 4 else 'result_one.xlsx'
    else:
        # Use sample files for testing
        client_file = 'client_origin.csv'
        home_file = 'home_origin.csv'
        # output_file = 'finn_revision_testing.xlsx'
        output_file = 'finn_revision_3.xlsx'
        
        print(f"Using default files:")
        print(f"  Client: {client_file}")
        print(f"  Home: {home_file}")
        print(f"  Output: {output_file}\n")
    
    tool = DocumentRevisionTool(client_file, home_file, output_file)
    tool.run()


if __name__ == "__main__":
    main()