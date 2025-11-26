import pandas as pd
import numpy as np
from datetime import datetime
import re
import sys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

class RevisionComparator:
    """Handles the comparison logic between client and home files."""
    
    def __init__(self, client_df, home_df):
        self.client_df = client_df.copy()
        self.home_df = home_df.copy()
        
        # Initialize result columns
        if 'Result' not in self.client_df.columns:
            self.client_df['Result'] = ''
        if 'Doc Call Number' not in self.client_df.columns:
            self.client_df['Doc Call Number'] = ''
        if 'Note' not in self.client_df.columns:
            self.client_df['Note'] = ''
    
    @staticmethod
    def normalize_basic_revision(rev_str):
        """Convert BASIC or BAS to 0."""
        if pd.isna(rev_str):
            return ''
        
        rev_str = str(rev_str).strip().upper()
        
        if rev_str in ['BASIC', 'BAS']:
            return '0'
        
        return rev_str
    
    @staticmethod
    def normalize_tr_string(tr_str):
        """Normalize TR string by removing spaces and converting to uppercase.
        Examples: 'TR 002' -> 'TR002', 'TR002' -> 'TR002', 'tr 002' -> 'TR002'
        """
        if pd.isna(tr_str) or str(tr_str).strip() == '':
            return ''
        
        return str(tr_str).strip().replace(' ', '').upper()
    
    @staticmethod
    def parse_date(date_str):
        """Parse date strings in multiple formats."""
        if pd.isna(date_str) or str(date_str).strip() == '':
            return None
        
        date_str = str(date_str).strip()
        
        date_formats = [
            '%m/%d/%Y', '%d-%b-%y', '%d-%b-%Y', '%m/%d/%y',
            '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%d-%m-%y'
        ]
        
        for date_format in date_formats:
            try:
                parsed_date = datetime.strptime(date_str, date_format)
                if parsed_date.year < 100:
                    if parsed_date.year < 50:
                        parsed_date = parsed_date.replace(year=parsed_date.year + 2000)
                    else:
                        parsed_date = parsed_date.replace(year=parsed_date.year + 1900)
                return parsed_date
            except ValueError:
                continue
        
        return None
    
    @staticmethod
    def compare_dates(date1_str, date2_str):
        """Compare two date strings."""
        date1 = RevisionComparator.parse_date(date1_str)
        date2 = RevisionComparator.parse_date(date2_str)
        
        if date1 is None and date2 is None:
            return True
        if date1 is None or date2 is None:
            return False
        
        return date1.date() == date2.date()
    
    def find_by_document_number(self, doc_no):
        """Find matching rows in home file by Document Number."""
        if pd.isna(doc_no):
            return []
        
        doc_no_str = str(doc_no).strip()
        
        matching_rows = self.home_df[
            self.home_df['Document Number'].astype(str).str.strip() == doc_no_str
        ]

        print("This is in find_by_document_number\n")
        print("matching rows: ", matching_rows)
        
        return matching_rows.to_dict('records') if len(matching_rows) > 0 else []
    
    def find_by_revision_description(self, doc_no):
        """Find matching rows by checking Doc. No. in Revision Description."""
        if pd.isna(doc_no):
            return []
        
        doc_no_str = str(doc_no).strip()
        
        matching_rows = self.home_df[
            self.home_df['Revision Description'].astype(str).str.contains(
                doc_no_str, case=False, na=False, regex=False
            )
        ]

        print("This is in find_by_revision_description to find doc-no")
        print(doc_no_str," :: ", matching_rows)   

        return matching_rows.to_dict('records') if len(matching_rows) > 0 else []
    
    #this one is working so far
    def compare_revision_and_date(self, idx, row, matching_rows):
        """Compare Revision No./Date with Revision Num/Date from home file."""
        client_rev_no = row.get('Revision No.')
        client_rev_date = row.get('Rev. Date')
        
        # Normalize BASIC/BAS
        client_rev_normalized = self.normalize_basic_revision(client_rev_no)
        
        results = []
        call_numbers = []
        notes = []
        
        for match in matching_rows:
            home_rev_num = match.get('Revision Num')
            home_rev_date = match.get('Revision Date')
            call_number = match.get('Call Number', '')
            
            # Normalize home revision
            home_rev_normalized = self.normalize_basic_revision(home_rev_num)
            
            print(f"\n  Comparing revisions:")
            print(f"  Client: '{client_rev_normalized}' vs Home: '{home_rev_normalized}'")
            
            # Compare revision numbers
            rev_match = self.compare_revisions(client_rev_normalized, home_rev_normalized)
            
            # Compare dates
            if pd.isna(client_rev_date) or str(client_rev_date).strip() == '':
                # No client revision date
                notes.append('No Revision Date is given')
                date_match = False
            else:
                date_match = self.compare_dates(client_rev_date, home_rev_date)
            
            if rev_match and date_match:
                results.append('Verified')
                call_numbers.append(str(call_number))
            else:
                # Format mismatch result
                # Convert home_rev_normalized to int if it's a numeric string
                if not pd.isna(home_rev_normalized) and str(home_rev_normalized).strip() != '':
                    try:
                        # Try to convert to int
                        home_rev_display = str(int(float(home_rev_normalized)))
                    except (ValueError, TypeError):
                        # If conversion fails, keep as string
                        home_rev_display = str(home_rev_normalized)
                else:
                    home_rev_display = ''
                
                home_date_display = str(home_rev_date) if not pd.isna(home_rev_date) else ''
                results.append(f"{home_rev_display}/{home_date_display}")
                call_numbers.append(str(call_number))
        
        # Handle results
        if len(results) > 1:
            # Multiple matches - mark as duplicated
            self.client_df.at[idx, 'Note'] = 'duplicated'
            # self.client_df.at[idx, 'Doc Call Number'] = ', '.join(call_numbers)
            # self.client_df.at[idx, 'Result'] = results[0]  # Use first result
             # Filter out empty call numbers and join
            valid_call_numbers = [cn for cn in call_numbers if cn]
            self.client_df.at[idx, 'Doc Call Number'] = ', '.join(valid_call_numbers) if valid_call_numbers else ''
            self.client_df.at[idx, 'Result'] = results[0]
        elif len(results) == 1:
            self.client_df.at[idx, 'Doc Call Number'] = call_numbers[0]
            self.client_df.at[idx, 'Result'] = results[0]
            if notes:
                self.client_df.at[idx, 'Note'] = notes[0]
        
        return len(results) > 0

    @staticmethod
    def compare_revisions(rev1, rev2):

        print("1111111111This is compare_revisions")
        """Compare two revision strings, handling numeric and TR formats.
        Examples:
        - "2" == "02" -> True (numeric comparison)
        - "TR01" == "TR 1" -> True (TR with numeric comparison)
        - "TR 1" == "TR 01" -> True (TR with numeric comparison)
        - "ABC" == "ABC" -> True (string comparison)
        """
        if pd.isna(rev1) and pd.isna(rev2):
            return True
        if pd.isna(rev1) or pd.isna(rev2):
            return False
        
        rev1_str = str(rev1).strip()
        rev2_str = str(rev2).strip()
        
        # Check if both are TR revisions
        rev1_upper = rev1_str.upper()
        rev2_upper = rev2_str.upper()
        
        if rev1_upper.startswith('TR') and rev2_upper.startswith('TR'):
            print("11111111111this is if cell starts with TR")
            # Extract the numeric part after TR
            # Remove 'TR' prefix and any spaces
            rev1_num_part = re.sub(r'^TR\s*', '', rev1_upper, flags=re.IGNORECASE).strip()
            rev2_num_part = re.sub(r'^TR\s*', '', rev2_upper, flags=re.IGNORECASE).strip()
            
            try:
                # Try to compare numerically
                num1 = int(float(rev1_num_part))
                num2 = int(float(rev2_num_part))
                result = (num1 == num2)
                print(f"  TR numeric comparison: TR{num1} == TR{num2} -> {result}")
                return result
            except (ValueError, TypeError):
                # If numeric conversion fails, fall back to string comparison
                result = (rev1_num_part == rev2_num_part)
                print(f"  TR string comparison: TR{rev1_num_part} == TR{rev2_num_part} -> {result}")
                return result
        
        # Try numeric comparison for non-TR revisions
        try:
            client_num = int(float(rev1_str))
            home_num = int(float(rev2_str))
            result = (client_num == home_num)
            print(f"  Numeric comparison: {client_num} == {home_num} -> {result}")
            return result
        except (ValueError, TypeError):
            # Fall back to string comparison for non-numeric values
            result = (rev1_str == rev2_str)
            print(f"  String comparison: '{rev1_str}' == '{rev2_str}' -> {result}")
            return result
                      
    def process_comparisons(self):
        """Main processing logic for comparisons."""
        total_rows = len(self.client_df)
        
        for idx, row in self.client_df.iterrows():
            if (idx + 1) % 100 == 0:
                print(f"Processing row {idx + 1}/{total_rows}...")
            
            doc_no = row.get('Doc. No.')
            publi_type = row.get('Publi. Type')
            formatted = row.get('Formatted', '')
            revision_no = row.get('Revision No.')
            
            print(f"\n{'='*60}")
            print(f"Processing Client Row {idx + 1}:")
            print(f"  Doc. No.: {doc_no}")
            print(f"  Revision No.: {revision_no}")
            print(f"  Formatted: '{formatted}'")
            print(f"{'='*60}")
            
            # Step 1: Try to find by Document Number
            matching_rows = self.find_by_document_number(doc_no)
            
            if matching_rows:
                # Check if Formatted column has a value
                if not pd.isna(formatted) and str(formatted).strip() != '':
                    # Use formatted comparison (handles TR and other values)
                    if self.compare_with_formatted(idx, row, matching_rows):
                        continue
                
                # Use standard revision/date comparison
                if self.compare_revision_and_date(idx, row, matching_rows):
                    continue
            
            # # Step 2: Try Title matching (with TR logic if applicable)
            # formatted_str = str(formatted).strip().upper() if not pd.isna(formatted) else ''
            
            # if 'TR' in formatted_str:
            #     # Use TR-aware title matching
            #     matching_rows = self.find_by_title_keywords(doc_no, revision_no, formatted_str)
            # else:
            #     # Regular title matching (Formatted is empty or other value)
            #     matching_rows = self.find_by_title_keywords(doc_no)
            
            # Step 2: Try Title matching
            matching_rows = self.find_by_title_keywords(doc_no)

            if matching_rows:

                # Set Doc Call Number first when it is found in title
                # call_numbers = [str(match.get('Call Number', '')) for match in matching_rows if match.get('Call Number')]
                # self.client_df.at[idx, 'Doc Call Number'] = ', '.join(call_numbers) if call_numbers else ''
                
                # Check if Formatted has a value
                if not pd.isna(formatted) and str(formatted).strip() != '':
                    if self.compare_with_formatted(idx, row, matching_rows):
                        continue
                
                # Standard comparison
                if self.compare_revision_and_date(idx, row, matching_rows):
                    continue
            
            # Step 3: Try Revision Description matching (only if Formatted is empty)
            if pd.isna(formatted) or str(formatted).strip() == '':
                matching_rows = self.find_by_revision_description(doc_no)
                
                if matching_rows:
                    if self.compare_revision_and_date(idx, row, matching_rows):
                        continue
            
            # No match found
            self.client_df.at[idx, 'Result'] = 'Not found'
        
        print("\n" + "="*50)
        print("Comparison completed successfully!")
        print("="*50)
        return self.client_df

    ##this is to add if doc. no. is found in either title or revision description
    def find_by_title_keywords(self, doc_no, revision_no=None, formatted=None):
        """Find matching rows by checking if Doc. No. appears in Title.
        Does NOT filter by TR - just finds matches by Doc. No. in Title."""
        if pd.isna(doc_no):
            return []
        
        doc_no_str = str(doc_no).strip()
        
        print(f"\n=== DEBUG: find_by_title_keywords ===")
        print(f"Searching for Doc. No.: '{doc_no_str}'")
        
        matching_rows = []
        
        for idx, row in self.home_df.iterrows():
            title = str(row.get('Title', ''))
            
            # Check if Doc. No. appears in Title
            doc_no_found = False
            
            # Strategy 1: Check if doc_no appears as a complete substring in title
            if doc_no_str in title:
                print(f"  ✓ Doc No MATCH (substring): '{doc_no_str}' found in '{title}'")
                doc_no_found = True
            # Strategy 2: For complex doc numbers with spaces/letters
            elif ' ' in doc_no_str or re.search(r'[A-Za-z]', doc_no_str):
                words = re.findall(r'[A-Za-z0-9]+', doc_no_str)
                if len(words) > 1:
                    all_found = all(word.upper() in title.upper() for word in words)
                    if all_found:
                        print(f"  ✓ Doc No MATCH (keywords): All words {words} found in '{title}'")
                        doc_no_found = True
            # print("doc_no_found:", doc_no_found)
            if doc_no_found:
                matching_rows.append(row.to_dict())
        
        print(f"Total matches found: {len(matching_rows)}\n")
        return matching_rows 

    ##this is to add if doc. no. is found in either title or revision description
    # def compare_with_formatted(self, idx, row, matching_rows):
    #     """Compare when Formatted column has a value (TR or other)."""
    #     formatted = row.get('Formatted', '')
        
    #     if pd.isna(formatted) or str(formatted).strip() == '':
    #         return False
        
    #     formatted_str = str(formatted).strip().upper()
    #     doc_no = str(row.get('Doc. No.', '')).strip()
        
    #     print(f"\n  Comparing in compare_with_formatted:")
    #     print(f"  Doc. No.: '{doc_no}'")
    #     print(f"  Formatted: '{formatted_str}'")
    #     print(f"  Matching rows count: {len(matching_rows)}")
        
    #     # matching_rows already contains rows where Doc. No. is in Title
    #     # No need to search again
        
    #     if not matching_rows:
    #         self.client_df.at[idx, 'Result'] = 'Not found'
    #         self.client_df.at[idx, 'Note'] = f'Doc. No. {doc_no} not found in Title'
    #         return True
        
    #     # Doc. No. found - set the Doc Call Number
    #     call_numbers = [str(match.get('Call Number', '')) for match in matching_rows if match.get('Call Number')]
    #     self.client_df.at[idx, 'Doc Call Number'] = ', '.join(call_numbers) if call_numbers else ''
        
    #     # Check if Formatted contains 'TR'
    #     if 'TR' in formatted_str:
    #         # For TR: Compare Revision No. with Revision Num AND Rev. Date with Revision Date
    #         rev_no = row.get('Revision No.')
    #         client_rev_date = row.get('Rev. Date')
            
    #         print(f"  Checking TR comparison:")
    #         print(f"  Client Revision No.: '{rev_no}'")
    #         print(f"  Client Rev. Date: '{client_rev_date}'")
            
    #         verified = []
    #         mismatches = []
            
    #         for match in matching_rows:
    #             home_rev_num = str(match.get('Revision Num', '')).strip()
    #             home_rev_date = match.get('Revision Date')
                
    #             print(f"  Home Revision Num: '{home_rev_num}'")
    #             print(f"  Home Revision Date: '{home_rev_date}'")
                
    #             # Check if Revision No. matches Revision Num
    #             rev_no_str = str(rev_no).strip() if not pd.isna(rev_no) else ''
    #             rev_num_match = (rev_no_str.upper() == home_rev_num.upper()) if rev_no_str and home_rev_num else False
                
    #             # Check if dates match
    #             date_match = self.compare_dates(client_rev_date, home_rev_date)
                
    #             print(f"  Revision Num match: {rev_num_match}")
    #             print(f"  Date match: {date_match}")
                
    #             if rev_num_match and date_match:
    #                 verified.append(match)
    #             else:
    #                 mismatches.append({
    #                     'match': match,
    #                     'home_rev_num': home_rev_num,
    #                     'home_date': str(home_rev_date) if not pd.isna(home_rev_date) else ''
    #                 })
            
    #         if verified:
    #             self.client_df.at[idx, 'Result'] = 'Verified'
    #             if len(verified) > 1:
    #                 self.client_df.at[idx, 'Note'] = 'duplicated'
    #         elif mismatches:
    #             # TR found but revision/date mismatch
    #             mismatch = mismatches[0]
    #             mismatch_details = []
                
    #             if mismatch['home_rev_num']:
    #                 mismatch_details.append(f"Rev. Num: {mismatch['home_rev_num']}")
    #             if mismatch['home_date']:
    #                 mismatch_details.append(f"Rev. Date: {mismatch['home_date']}")
                
    #             self.client_df.at[idx, 'Result'] = '/ '.join(mismatch_details) if mismatch_details else 'Mismatch'
                
    #             if len(mismatches) > 1:
    #                 current_note = self.client_df.at[idx, 'Note']
    #                 if pd.isna(current_note) or str(current_note).strip() == '':
    #                     self.client_df.at[idx, 'Note'] = 'duplicated'
    #         else:
    #             # Doc. No. found but TR not found
    #             self.client_df.at[idx, 'Result'] = 'TR not found in Revision Num'
            
    #         return True
        
    #     else:
    #         # Formatted has a specific value (like STATEMENT number)
    #         # Compare this value with Revision Description
    #         found = False
    #         for match in matching_rows:
    #             rev_desc = str(match.get('Revision Description', ''))
                
    #             if formatted_str in rev_desc.upper():
    #                 found = True
    #                 self.client_df.at[idx, 'Result'] = 'Verified'
    #                 break
            
    #         if not found:
    #             self.client_df.at[idx, 'Result'] = f'{formatted_str} not found in Revision Description'
            
    #         return True

    def compare_with_formatted(self, idx, row, matching_rows):
        """Compare when Formatted column has a value (TR or other)."""
        formatted = row.get('Formatted', '')
        
        if pd.isna(formatted) or str(formatted).strip() == '':
            return False
        
        formatted_str = str(formatted).strip().upper()
        doc_no = str(row.get('Doc. No.', '')).strip()
        
        print(f"\n  Comparing in compare_with_formatted:")
        print(f"  Doc. No.: '{doc_no}'")
        print(f"  Formatted: '{formatted_str}'")
        print(f"  Matching rows count: {len(matching_rows)}")
        
        if not matching_rows:
            self.client_df.at[idx, 'Result'] = 'Not found'
            self.client_df.at[idx, 'Note'] = f'Doc. No. {doc_no} not found in Title'
            return True
        
        # Doc. No. found - set the Doc Call Number
        call_numbers = [str(match.get('Call Number', '')) for match in matching_rows if match.get('Call Number')]
        self.client_df.at[idx, 'Doc Call Number'] = ', '.join(call_numbers) if call_numbers else ''
        
        # Check if Formatted contains 'TR'
        if 'TR' in formatted_str:
            # For TR: Compare Revision No. with Revision Description (not Revision Num!)
            # and Rev. Date with Revision Date
            rev_no = row.get('Revision No.')
            client_rev_date = row.get('Rev. Date')
            
            print(f"  Checking TR comparison:")
            print(f"  Client Revision No.: '{rev_no}'")
            print(f"  Client Rev. Date: '{client_rev_date}'")
            
            # Normalize the client revision number for TR comparison
            client_rev_normalized = self.normalize_tr_string(rev_no)
            
            verified = []
            mismatches = []
            
            for match in matching_rows:
                home_rev_desc = str(match.get('Revision Description', '')).strip()
                home_rev_num = str(match.get('Revision Num', '')).strip()
                home_rev_date = match.get('Revision Date')
                
                print(f"  Home Revision Description: '{home_rev_desc}'")
                print(f"  Home Revision Num: '{home_rev_num}'")
                print(f"  Home Revision Date: '{home_rev_date}'")
                
                # Normalize the home revision description for comparison
                home_rev_desc_normalized = self.normalize_tr_string(home_rev_desc)
                
                # Check if client Revision No. matches the TR in Revision Description
                rev_match = False
                if home_rev_desc_normalized and client_rev_normalized in home_rev_desc_normalized:
                    rev_match = True
                    print(f"  ✓ TR Match: '{client_rev_normalized}' found in '{home_rev_desc_normalized}'")
                else:
                    print(f"  ✗ TR No Match: '{client_rev_normalized}' NOT in '{home_rev_desc_normalized}'")
                
                # Check if dates match
                date_match = self.compare_dates(client_rev_date, home_rev_date)
                
                print(f"  Revision match: {rev_match}")
                print(f"  Date match: {date_match}")
                
                if rev_match and date_match:
                    verified.append(match)
                else:
                    # Store mismatch information
                    mismatch_details = []
                    if home_rev_num:
                        mismatch_details.append(f"Rev. Num: {home_rev_num}")
                    if home_rev_date and not pd.isna(home_rev_date):
                        mismatch_details.append(f"Rev. Date: {str(home_rev_date)}")
                    
                    mismatches.append({
                        'match': match,
                        'details': '/ '.join(mismatch_details) if mismatch_details else 'Mismatch'
                    })
            
            if verified:
                self.client_df.at[idx, 'Result'] = 'Verified'
                if len(verified) > 1:
                    self.client_df.at[idx, 'Note'] = 'duplicated'
            elif mismatches:
                # TR found in title but revision/date mismatch
                self.client_df.at[idx, 'Result'] = mismatches[0]['details']
                
                if len(mismatches) > 1:
                    current_note = self.client_df.at[idx, 'Note']
                    if pd.isna(current_note) or str(current_note).strip() == '':
                        self.client_df.at[idx, 'Note'] = 'duplicated'
            else:
                # Doc. No. found but no TR matches at all
                self.client_df.at[idx, 'Result'] = 'TR not found'
            
            return True
        
        else:
            # Formatted has a specific value (like STATEMENT number)
            # Compare this value with Revision Description
            found = False
            for match in matching_rows:
                rev_desc = str(match.get('Revision Description', ''))
                
                if formatted_str in rev_desc.upper():
                    found = True
                    self.client_df.at[idx, 'Result'] = 'Verified'
                    break
            
            if not found:
                self.client_df.at[idx, 'Result'] = f'{formatted_str} not found in Revision Description'
            
            return True

