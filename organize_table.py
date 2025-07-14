import os
import pandas as pd
import glob
import re

def extract_company_name(csv_path):
    """Extract company name from the CSV filename"""
    filename = os.path.basename(csv_path)
    # Look for pattern like "2023-006helixcore" and extract just "helixcore"
    match = re.search(r'2023-\d{3}([^_]+)', filename)
    if match:
        return match.group(1)
    # Fallback to original pattern if the new one doesn't match
    match = re.search(r'(\d{4}\d{3}[^_]+)', filename)
    if match:
        company_name = match.group(1)
        # Remove the "2023-xxx" pattern from the company name
        company_name = re.sub(r'2023-\d{3}', '', company_name)
        return company_name
    return filename.replace('_Applicability_Block.csv', '')

def extract_scores_from_csv(csv_path):
    """Extract scores from the CSV file"""
    df = pd.read_csv(csv_path, header=None)
    company_name = extract_company_name(csv_path)
    results = []
    
    # Skip patterns for unwanted rows
    skip_patterns = [
        'APPLICABILITY OF SCORING CRITERIA', 'Maximum Point Totals', 'Type of Credit:',
        'Set-Aside Elections', 'NonProfit', 'General', 'Disaster Recovery', 'Preservation:',
        'HUD RAD', 'HTC', 'HUD RA', 'CHDO', 'NOAH', 'HUD RAD is also'
    ]
    
    for idx, row in df.iterrows():
        if idx < 2:  # Skip header rows
            continue
            
        criterion = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
        
        if not criterion or criterion == 'nan' or len(criterion) <= 3:
            continue
            
        # Skip unwanted rows
        if any(pattern in criterion for pattern in skip_patterns):
            continue
        
        # Remove roman numerals from criterion name
        criterion_clean = re.sub(r'^[IVX]+\.\s*', '', criterion)
        
        # Extract numeric values from specific columns
        max_score = None
        actual_score = None
        
        # Look for max_score in columns 5-11
        for col_idx in range(5, min(12, len(row))):
            val = row.iloc[col_idx]
            if pd.notna(val):
                try:
                    max_score = float(val)
                    break
                except (ValueError, TypeError):
                    continue
        
        # Look for actual_score in columns 23-36
        for col_idx in range(23, min(37, len(row))):
            val = row.iloc[col_idx]
            if pd.notna(val):
                try:
                    actual_score = float(val)
                    break
                except (ValueError, TypeError):
                    continue
        
        # Add valid scores
        if max_score is not None and actual_score is not None:
            results.append((criterion_clean, max_score, actual_score))
    
    return company_name, results

def clean_column_name(attr):
    """Clean attribute name for use as column name"""
    return (attr.replace(" ", "_")
               .replace("/", "_")
               .replace("\"", "")
               .replace("'", "")
               .replace("(", "")
               .replace(")", "")
               .replace("-", "_")
               .replace(".", "")
               .replace(",", ""))

def main():
    input_dir = "excel_to_csv_output"
    
    if not os.path.exists(input_dir):
        print(f"Directory '{input_dir}' does not exist.")
        return
    
    csv_files = glob.glob(os.path.join(input_dir, "*_Applicability_Block.csv"))
    
    if not csv_files:
        print(f"No CSV files found in '{input_dir}'")
        return
    
    print(f"Processing {len(csv_files)} CSV files...")
    
    all_rows = []
    for csv_file in csv_files:
        print(f"Processing: {os.path.basename(csv_file)}")
        
        try:
            company_name, scores = extract_scores_from_csv(csv_file)
            
            if not scores:
                print(f"  Warning: No scores extracted")
                continue
            
            row = {
                "company_name": company_name
            }
            
            for attr, max_score, actual_score in scores:
                col_base = clean_column_name(attr)
                row[f"{col_base}_max"] = max_score
                row[f"{col_base}_actual"] = actual_score
            
            all_rows.append(row)
            print(f"  Extracted {len(scores)} scoring criteria")
            
        except Exception as e:
            print(f"  Error: {str(e)}")
            continue
    
    if not all_rows:
        print("No data extracted.")
        return
    
    # Create and save summary DataFrame
    summary_df = pd.DataFrame(all_rows).fillna(0)
    output_file = "applicability_scores_summary_wide.csv"
    summary_df.to_csv(output_file, index=False)
    
    print(summary_df.head())

if __name__ == "__main__":
    main()