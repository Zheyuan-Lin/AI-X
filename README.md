# CSV Data Organizer

A Python script suite that downloads Excel files from Georgia DCA website, extracts scoring criteria data, and organizes them into a structured summary table.

## Requirements

```bash
pip install -r requirements.txt
```

## Input Files

The download script requires two text files:

**download_links.txt** - URLs to download (one per line):
```
https://dca.georgia.gov/sites/default/files/2023-001johngrhmcoreapp.xlsx
```

**filenames.txt** - Corresponding filenames (one per line):
```
2023-001johngrhmcoreapp.xlsx
```

### Step 1: Download Excel files
```bash
python download_files.py
```

### Step 2: Extract scoring criteria to csv file
```bash
python excel_scraper.py
```

### Step 3: Organize data into summary table
```bash
python organize_table.py
```

## Output

The final file `applicability_scores_summary_wide.csv` will contain:
- Company names (cleaned)
- Scoring criteria with maximum and actual scores
- One row per company 