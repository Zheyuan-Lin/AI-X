#!/usr/bin/env python3
"""
Download Excel files from Georgia DCA website
"""

import requests
import os
from urllib.parse import urljoin

def download_files():
    # Base URL
    base_url = "https://dca.georgia.gov"
    
    # Read the download links
    with open('download_links.txt', 'r') as f:
        download_urls = [line.strip() for line in f if line.strip()]
    
    # Read the filenames
    with open('filenames.txt', 'r') as f:
        filenames = [line.strip() for line in f if line.strip()]
    
    # Create output directory
    output_dir = "2023_application"
    os.makedirs(output_dir, exist_ok=True)
    
    # Download each file
    for i, (url, filename) in enumerate(zip(download_urls, filenames)):
        if i >= len(filenames):  # Handle case where we have more URLs than filenames
            break
            
        output_path = os.path.join(output_dir, filename)
        
        print(f"Downloading {i+1}/{len(filenames)}: {filename}")
        
        try:
            response = requests.get(url, stream=True)
            response.raise_for_status()
            
            with open(output_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            print(f"  ✓ Downloaded: {filename}")
            
        except Exception as e:
            print(f"  ✗ Error downloading {filename}: {e}")
    
    print(f"\nDownload completed! Files saved to {output_dir}/")

if __name__ == "__main__":
    download_files() 