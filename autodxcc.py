## AUTODXCC v2.0 - 24/07/2025 ##
# Daniel R. Martins (danielrmartins00@gmail.com)

import pandas as pd
import requests
import fitz
import os
import glob
import re
from datetime import datetime
import shutil

def download_standings(date, folder):
    """Downloads latest DXCC Standings PDFs"""
    bands = ["MIXED", "PHONE", "CW", "RTTY", "SATELLITE", "HR", "CHAL", "160M", 
            "80M", "40M", "30M", "20M", "17M", "15M", "12M", "10M", "6M", "2M", 
            "70CM", "23CM"]
    expected_files = {f"DXCC-{band}-{date}-A4.pdf" for band in bands}

    if os.path.exists(folder):
        existing_files = set(os.listdir(folder))
        if expected_files.issubset(existing_files):
            print("Your standings are already up to date. Skipping download...")
            return

    if os.path.exists(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)

    for band in bands:
        filename = f"DXCC-{band}-{date}-A4.pdf"
        url = f"https://www.arrl.org/system/dxcc/view/{filename}"
        filepath = os.path.join(folder, filename)

        try:
            response = requests.get(url)
            response.raise_for_status()
            with open(filepath, 'wb') as f:
                f.write(response.content)
            print(f"Downloaded: {filename}")
        except Exception as e:
            print(f"Failed to download: {url} – {e}")

def extract_lines_from_pdf(pdf_path):
    """Extract all lines from a PDF into a single array."""
    all_lines = []
    with fitz.open(pdf_path) as pdf:
        for page_num in range(len(pdf)):
            text = pdf[page_num].get_text()
            lines = [line.strip() for line in text.splitlines()]  # remove spaces
            all_lines.extend(lines)
    return all_lines

def countries_before_callsign(lines, callsign):
    """Find the number of countries that appears before a given callsign, in a list of lines."""
    countries = None
    for line in lines:
        if line.isdigit() and len(line) >= 3:  # number always >= 100
            countries = int(line)
        elif callsign in line:
            return countries
    return None

def countries_before_callsign_hr(lines, callsign):
    """Find the number of countries that appears before a given callsign, in a list of lines, 
       for each mode of Honor Roll (Mixed, Phone, CW, Digital)."""
    callsign_pattern = r'\b' + re.escape(callsign) + r'\/'

    modes = ["Mixed", "Phone", "CW", "Digital"]
    result = {mode: None for mode in modes}
    
    current_mode = None
    countries = None
    for line in lines:
        if line.strip() in modes:
            current_mode = line.strip()
            continue
        
        if line.isdigit() and len(line) >= 3:
            countries = int(line)
        
        if current_mode and re.search(callsign_pattern, line):
            match = re.search(r'(\d+)$', line)
            if match:
                result[current_mode] = countries
            current_mode = None
            
    return result

def extract_modality(pdf_name):
    """Extract modality from pdf."""
    parts = pdf_name.split('-')
    if len(parts) >= 3:
        return parts[1]
    return pdf_name

def process_pdfs_and_callsigns(pdf_folder, callsigns, output_file):
    """Process PDFs and extract data for each callsign."""
    pdfs = glob.glob(os.path.join(pdf_folder, "*.pdf"))
    
    results = pd.DataFrame({'CALL': callsigns})

    for pdf_path in pdfs:
        pdf_name = os.path.basename(pdf_path)
        column_name = extract_modality(pdf_name)
        print(f"Processing: {pdf_name}")
        
        lines = extract_lines_from_pdf(pdf_path)
        if column_name == "HR":
            hr_data = [countries_before_callsign_hr(lines, call) for call in callsigns]
            
            results[f"HR MIX"] = [hr_data[i].get('Mixed', None) for i in range(len(callsigns))]
            results[f"HR PH"] = [hr_data[i].get('Phone', None) for i in range(len(callsigns))]
            results[f"HR DIG"] = [hr_data[i].get('Digital', None) for i in range(len(callsigns))]
            results[f"HR CW"] = [hr_data[i].get('CW', None) for i in range(len(callsigns))]
        else:
            results[column_name] = [countries_before_callsign(lines, call) for call in callsigns]

    column_order = ["CALL","HR MIX", "HR PH", "HR CW", "HR DIG", "MIXED", "PHONE", "CW", "RTTY", "SATELLITE",
                    "CHAL", "160M", "80M", "40M", "30M", "20M", "17M", "15M", "12M", "10M", "6M", "2M", 
                    "70CM", "23CM"]
    results = results[column_order]

    results.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

def main():
    timestamp = datetime.now().strftime("%Y%m%d")
    script_directory = os.path.dirname(os.path.abspath(__file__))
    output_file_path = os.path.join(script_directory, f"{timestamp}-autodxcc-report.xlsx")

    excel_file_name = input("Enter the callsigns file name (Excel file): ").strip() # with or without .xlsx
    if not excel_file_name.lower().endswith(".xlsx"):
        excel_file_name += ".xlsx"
    excel_file_path = os.path.join(script_directory, excel_file_name)    
    if not os.path.isfile(excel_file_path):
        print(f"Error: The file {excel_file_name} does not exist in the current directory.")
        return
    
    callsigns_df = pd.read_excel(excel_file_path)
    callsigns = callsigns_df['CALL'].tolist()

    pdf_folder = os.path.join(script_directory, 'pdfs')
    if not os.path.isdir(pdf_folder):
        print(f"Error: The 'pdfs' folder does not exist in {script_directory}.")
        return

    download_standings(timestamp, pdf_folder)
    process_pdfs_and_callsigns(pdf_folder, callsigns, output_file_path)

if __name__ == "__main__":
    main()
