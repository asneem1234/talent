"""
BRSR Simple Data Extractor - Focused on 3 Core Sections Only
Extracts:
1. Section A, IV.18.a - Employees by gender
2. Section C, P3, Q21 - Board/KMP composition  
3. Section C, P3, Q22 - Turnover rates
"""

import PyPDF2
import re
import csv
import os
from pathlib import Path


def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file"""
    try:
        print(f"\n{'='*60}")
        print(f"Processing: {Path(pdf_path).name}")
        print(f"{'='*60}")
        
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            total_pages = len(pdf_reader.pages)
            print(f"Total pages: {total_pages}")
            
            text = ""
            print("Extracting text...")
            for i, page in enumerate(pdf_reader.pages, 1):
                text += page.extract_text() + "\n"
                if i % 50 == 0:
                    print(f"  Processed {i}/{total_pages} pages...")
            
            print("✓ Text extracted successfully!")
            return text, Path(pdf_path).stem
    except Exception as e:
        print(f"✗ Error: {e}")
        return None, None


def extract_employees_data(text, data):
    """Extract Section A, IV.18.a - Employees by gender"""
    
    # Pattern for Permanent Employees (D)
    patterns_perm = [
        r'1[.\s]+Permanent\s*\([DdEe]\)\s*([\d,]+)\s+([\d,]+)\s+([\d.]+)%?\s+([\d,]+)\s+([\d.]+)%?',
        r'Permanent\s*\([DdEe]\)\s*([\d,]+)\s+([\d,]+)\s+([\d.]+)%?\s+([\d,]+)\s+([\d.]+)%?',
    ]
    
    # Pattern for Other than Permanent (E)
    patterns_other = [
        r'2[.\s]+Other\s+than\s+[Pp]ermanent\s*\([EeFf]\)\s*([\d,]+)\s+([\d,]+)\s+([\d.]+)%?\s+([\d,]+)\s+([\d.]+)%?',
        r'Other\s+than\s+[Pp]ermanent\s*\([EeFf]\)\s*([\d,]+)\s+([\d,]+)\s+([\d.]+)%?\s+([\d,]+)\s+([\d.]+)%?',
    ]
    
    # Pattern for Total Employees (D + E)
    patterns_total = [
        r'3[.\s]+T\s*otal\s+employees\s*\([DdEe]\s*\+\s*[EeFf]\)\s*([\d,]+)\s+([\d,]+)\s+([\d.]+)%?\s+([\d,]+)\s+([\d.]+)%?',
        r'Total\s+employees\s*\([DdEe]\s*\+\s*[EeFf]\)\s*([\d,]+)\s+([\d,]+)\s+([\d.]+)%?\s+([\d,]+)\s+([\d.]+)%?',
    ]
    
    # Extract Permanent Employees
    for pattern in patterns_perm:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['Permanent Employees Male Number'] = int(match.group(2).replace(',', ''))
            data['Permanent Employees Male %'] = float(match.group(3))
            data['Permanent Employees Female Number'] = int(match.group(4).replace(',', ''))
            data['Permanent Employees Female %'] = float(match.group(5))
            print(f"  ✓ Permanent Employees: {match.group(1).replace(',', '')}")
            break
    
    # Extract Other than Permanent
    for pattern in patterns_other:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['Other than Permanent Employees Male Number'] = int(match.group(2).replace(',', ''))
            data['Other than Permanent Employees Male %'] = float(match.group(3))
            data['Other than Permanent Employees Female Number'] = int(match.group(4).replace(',', ''))
            data['Other than Permanent Employees Female %'] = float(match.group(5))
            print(f"  ✓ Other than Permanent: {match.group(1).replace(',', '')}")
            break
    
    # Extract Total Employees
    for pattern in patterns_total:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['Total Employees Male Number'] = int(match.group(2).replace(',', ''))
            data['Total Employees Male %'] = float(match.group(3))
            data['Total Employees Female Number'] = int(match.group(4).replace(',', ''))
            data['Total Employees Female %'] = float(match.group(5))
            print(f"  ✓ Total Employees: {match.group(1).replace(',', '')}")
            break


def extract_board_kmp_data(text, data):
    """Extract Section C, P3, Q21 - Board and KMP composition"""
    
    # Board of Directors patterns
    board_patterns = [
        r'Board\s+of\s+Directors[^\n]*\s+([\d,]+)\s+([\d,]+)\s+([\d.]+)%?',
        r'BoD[^\n]*\s+([\d,]+)\s+([\d,]+)\s+([\d.]+)%?',
    ]
    
    # KMP patterns
    kmp_patterns = [
        r'Key\s+Management\s+Personnel[^\n]*\s+([\d,]+)\s+([\d,]+)\s+([\d.]+)%?',
        r'KMP[^\n]*\s+([\d,]+)\s+([\d,]+)\s+([\d.]+)%?',
    ]
    
    # Extract Board data
    for pattern in board_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['Board of Directors Total'] = int(match.group(1).replace(',', ''))
            data['Board of Directors Female Number'] = int(match.group(2).replace(',', ''))
            data['Board of Directors Female %'] = float(match.group(3))
            print(f"  ✓ Board: {match.group(1)} members")
            break
    
    # Extract KMP data
    for pattern in kmp_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['Key Management Personnel Total'] = int(match.group(1).replace(',', ''))
            data['Key Management Personnel Female Number'] = int(match.group(2).replace(',', ''))
            data['Key Management Personnel Female %'] = float(match.group(3))
            print(f"  ✓ KMP: {match.group(1)} members")
            break


def extract_turnover_data(text, data):
    """Extract Section C, P3, Q22 - Turnover rates for 3 years"""
    
    # Pattern for turnover table - looking for Permanent Employees line
    pattern = r'Permanent\s+Employees\s+([\d.]+)%?\s+([\d.]+)%?\s+([\d.]+)%?\s+([\d.]+)%?\s+([\d.]+)%?\s+([\d.]+)%?\s+([\d.]+)%?\s+([\d.]+)%?\s+([\d.]+)%?'
    
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        # Current Year (FY 2025)
        data['Turnover Current Year Permanent Employees Male %'] = float(match.group(1))
        data['Turnover Current Year Permanent Employees Female %'] = float(match.group(2))
        data['Turnover Current Year Permanent Employees Total %'] = float(match.group(3))
        
        # Previous Year (FY 2024)
        data['Turnover Previous Year Permanent Employees Male %'] = float(match.group(4))
        data['Turnover Previous Year Permanent Employees Female %'] = float(match.group(5))
        data['Turnover Previous Year Permanent Employees Total %'] = float(match.group(6))
        
        # Prior Year (FY 2023)
        data['Turnover Prior Year Permanent Employees Male %'] = float(match.group(7))
        data['Turnover Prior Year Permanent Employees Female %'] = float(match.group(8))
        data['Turnover Prior Year Permanent Employees Total %'] = float(match.group(9))
        
        print(f"  ✓ Turnover data extracted for 3 years")


def extract_brsr_data(text, company_name):
    """Extract only the 3 core BRSR sections"""
    
    data = {
        'Company': company_name,
        
        # Section A, IV.18.a - Employees by gender
        'Permanent Employees Male Number': 0,
        'Permanent Employees Male %': 0,
        'Permanent Employees Female Number': 0,
        'Permanent Employees Female %': 0,
        'Other than Permanent Employees Male Number': 0,
        'Other than Permanent Employees Male %': 0,
        'Other than Permanent Employees Female Number': 0,
        'Other than Permanent Employees Female %': 0,
        'Total Employees Male Number': 0,
        'Total Employees Male %': 0,
        'Total Employees Female Number': 0,
        'Total Employees Female %': 0,
        
        # Section C, P3, Q21 - Board and KMP
        'Board of Directors Total': 0,
        'Board of Directors Female Number': 0,
        'Board of Directors Female %': 0,
        'Key Management Personnel Total': 0,
        'Key Management Personnel Female Number': 0,
        'Key Management Personnel Female %': 0,
        
        # Section C, P3, Q22 - Turnover rates
        'Turnover Current Year Permanent Employees Male %': 0,
        'Turnover Current Year Permanent Employees Female %': 0,
        'Turnover Current Year Permanent Employees Total %': 0,
        'Turnover Previous Year Permanent Employees Male %': 0,
        'Turnover Previous Year Permanent Employees Female %': 0,
        'Turnover Previous Year Permanent Employees Total %': 0,
        'Turnover Prior Year Permanent Employees Male %': 0,
        'Turnover Prior Year Permanent Employees Female %': 0,
        'Turnover Prior Year Permanent Employees Total %': 0,
    }
    
    print(f"\nExtracting BRSR data for {company_name}...")
    
    # Extract the 3 core sections
    extract_employees_data(text, data)
    extract_board_kmp_data(text, data)
    extract_turnover_data(text, data)
    
    return data


def display_preview(data):
    """Display formatted preview of extracted data"""
    print(f"\n{'='*60}")
    print(f"DATA PREVIEW - {data['Company']}")
    print(f"{'='*60}\n")
    
    # Employees
    perm_total = data['Permanent Employees Male Number'] + data['Permanent Employees Female Number']
    other_total = data['Other than Permanent Employees Male Number'] + data['Other than Permanent Employees Female Number']
    total_total = data['Total Employees Male Number'] + data['Total Employees Female Number']
    
    print(f"📊 EMPLOYEES:")
    print(f"   Permanent: {perm_total:,} (M:{data['Permanent Employees Male Number']:,}/{data['Permanent Employees Male %']:.1f}% F:{data['Permanent Employees Female Number']:,}/{data['Permanent Employees Female %']:.1f}%)")
    print(f"   Other: {other_total:,} (M:{data['Other than Permanent Employees Male Number']:,}/{data['Other than Permanent Employees Male %']:.1f}% F:{data['Other than Permanent Employees Female Number']:,}/{data['Other than Permanent Employees Female %']:.1f}%)")
    print(f"   TOTAL: {total_total:,} (M:{data['Total Employees Male Number']:,}/{data['Total Employees Male %']:.1f}% F:{data['Total Employees Female Number']:,}/{data['Total Employees Female %']:.1f}%)\n")
    
    # Board & KMP
    print(f"👔 BOARD OF DIRECTORS:")
    print(f"   Total: {data['Board of Directors Total']}, Women: {data['Board of Directors Female Number']} ({data['Board of Directors Female %']:.2f}%)\n")
    
    print(f"🔑 KEY MANAGEMENT PERSONNEL:")
    print(f"   Total: {data['Key Management Personnel Total']}, Women: {data['Key Management Personnel Female Number']} ({data['Key Management Personnel Female %']:.1f}%)\n")
    
    # Turnover
    print(f"📈 TURNOVER RATES:")
    print(f"   Current Year: M:{data['Turnover Current Year Permanent Employees Male %']:.1f}% F:{data['Turnover Current Year Permanent Employees Female %']:.1f}% Total:{data['Turnover Current Year Permanent Employees Total %']:.1f}%")
    print(f"   Previous Year: M:{data['Turnover Previous Year Permanent Employees Male %']:.1f}% F:{data['Turnover Previous Year Permanent Employees Female %']:.1f}% Total:{data['Turnover Previous Year Permanent Employees Total %']:.1f}%")
    print(f"   Prior Year: M:{data['Turnover Prior Year Permanent Employees Male %']:.1f}% F:{data['Turnover Prior Year Permanent Employees Female %']:.1f}% Total:{data['Turnover Prior Year Permanent Employees Total %']:.1f}%")
    
    print(f"\n{'='*60}\n")


def update_csv(data, csv_file):
    """Update CSV file with extracted data"""
    
    fieldnames = [
        'Company',
        # Employees
        'Permanent Employees Male Number', 'Permanent Employees Male %',
        'Permanent Employees Female Number', 'Permanent Employees Female %',
        'Other than Permanent Employees Male Number', 'Other than Permanent Employees Male %',
        'Other than Permanent Employees Female Number', 'Other than Permanent Employees Female %',
        'Total Employees Male Number', 'Total Employees Male %',
        'Total Employees Female Number', 'Total Employees Female %',
        # Board & KMP
        'Board of Directors Total', 'Board of Directors Female Number', 'Board of Directors Female %',
        'Key Management Personnel Total', 'Key Management Personnel Female Number', 'Key Management Personnel Female %',
        # Turnover
        'Turnover Current Year Permanent Employees Male %',
        'Turnover Current Year Permanent Employees Female %',
        'Turnover Current Year Permanent Employees Total %',
        'Turnover Previous Year Permanent Employees Male %',
        'Turnover Previous Year Permanent Employees Female %',
        'Turnover Previous Year Permanent Employees Total %',
        'Turnover Prior Year Permanent Employees Male %',
        'Turnover Prior Year Permanent Employees Female %',
        'Turnover Prior Year Permanent Employees Total %',
    ]
    
    # Read existing data
    existing_data = []
    company_exists = False
    
    if os.path.exists(csv_file):
        with open(csv_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row['Company'].lower() == data['Company'].lower():
                    existing_data.append(data)
                    company_exists = True
                else:
                    existing_data.append(row)
    
    # Add new data if company doesn't exist
    if not company_exists:
        existing_data.append(data)
    
    # Write to CSV
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(existing_data)
    
    action = "Updated" if company_exists else "Added"
    print(f"  ✓ {action} {data['Company']} in CSV")


def batch_process_pdfs(folder_path, csv_file='brsr_simple_analysis.csv'):
    """Process all PDFs in folder and subfolders"""
    
    pdf_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    print(f"\n{'#'*60}")
    print(f"# BATCH PROCESSING: {len(pdf_files)} PDF files found")
    print(f"{'#'*60}\n")
    
    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"\n[{i}/{len(pdf_files)}] Processing: {Path(pdf_path).name}\n")
        
        # Extract text
        text, company_name = extract_text_from_pdf(pdf_path)
        if not text:
            continue
        
        # Extract BRSR data
        data = extract_brsr_data(text, company_name)
        
        # Display preview
        display_preview(data)
        
        # Update CSV
        update_csv(data, csv_file)
        
        print(f"✓ COMPLETED: {company_name}\n")
    
    print(f"\n{'='*60}")
    print(f"✓ BATCH PROCESSING COMPLETE!")
    print(f"  Processed: {len(pdf_files)} PDFs")
    print(f"  Output: {csv_file}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python brsr_simple_extractor.py <pdf_path_or_folder>")
        sys.exit(1)
    
    path = sys.argv[1]
    
    if os.path.isdir(path):
        # Batch process folder
        batch_process_pdfs(path)
    else:
        # Process single PDF
        text, company_name = extract_text_from_pdf(path)
        if text:
            data = extract_brsr_data(text, company_name)
            display_preview(data)
            update_csv(data, 'brsr_simple_analysis.csv')
