#!/usr/bin/env python3
"""
Stock Replenishment Report Generator

This script accesses the STOCK SHEET, identifies items where Opening Balance
is below Minimum Level, and generates both a terminal display and CSV export
of items needing replenishment.
"""

import os
import sys
import csv
import re
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tabulate import tabulate

# Load environment variables from .env file
load_dotenv()

# Configuration
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE')
SHEET_NAME = 'STOCK SHEET (Add New Item here)'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
OUTPUT_CSV = 'replenishment_items.csv'

# Validate required environment variables
if not SPREADSHEET_ID:
    raise ValueError("SPREADSHEET_ID environment variable is required. Please set it in .env file.")
if not SERVICE_ACCOUNT_FILE:
    raise ValueError("SERVICE_ACCOUNT_FILE environment variable is required. Please set it in .env file.")


def load_credentials():
    """Load service account credentials from JSON file."""
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(
            f"Service account file not found: {SERVICE_ACCOUNT_FILE}\n"
            f"Please ensure the file exists in the current directory."
        )
    
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=SCOPES
        )
        return credentials
    except Exception as e:
        raise ValueError(
            f"Failed to load credentials from {SERVICE_ACCOUNT_FILE}:\n{str(e)}"
        )


def create_sheets_service(credentials):
    """Create and return a Google Sheets API service object."""
    try:
        service = build('sheets', 'v4', credentials=credentials)
        return service
    except Exception as e:
        raise RuntimeError(f"Failed to create Sheets API service: {str(e)}")


def read_sheet_data(service, spreadsheet_id, sheet_name, range_name=None):
    """Read data from a specific sheet in the spreadsheet."""
    if range_name:
        range_to_read = f"{sheet_name}!{range_name}"
    else:
        range_to_read = sheet_name
    
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_to_read
        ).execute()
        return result.get('values', [])
    except HttpError as e:
        if e.resp.status == 404:
            raise ValueError(f"Sheet or range not found: {range_to_read}")
        elif e.resp.status == 403:
            raise PermissionError(
                f"Permission denied when reading sheet data.\n"
                f"Please share the spreadsheet with: hexa-service@sheets-api-473619.iam.gserviceaccount.com\n"
                f"Error: {str(e)}"
            )
        else:
            raise RuntimeError(f"Failed to read sheet data: {str(e)}")
    except Exception as e:
        raise RuntimeError(f"Unexpected error reading sheet data: {str(e)}")


def normalize_text(text):
    """Normalize text for comparison (lowercase, remove extra spaces/punctuation)."""
    if not text:
        return ""
    # Convert to string, lowercase, remove extra spaces
    text = str(text).lower().strip()
    # Remove extra spaces
    text = re.sub(r'\s+', ' ', text)
    return text


def find_column_index(header_row, search_terms):
    """
    Find column index by searching for one of the provided terms in the header row.
    
    Args:
        header_row: List of header cell values
        search_terms: List of possible search terms (case-insensitive)
    
    Returns:
        Column index if found, None otherwise
    """
    for idx, cell in enumerate(header_row):
        normalized_cell = normalize_text(cell)
        for term in search_terms:
            normalized_term = normalize_text(term)
            # Check for exact match or if term is contained in cell
            if normalized_cell == normalized_term or normalized_term in normalized_cell:
                return idx
    return None


def find_header_row(data, max_rows_to_check=10):
    """
    Find the header row by searching for 'Sno' or 'UID' in the first few rows.
    
    Returns:
        Tuple of (header_row_index, header_row_data)
    """
    for i in range(min(max_rows_to_check, len(data))):
        row = data[i]
        if not row:
            continue
        
        # Check if this row contains header indicators
        row_text = ' '.join(str(cell) for cell in row).lower()
        if 'sno' in row_text and ('uid' in row_text or 'item name' in row_text):
            return i, row
    
    raise ValueError("Could not find header row. Expected to find 'Sno' and 'UID' columns.")


def find_all_columns(header_row):
    """
    Find all required column indices dynamically.
    
    Returns:
        Dictionary mapping column names to indices
    """
    columns = {}
    
    # Define search terms for each column
    column_definitions = {
        'sno': [['sno', 'sno.']],
        'uid': [['uid']],
        'bush': [['bush']],
        'group': [['group']],
        'last_io_raised': [['last i. o. raised', 'last i.o raised', 'last io raised']],
        'category': [['category']],
        'stock_location': [['stock location']],
        'min_lvl': [['min lvl', 'min. lvl', 'min lv', 'min. lv', 'minimum level']],
        'opn_bal': [['opn. bal', 'opn bal', 'opening balance']]
    }
    
    for col_name, search_terms_list in column_definitions.items():
        # Flatten the search terms list
        search_terms = [term for sublist in search_terms_list for term in sublist]
        idx = find_column_index(header_row, search_terms)
        if idx is None:
            raise ValueError(
                f"Required column '{col_name}' not found in header row.\n"
                f"Searched for: {', '.join(search_terms)}\n"
                f"Header row: {header_row}"
            )
        columns[col_name] = idx
    
    return columns


def safe_float(value):
    """Safely convert a value to float, returning None if conversion fails."""
    if value is None or value == '':
        return None
    try:
        # Remove any commas or other formatting
        cleaned = str(value).replace(',', '').strip()
        return float(cleaned)
    except (ValueError, TypeError):
        return None


def extract_row_data(row, columns):
    """Extract data from a row using column indices."""
    def get_cell(idx):
        if idx < len(row):
            return row[idx]
        return ""
    
    return {
        'sno': get_cell(columns['sno']),
        'uid': get_cell(columns['uid']),
        'bush': get_cell(columns['bush']),
        'group': get_cell(columns['group']),
        'last_io_raised': get_cell(columns['last_io_raised']),
        'category': get_cell(columns['category']),
        'stock_location': get_cell(columns['stock_location']),
        'opn_bal': get_cell(columns['opn_bal']),
        'min_lvl': get_cell(columns['min_lvl'])
    }


def filter_replenishment_items(data, header_row_idx, columns):
    """
    Filter rows where OPN. BAL < MIN LVL.
    
    Returns:
        List of dictionaries containing items needing replenishment
    """
    replenishment_items = []
    
    for i in range(header_row_idx + 1, len(data)):
        row = data[i]
        
        # Skip empty rows
        if not row or all(not cell for cell in row):
            continue
        
        # Extract row data
        row_data = extract_row_data(row, columns)
        
        # Convert OPN. BAL and MIN LVL to floats
        opn_bal = safe_float(row_data['opn_bal'])
        min_lvl = safe_float(row_data['min_lvl'])
        
        # Skip if we can't parse the values
        if opn_bal is None or min_lvl is None:
            continue
        
        # Check if OPN. BAL < MIN LVL
        if opn_bal < min_lvl:
            # Add to replenishment list (exclude the numeric fields from output)
            item = {
                'Sno.': row_data['sno'],
                'UID': row_data['uid'],
                'Bush': row_data['bush'],
                'Group': row_data['group'],
                'Last I.O Raised': row_data['last_io_raised'],
                'Category': row_data['category'],
                'Stock Location': row_data['stock_location']
            }
            replenishment_items.append(item)
    
    return replenishment_items


def display_results(items):
    """Display results in a formatted table in the terminal."""
    if not items:
        print("\n" + "=" * 80)
        print("No items found that need replenishment.")
        print("All items have Opening Balance >= Minimum Level.")
        print("=" * 80)
        return
    
    # Prepare data for tabulate
    headers = ['Sno.', 'UID', 'Bush', 'Group', 'Last I.O Raised', 'Category', 'Stock Location']
    table_data = []
    
    for item in items:
        table_data.append([
            item['Sno.'],
            item['UID'],
            item['Bush'],
            item['Group'],
            item['Last I.O Raised'],
            item['Category'],
            item['Stock Location']
        ])
    
    print("\n" + "=" * 80)
    print(f"ITEMS NEEDING REPLENISHMENT (OPN. BAL < MIN LVL)")
    print(f"Total items: {len(items)}")
    print("=" * 80)
    print()
    print(tabulate(table_data, headers=headers, tablefmt='grid', maxcolwidths=[5, 40, 8, 8, 20, 10, 15]))
    print()


def save_to_csv(items, filename):
    """Save results to a CSV file."""
    if not items:
        print(f"\nNo items to save to CSV.")
        return
    
    headers = ['Sno.', 'UID', 'Bush', 'Group', 'Last I.O Raised', 'Category', 'Stock Location']
    
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=headers)
            writer.writeheader()
            writer.writerows(items)
        print(f"\n✓ Results saved to: {filename}")
    except Exception as e:
        raise RuntimeError(f"Failed to save CSV file: {str(e)}")


def main():
    """Main function to generate stock replenishment report."""
    print("Stock Replenishment Report Generator")
    print("=" * 80)
    
    try:
        # Load credentials
        print(f"\n1. Loading credentials from {SERVICE_ACCOUNT_FILE}...")
        credentials = load_credentials()
        print("   ✓ Credentials loaded successfully")
        
        # Create Sheets service
        print("\n2. Creating Google Sheets API service...")
        service = create_sheets_service(credentials)
        print("   ✓ Service created successfully")
        
        # Read sheet data
        print(f"\n3. Reading data from sheet: '{SHEET_NAME}'...")
        data = read_sheet_data(service, SPREADSHEET_ID, SHEET_NAME)
        print(f"   ✓ Data read successfully ({len(data)} rows)")
        
        # Find header row
        print("\n4. Identifying header row...")
        header_row_idx, header_row = find_header_row(data)
        print(f"   ✓ Header row found at index {header_row_idx}")
        
        # Find all required columns
        print("\n5. Finding required columns...")
        columns = find_all_columns(header_row)
        print("   ✓ All required columns found:")
        for col_name, idx in columns.items():
            header_value = header_row[idx] if idx < len(header_row) else "N/A"
            print(f"      - {col_name}: column {idx} ('{header_value}')")
        
        # Filter replenishment items
        print("\n6. Filtering items where OPN. BAL < MIN LVL...")
        replenishment_items = filter_replenishment_items(data, header_row_idx, columns)
        print(f"   ✓ Found {len(replenishment_items)} items needing replenishment")
        
        # Display results
        print("\n7. Displaying results...")
        display_results(replenishment_items)
        
        # Save to CSV
        print("\n8. Saving results to CSV...")
        save_to_csv(replenishment_items, OUTPUT_CSV)
        
        print("\n" + "=" * 80)
        print("SUCCESS: Stock replenishment report generated!")
        print("=" * 80)
        return 0
        
    except FileNotFoundError as e:
        print(f"\n❌ ERROR: {str(e)}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"\n❌ ERROR: {str(e)}", file=sys.stderr)
        return 1
    except PermissionError as e:
        print(f"\n❌ ERROR: {str(e)}", file=sys.stderr)
        return 1
    except RuntimeError as e:
        print(f"\n❌ ERROR: {str(e)}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"\n❌ UNEXPECTED ERROR: {str(e)}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())

