#!/usr/bin/env python3
"""
Google Sheets Access Verification Script

This script authenticates with Google Sheets API using a service account
and reads data from a specified spreadsheet to verify access.
"""

import os
import sys
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables from .env file
load_dotenv()

# Configuration
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

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


def get_spreadsheet_info(service, spreadsheet_id):
    """Get basic information about the spreadsheet."""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return spreadsheet
    except HttpError as e:
        if e.resp.status == 404:
            raise ValueError(
                f"Spreadsheet not found. Please check the spreadsheet ID: {spreadsheet_id}\n"
                f"Error: {str(e)}"
            )
        elif e.resp.status == 403:
            raise PermissionError(
                f"Permission denied. The service account does not have access to this spreadsheet.\n"
                f"Please share the spreadsheet with: hexa-service@sheets-api-473619.iam.gserviceaccount.com\n"
                f"Error: {str(e)}"
            )
        else:
            raise RuntimeError(f"Failed to access spreadsheet: {str(e)}")
    except Exception as e:
        raise RuntimeError(f"Unexpected error accessing spreadsheet: {str(e)}")


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
                f"Error: {str(e)}"
            )
        else:
            raise RuntimeError(f"Failed to read sheet data: {str(e)}")
    except Exception as e:
        raise RuntimeError(f"Unexpected error reading sheet data: {str(e)}")


def display_data(data, max_rows=10):
    """Display spreadsheet data in a readable format."""
    if not data:
        print("No data found in the sheet.")
        return
    
    print(f"\n{'='*60}")
    print(f"Sheet Data (showing up to {min(len(data), max_rows)} rows):")
    print(f"{'='*60}\n")
    
    for i, row in enumerate(data[:max_rows], 1):
        print(f"Row {i}: {row}")
    
    if len(data) > max_rows:
        print(f"\n... ({len(data) - max_rows} more rows not shown)")
    
    print(f"\n{'='*60}")
    print(f"Total rows: {len(data)}")
    print(f"{'='*60}\n")


def main():
    """Main function to verify Google Sheets access."""
    print("Google Sheets Access Verification")
    print("=" * 60)
    
    try:
        # Load credentials
        print(f"\n1. Loading credentials from {SERVICE_ACCOUNT_FILE}...")
        credentials = load_credentials()
        print("   ✓ Credentials loaded successfully")
        
        # Create Sheets service
        print("\n2. Creating Google Sheets API service...")
        service = create_sheets_service(credentials)
        print("   ✓ Service created successfully")
        
        # Get spreadsheet info
        print(f"\n3. Accessing spreadsheet: {SPREADSHEET_ID}...")
        spreadsheet_info = get_spreadsheet_info(service, SPREADSHEET_ID)
        print("   ✓ Spreadsheet accessed successfully")
        
        # Display spreadsheet metadata
        print(f"\n   Spreadsheet Title: {spreadsheet_info.get('properties', {}).get('title', 'N/A')}")
        sheets = spreadsheet_info.get('sheets', [])
        print(f"   Number of sheets: {len(sheets)}")
        
        if sheets:
            sheet_names = [sheet['properties']['title'] for sheet in sheets]
            print(f"   Sheet names: {', '.join(sheet_names)}")
            
            # Read data from the first sheet
            first_sheet_name = sheet_names[0]
            print(f"\n4. Reading data from sheet: '{first_sheet_name}'...")
            data = read_sheet_data(service, SPREADSHEET_ID, first_sheet_name)
            print(f"   ✓ Data read successfully ({len(data)} rows)")
            
            # Display the data
            display_data(data)
        
        print("\n" + "=" * 60)
        print("SUCCESS: Google Sheets access verified!")
        print("=" * 60)
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

