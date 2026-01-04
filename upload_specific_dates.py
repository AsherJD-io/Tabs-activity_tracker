"""
Upload only specific dates (Dec 27-29)
"""

import json
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

CREDENTIALS_FILE = "credentials.json"
GOOGLE_SHEET_NAME = "Activity Tracker"
DATA_FILE = "activity_data.json"

def format_time(minutes):
    hours = int(minutes // 60)
    mins = int(minutes % 60)
    return f"{hours}h {mins}m"

# Dates to upload
dates_to_upload = ["2025-12-27", "2025-12-28", "2025-12-29"]

print("="*60)
print("UPLOADING SPECIFIC DATES")
print("="*60)
print()

# Load data
try:
    with open(DATA_FILE, 'r') as f:
        activity_data = json.load(f)
except:
    print("Error: Could not load activity_data.json")
    input("Press Enter to exit...")
    exit()

# Setup Google Sheets API
try:
    scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
    
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        CREDENTIALS_FILE, scope
    )
    client = gspread.authorize(creds)
    sheet = client.open(GOOGLE_SHEET_NAME)
    worksheet = sheet.worksheet("December 2025")
    
    print(f"Connected to: {GOOGLE_SHEET_NAME}")
    print(f"Worksheet: December 2025")
    print()
    
    for date_str in dates_to_upload:
        if date_str not in activity_data:
            print(f"⊘ {date_str} - No data in activity_data.json")
            continue
        
        print(f"\n{'='*60}")
        print(f"Uploading: {date_str}")
        
        sorted_activities = sorted(
            activity_data[date_str].items(),
            key=lambda x: x[1],
            reverse=True
        )
        
        print(f"Found {len(sorted_activities)} apps/tabs")
        
        # Prepare rows
        rows_to_upload = []
        for app, minutes in sorted_activities:
            rows_to_upload.append([
                date_str,
                app[:500],  # Truncate long names
                round(minutes, 2),
                format_time(minutes)
            ])
        
        # Upload in batches
        batch_size = 500
        total = 0
        
        try:
            for i in range(0, len(rows_to_upload), batch_size):
                batch = rows_to_upload[i:i + batch_size]
                print(f"  Batch {i//batch_size + 1}: Uploading {len(batch)} rows...")
                
                worksheet.append_rows(batch)
                total += len(batch)
                print(f"  ✓ Progress: {total}/{len(rows_to_upload)}")
                
                if i + batch_size < len(rows_to_upload):
                    import time
                    time.sleep(3)
            
            print(f"✓ {date_str} - Successfully uploaded {total} entries")
            
        except Exception as e:
            print(f"✗ {date_str} - Error: {e}")
            import traceback
            traceback.print_exc()
    
    print()
    print("="*60)
    print("UPLOAD COMPLETE")
    print("="*60)
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()

print()
input("Press Enter to exit...")
