"""
Manual upload script - uploads any date's data to Google Sheets
Use this if you missed the automatic 6 PM upload
"""

import json
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

CREDENTIALS_FILE = "credentials.json"
GOOGLE_SHEET_NAME = "Activity Tracker"
DATA_FILE = "activity_data.json"

def format_time(minutes):
    """Convert minutes to hours and minutes format"""
    hours = int(minutes // 60)
    mins = int(minutes % 60)
    return f"{hours}h {mins}m"

def upload_date(date_str):
    """Upload specific date's data to Google Sheets"""
    
    # Load activity data
    try:
        with open(DATA_FILE, 'r') as f:
            activity_data = json.load(f)
    except FileNotFoundError:
        print(f"Error: {DATA_FILE} not found!")
        return False
    
    if date_str not in activity_data:
        print(f"No data found for {date_str}")
        print(f"Available dates: {list(activity_data.keys())}")
        return False
    
    print(f"\nUploading data for {date_str}...")
    
    # Sort activities by time spent
    sorted_activities = sorted(
        activity_data[date_str].items(),
        key=lambda x: x[1],
        reverse=True
    )
    
    print(f"Found {len(sorted_activities)} applications/tabs")
    
    # Setup Google Sheets API
    try:
        scope = ['https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive']
        
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            CREDENTIALS_FILE, scope
        )
        client = gspread.authorize(creds)
        
        # Open spreadsheet
        sheet = client.open(GOOGLE_SHEET_NAME)
        
        # Get or create worksheet for current month
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        month_name = date_obj.strftime("%B %Y")
        
        try:
            worksheet = sheet.worksheet(month_name)
            print(f"Using existing worksheet: {month_name}")
        except:
            worksheet = sheet.add_worksheet(title=month_name, rows=100, cols=10)
            worksheet.append_row(["Date", "Application/Tab", "Time (minutes)", "Time (formatted)"])
            print(f"Created new worksheet: {month_name}")
        
        # Prepare all rows for batch upload
        rows_to_upload = []
        for app, minutes in sorted_activities:
            rows_to_upload.append([
                date_str,
                app,
                round(minutes, 2),
                format_time(minutes)
            ])
        
        # Batch upload all rows at once (much faster, avoids rate limits)
        if rows_to_upload:
            worksheet.append_rows(rows_to_upload)
        
        print(f"✓ Successfully uploaded {len(sorted_activities)} entries to Google Sheets!")
        print(f"  Spreadsheet: {GOOGLE_SHEET_NAME}")
        print(f"  Worksheet: {month_name}")
        
        return True
        
    except Exception as e:
        print(f"Error uploading to Google Sheets: {e}")
        return False

def main():
    print("="*60)
    print("Manual Data Upload to Google Sheets")
    print("="*60)
    print()
    
    # Load available dates
    try:
        with open(DATA_FILE, 'r') as f:
            activity_data = json.load(f)
        
        if not activity_data:
            print("No data available to upload!")
            return
        
        print("Available dates:")
        for date in sorted(activity_data.keys(), reverse=True):
            print(f"  - {date}")
        print()
        
    except FileNotFoundError:
        print(f"Error: {DATA_FILE} not found!")
        print("Make sure the tracker has been running and collecting data.")
        return
    
    # Get date from user
    print("Options:")
    print("  1. Upload today's data")
    print("  2. Upload yesterday's data")
    print("  3. Upload specific date")
    print("  4. Upload ALL dates")
    print()
    
    choice = input("Enter choice (1-4): ").strip()
    
    if choice == "1":
        date_str = str(datetime.date.today())
        upload_date(date_str)
        
    elif choice == "2":
        date_str = str(datetime.date.today() - datetime.timedelta(days=1))
        upload_date(date_str)
        
    elif choice == "3":
        date_str = input("Enter date (YYYY-MM-DD): ").strip()
        upload_date(date_str)
        
    elif choice == "4":
        print("\nUploading all dates...")
        success_count = 0
        failed_dates = []
        
        for date_str in sorted(activity_data.keys()):
            print(f"\n{'='*60}")
            try:
                if upload_date(date_str):
                    success_count += 1
                    # Small delay to avoid rate limits between dates
                    import time
                    time.sleep(2)
                else:
                    failed_dates.append(date_str)
            except Exception as e:
                print(f"✗ Failed to upload {date_str}: {e}")
                failed_dates.append(date_str)
        
        print(f"\n{'='*60}")
        print(f"UPLOAD COMPLETE")
        print(f"{'='*60}")
        print(f"✓ Successfully uploaded: {success_count} out of {len(activity_data)} dates")
        
        if failed_dates:
            print(f"✗ Failed dates: {', '.join(failed_dates)}")
        
    else:
        print("Invalid choice!")

if __name__ == "__main__":
    main()
    print()
    input("Press Enter to exit...")
