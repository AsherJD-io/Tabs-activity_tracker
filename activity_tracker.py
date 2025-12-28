"""
Activity Time Tracker - Monitors active windows/tabs and logs to Google Sheets
Requires: pip install psutil pygetwindow gspread oauth2client schedule matplotlib pandas
"""

import psutil
import time
import datetime
from collections import defaultdict
import json
import os
import schedule
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import matplotlib.pyplot as plt

# Configuration
TRACKING_INTERVAL = 5  # Check active window every 5 seconds
DAILY_REPORT_TIME = "20:00"  # 8 PM
GOOGLE_SHEET_NAME = "Activity Tracker"
CREDENTIALS_FILE = "credentials.json"  # Google Sheets API credentials
DATA_FILE = "activity_data.json"

class ActivityTracker:
    def __init__(self):
        self.activity_data = self.load_data()
        self.current_date = datetime.date.today()
        self.last_active_window = None
        self.last_check_time = time.time()
        
    def load_data(self):
        """Load existing activity data from file"""
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, 'r') as f:
                return json.load(f)
        return {}
    
    def save_data(self):
        """Save activity data to file"""
        with open(DATA_FILE, 'w') as f:
            json.dump(self.activity_data, f, indent=2)
            f.flush()  # Force write to disk immediately
            os.fsync(f.fileno())  # Ensure it's written to disk
    
    def get_active_window(self):
        """Get the currently active window title and process name"""
        try:
            # For Windows
            import pygetwindow as gw
            active_window = gw.getActiveWindow()
            if active_window:
                return active_window.title
        except:
            pass
        
        try:
            # For macOS
            from AppKit import NSWorkspace # type: ignore
            active_app = NSWorkspace.sharedWorkspace().activeApplication()
            return active_app['NSApplicationName']
        except:
            pass
        
        try:
            # For Linux
            import subprocess
            window_id = subprocess.check_output(['xdotool', 'getactivewindow']).decode().strip()
            window_name = subprocess.check_output(['xdotool', 'getwindowname', window_id]).decode().strip()
            return window_name
        except:
            pass
        
        return "Unknown"
    
    def track_activity(self):
        """Track current activity and update time spent"""
        current_time = time.time()
        time_diff = current_time - self.last_check_time
        
        active_window = self.get_active_window()
        
        if active_window and active_window != "Unknown":
            date_key = str(self.current_date)
            
            if date_key not in self.activity_data:
                self.activity_data[date_key] = {}
            
            if active_window not in self.activity_data[date_key]:
                self.activity_data[date_key][active_window] = 0
            
            # Add time spent (convert to minutes)
            self.activity_data[date_key][active_window] += time_diff / 60
        
        self.last_check_time = current_time
        self.last_active_window = active_window
        
        # Save data more frequently for immediate updates
        self.save_data()
    
    def format_time(self, minutes):
        """Convert minutes to hours and minutes format"""
        hours = int(minutes // 60)
        mins = int(minutes % 60)
        return f"{hours}h {mins}m"
    
    def generate_daily_report(self):
        """Generate daily report and upload to Google Sheets"""
        date_key = str(self.current_date)
        
        if date_key not in self.activity_data:
            print(f"No data for {date_key}")
            return
        
        print(f"\n=== Daily Report for {date_key} ===")
        
        # Sort by time spent
        sorted_activities = sorted(
            self.activity_data[date_key].items(),
            key=lambda x: x[1],
            reverse=True
        )
        
        for app, minutes in sorted_activities:
            print(f"{app}: {self.format_time(minutes)}")
        
        # Upload to Google Sheets
        self.upload_to_sheets(date_key, sorted_activities)
        
        # Check if it's end of month
        tomorrow = self.current_date + datetime.timedelta(days=1)
        if tomorrow.month != self.current_date.month:
            self.generate_monthly_graph()
    
    def upload_to_sheets(self, date_key, activities):
        """Upload daily data to Google Sheets"""
        try:
            # Setup Google Sheets API
            scope = ['https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive']
            
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                CREDENTIALS_FILE, scope
            )
            client = gspread.authorize(creds)
            
            # Open existing spreadsheet (don't try to create)
            try:
                sheet = client.open(GOOGLE_SHEET_NAME)
            except Exception as e:
                print(f"Error: Could not open '{GOOGLE_SHEET_NAME}' spreadsheet.")
                print(f"Please create it manually at sheets.google.com and share with service account.")
                print(f"Error details: {e}")
                return
            
            # Get or create worksheet for current month
            month_name = datetime.datetime.strptime(date_key, "%Y-%m-%d").strftime("%B %Y")
            try:
                worksheet = sheet.worksheet(month_name)
            except:
                worksheet = sheet.add_worksheet(title=month_name, rows=100, cols=10)
                worksheet.append_row(["Date", "Application/Tab", "Time (minutes)", "Time (formatted)"])
            
            # Check if data for this date already exists (make idempotent)
            try:
                all_values = worksheet.get_all_values()
                existing_dates = [row[0] for row in all_values[1:] if row]  # Skip header
                
                if date_key in existing_dates:
                    print(f"⚠ Data for {date_key} already exists in sheet. Removing old entries...")
                    # Find and delete rows with this date
                    rows_to_delete = []
                    for idx, row in enumerate(all_values[1:], start=2):  # Start from row 2 (after header)
                        if row and row[0] == date_key:
                            rows_to_delete.append(idx)
                    
                    # Delete in reverse order to maintain row indices
                    for row_idx in reversed(rows_to_delete):
                        worksheet.delete_rows(row_idx)
                    print(f"✓ Removed {len(rows_to_delete)} old entries for {date_key}")
            except Exception as e:
                print(f"Note: Could not check for existing data: {e}")
            
            # Prepare all rows for batch upload (avoids rate limits)
            rows_to_upload = []
            for app, minutes in activities:
                rows_to_upload.append([
                    date_key,
                    app,
                    round(minutes, 2),
                    self.format_time(minutes)
                ])
            
            # Batch upload all rows at once
            if rows_to_upload:
                worksheet.append_rows(rows_to_upload)
            
            print(f"✓ Data uploaded to Google Sheets: {GOOGLE_SHEET_NAME}")
            
        except Exception as e:
            print(f"Error uploading to Google Sheets: {e}")
            print("Make sure you have set up Google Sheets API credentials")
    
    def generate_monthly_graph(self):
        """Generate monthly activity graph"""
        current_month = self.current_date.replace(day=1)
        month_data = {}
        
        # Aggregate data for current month
        for date_str, activities in self.activity_data.items():
            date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            if date.year == current_month.year and date.month == current_month.month:
                for app, minutes in activities.items():
                    if app not in month_data:
                        month_data[app] = 0
                    month_data[app] += minutes
        
        if not month_data:
            print("No data available for monthly graph")
            return
        
        # Sort and get top 10 applications
        sorted_data = sorted(month_data.items(), key=lambda x: x[1], reverse=True)[:10]
        apps = [item[0][:30] + "..." if len(item[0]) > 30 else item[0] for item in sorted_data]
        hours = [item[1] / 60 for item in sorted_data]  # Convert to hours
        
        # Create pie chart
        plt.figure(figsize=(12, 8))
        colors = plt.cm.Set3(range(len(apps)))
        
        # Create pie chart with percentage labels
        wedges, texts, autotexts = plt.pie(
            hours, 
            labels=apps, 
            autopct='%1.1f%%',
            startangle=90,
            colors=colors,
            textprops={'fontsize': 10}
        )
        
        # Make percentage text bold and white
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(9)
        
        plt.title(f'Top 10 Applications - {current_month.strftime("%B %Y")}', 
                 fontsize=16, fontweight='bold', pad=20)
        
        # Add legend with hours
        legend_labels = [f'{app}: {hour:.1f}h' for app, hour in zip(apps, hours)]
        plt.legend(legend_labels, loc='center left', bbox_to_anchor=(1, 0, 0.5, 1))
        
        plt.tight_layout()
        
        filename = f'activity_report_{current_month.strftime("%Y_%m")}.png'
        plt.savefig(filename, dpi=300, bbox_inches='tight')
        print(f"✓ Monthly pie chart saved: {filename}")
        plt.close()
    
    def run(self):
        """Main loop to run the tracker"""
        print("Activity Tracker Started")
        print(f"Tracking active windows every {TRACKING_INTERVAL} seconds")
        print(f"Daily reports at {DAILY_REPORT_TIME}")
        print("Press Ctrl+C to stop\n")
        
        # Schedule daily report
        schedule.every().day.at(DAILY_REPORT_TIME).do(self.generate_daily_report)
        
        try:
            while True:
                # Check if date has changed
                if datetime.date.today() != self.current_date:
                    self.current_date = datetime.date.today()
                
                # Track activity
                self.track_activity()
                
                # Run scheduled tasks
                schedule.run_pending()
                
                # Wait for next interval
                time.sleep(TRACKING_INTERVAL)
                
        except KeyboardInterrupt:
            print("\n\nStopping tracker...")
            self.save_data()
            print("Data saved. Goodbye!")

if __name__ == "__main__":
    # Setup instructions
    print("""
    Setup Instructions:
    
    1. Install required packages:
       pip install psutil pygetwindow gspread oauth2client schedule matplotlib pandas
    
    2. Set up Google Sheets API:
       - Go to https://console.cloud.google.com/
       - Create a new project
       - Enable Google Sheets API
       - Create a Service Account
       - Download credentials.json and place in same folder as this script
    
    3. Run the script:
       python activity_tracker.py
    
    The script will:
    - Track active windows every 5 seconds
    - Generate daily reports at 8 PM
    - Upload data to Google Sheets
    - Create monthly graphs at end of each month
    """)
    
    input("Press Enter to start tracking...")
    
    tracker = ActivityTracker()
    tracker.run()