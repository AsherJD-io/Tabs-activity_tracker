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
<<<<<<< HEAD
DAILY_REPORT_TIME = "20:00"  # 8 PM
=======
DAILY_REPORT_TIME = "18:00"  # 6 PM
>>>>>>> 5c0110aebe403c6ece1a33f75f6c631d5c7a1517
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
<<<<<<< HEAD
            f.flush()  # Force write to disk immediately
            os.fsync(f.fileno())  # Ensure it's written to disk
=======
>>>>>>> 5c0110aebe403c6ece1a33f75f6c631d5c7a1517
    
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
<<<<<<< HEAD
            from AppKit import NSWorkspace # type: ignore
=======
            from AppKit import NSWorkspace
>>>>>>> 5c0110aebe403c6ece1a33f75f6c631d5c7a1517
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
        
<<<<<<< HEAD
        # Save data more frequently for immediate updates
        self.save_data()
=======
        # Save data periodically
        if int(current_time) % 60 == 0:  # Save every minute
            self.save_data()
>>>>>>> 5c0110aebe403c6ece1a33f75f6c631d5c7a1517
    
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
<<<<<<< HEAD
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive.file']
=======
                    'https://www.googleapis.com/auth/drive']
>>>>>>> 5c0110aebe403c6ece1a33f75f6c631d5c7a1517
            
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                CREDENTIALS_FILE, scope
            )
            client = gspread.authorize(creds)
            
<<<<<<< HEAD
            # Open existing spreadsheet (don't try to create)
            try:
                sheet = client.open(GOOGLE_SHEET_NAME)
            except Exception as e:
                print(f"Error: Could not open '{GOOGLE_SHEET_NAME}' spreadsheet.")
                print(f"Please create it manually at sheets.google.com and share with service account.")
                print(f"Error details: {e}")
                return
=======
            # Open or create spreadsheet
            try:
                sheet = client.open(GOOGLE_SHEET_NAME)
            except:
                sheet = client.create(GOOGLE_SHEET_NAME)
>>>>>>> 5c0110aebe403c6ece1a33f75f6c631d5c7a1517
            
            # Get or create worksheet for current month
            month_name = datetime.datetime.strptime(date_key, "%Y-%m-%d").strftime("%B %Y")
            try:
                worksheet = sheet.worksheet(month_name)
            except:
                worksheet = sheet.add_worksheet(title=month_name, rows=100, cols=10)
                worksheet.append_row(["Date", "Application/Tab", "Time (minutes)", "Time (formatted)"])
            
            # Append data
            for app, minutes in activities:
                worksheet.append_row([
                    date_key,
                    app,
                    round(minutes, 2),
                    self.format_time(minutes)
                ])
            
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
        apps = [item[0][:30] for item in sorted_data]  # Truncate long names
        hours = [item[1] / 60 for item in sorted_data]  # Convert to hours
        
        # Create graph
        plt.figure(figsize=(12, 8))
        plt.barh(apps, hours, color='steelblue')
        plt.xlabel('Hours Spent')
        plt.ylabel('Application/Tab')
        plt.title(f'Top 10 Applications - {current_month.strftime("%B %Y")}')
        plt.tight_layout()
        
        filename = f'activity_report_{current_month.strftime("%Y_%m")}.png'
        plt.savefig(filename, dpi=300, bbox_inches='tight')
        print(f"✓ Monthly graph saved: {filename}")
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
<<<<<<< HEAD
    - Generate daily reports at 8 PM
=======
    - Generate daily reports at 6 PM
>>>>>>> 5c0110aebe403c6ece1a33f75f6c631d5c7a1517
    - Upload data to Google Sheets
    - Create monthly graphs at end of each month
    """)
    
    input("Press Enter to start tracking...")
    
    tracker = ActivityTracker()
    tracker.run()
<<<<<<< HEAD
import os
print("Current working directory:", os.getcwd())
print("Files here:", os.listdir())
=======
>>>>>>> 5c0110aebe403c6ece1a33f75f6c631d5c7a1517
