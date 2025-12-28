# Activity Time Tracker

Tracks how much time you spend across applications and browser tabs, log daily totals to Google Sheets, and generate a monthly pie chart. `applicable as a time-sheet generator app in a corporate setting`

## Features

    Tracks active window / tab every few seconds

    Aggregates total time per app/tab per day

    Automatically logs results daily at 8:00 PM

    Uploads data to Google Sheets (one worksheet per month)

    Generates a monthly pie chart (top 10 apps)

    Persists data locally to survive restarts/crashes

    Idempotent sheet writes (safe to re-run)


### Supports

Windows: (Fully supported) macOS: Supported (uses AppKit) Linux: (Requires xdotool)

### Install Dependencies

```bash
   python -m pip install psutil pygetwindow gspread oauth2client schedule matplotlib pandas
```

Run `activity_tracker.py`

Data is stored locally in: `activity_data.json`


# Google Sheet/s Setup

Go to <https://console.cloud.google.com/>

Enable Drive + Sheet API. Create a Service Account

Download JSON key file, rename Credentials.json

Share Google Sheet with Service Account Email


## Monthly Summary**

```bash
   #- Aggregates monthly totals
   #- Generates a pie chart of top 10 applications
```