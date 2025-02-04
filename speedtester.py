import speedtest
import pandas as pd
import time
from datetime import datetime

# Configuration
interval = 60  # Interval in seconds (modify as needed)
excel_file = "internet_speed_log.xlsx"

def test_speed():
    """Sppedtest kore Mbps return korbe"""
    st = speedtest.Speedtest()
    st.get_best_server()
    download_speed = st.download() / 1_000_000  # Convert to Mbps
    return round(download_speed, 2)

def log_speed():
    """Logs results into Excel file."""
    while True:
        speed = test_speed()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Load existing data or create new DataFrame
        try:
            df = pd.read_excel(excel_file)
        except FileNotFoundError:
            df = pd.DataFrame(columns=["Timestamp", "Download Speed (Mbps)"])

        # Append new data
        new_entry = pd.DataFrame([[timestamp, speed]], columns=df.columns)
        df = pd.concat([df, new_entry], ignore_index=True)

        # Save to Excel
        df.to_excel(excel_file, index=False, engine='openpyxl')
        print(f"Logged: {timestamp} - {speed} Mbps")

        # Wait for next test
        time.sleep(interval)

if __name__ == "__main__":
    log_speed()
