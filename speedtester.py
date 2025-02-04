import speedtest
import pandas as pd
import time
from datetime import datetime
import os

# Function to test download speed
def test_speed(url):
    st = speedtest.Speedtest()
    st.get_best_server()
    st.download()
    return round(st.results.download / 1_000_000, 2)  # Convert to Mbps

# Get user inputs
url = "https://novoserve.com/"  #input("Enter the URL to hit (this is ignored, as speedtest selects the best server): ")
interval = .5 * 60 #float(input("Enter the interval in minutes: "))
end_time_str = "04/02/2025 07:21:00 PM" #input("Enter the end time (dd/mm/yyyy hh:mm:ss AM/PM): ")

# Convert end time to datetime object
end_time = datetime.strptime(end_time_str, "%d/%m/%Y %I:%M:%S %p")

# Define log file
log_file = os.path.join(os.path.expanduser("~"), "internet_speed_log.xlsx")

# Function to log speed in Excel
def log_speed():
    while datetime.now() < end_time:
        speed = test_speed(url)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Load or create DataFrame
        try:
            df = pd.read_excel(log_file)
        except FileNotFoundError:
            df = pd.DataFrame(columns=["Timestamp", "Download Speed (Mbps)"])

        # Append new entry
        new_entry = pd.DataFrame([[timestamp, speed]], columns=df.columns)
        df = pd.concat([df, new_entry], ignore_index=True)

        # Save to Excel
        df.to_excel(log_file, index=False, engine='openpyxl')
        print(f"Logged: {timestamp} - {speed} Mbps")

        # Wait for the next test
        time.sleep(interval)

    print("End time reached. Stopping the script.")

# Run the script
if __name__ == "__main__":
    log_speed()
