import pandas as pd
import logging
from pathlib import Path
import sys

logging.basicConfig(level=logging.DEBUG, filename="output.txt", filemode="w")

def analyze_file(file_path):
    """Analyzes the employee work schedule file and prints the names of employees who meet the criteria specified in the question."""

    try:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(file_path)

        if not df.empty:
            # Convert the 'Time' and 'Time Out' columns to datetime type
            df['Time'] = pd.to_datetime(df['Time'])
            df['Time Out'] = pd.to_datetime(df['Time Out'])

            # Sort the DataFrame by 'Employee Name' and 'Time' for consecutive day analysis
            df = df.sort_values(by=['Employee Name', 'Time']).reset_index(drop=True)

            # Iterate through the DataFrame to analyze employee data
            for idx, row in df.iterrows():
                employee = row['Employee Name']
                date = row['Time']

                # Convert the 'Timecard Hours (as Time)' to hours and minutes
                hours_minutes_str = str(row['Timecard Hours (as Time)'])
                if ':' in hours_minutes_str:
                    hours_minutes_list = [int(x) for x in hours_minutes_str.split(':')]
                    hours_worked = hours_minutes_list[0] + hours_minutes_list[1] / 60
                else:
                    # Handle numeric format (assuming it represents hours)
                    hours_worked = float(hours_minutes_str)

                # Check for employees who have worked for 7 consecutive days
                consecutive_days = df[(df['Employee Name'] == employee) & (df['Time'] >= date) & (df['Time'] <= date + pd.DateOffset(days=6))]
                if len(consecutive_days) == 7:
                    logging.info("%s has worked for 7 consecutive days starting from %s", employee, date)

                # Check for employees with less than 10 hours between shifts but greater than 1 hour
                next_shift_candidates = df[(df['Employee Name'] == employee) & (df['Time'] > date)]
                if not next_shift_candidates.empty:
                    next_shift = next_shift_candidates.iloc[0, :]
                    time_between_shifts = (next_shift['Time'] - date).total_seconds() / 3600
                    if 1 < time_between_shifts < 10:
                        logging.info("%s has less than 10 hours between shifts but greater than 1 hour (%.2f hours) starting from %s", employee, time_between_shifts, date)

                # Check for employees who have worked for more than 14 hours in a single shift
                if hours_worked > 14:
                    logging.info("%s has worked for more than 14 hours in a single shift on %s", employee, date)

    except KeyError as e:
        # Handle the error if the 'Time' or 'Time Out' columns do not exist
        logging.error("Error: The 'Time' or 'Time Out' columns are not present in the DataFrame. %s", e)
        logging.info("Column Names: %s", df.columns)

if __name__ == "__main__":
    file_path = Path(r"C:\Users\VATSAL\Documents\Vatsal_py\Assignment_Timecard.xlsx")
    analyze_file(file_path)
