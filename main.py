import pandas as pd
from datetime import datetime
import os

def get_valid_file_path(prompt):
    while True:
        path = input(prompt)
        if os.path.isfile(path):
            return path
        else:
            print(f"Error: The file '{path}' does not exist. Please try again.")

def get_valid_directory_path(prompt):
    while True:
        path = input(prompt)
        if os.path.isdir(os.path.dirname(path)):
            return path
        else:
            print(f"Error: The directory '{os.path.dirname(path)}' does not exist or is incorrect. Please try again.")

def process_signup_roster():


    a = """
██████╗ ███████╗███████╗██╗     ██╗███████╗███████╗
██╔══██╗██╔════╝██╔════╝██║     ██║██╔════╝██╔════╝
██████╔╝█████╗  ███████╗██║     ██║█████╗  █████╗  
██╔══██╗██╔══╝  ╚════██║██║     ██║██╔══╝  ██╔══╝  
██║  ██║███████╗███████║███████╗██║██║     ███████╗
╚═╝  ╚═╝╚══════╝╚══════╝╚══════╝╚═╝╚═╝     ╚══════╝
"""
    # Prompt user for file paths with validation

    print(a)
    base_roster_file = get_valid_file_path("Enter the path to the base roster file (spreadsheet 1): ")
    signup_file = get_valid_file_path("Enter the path to the sign-up form file (spreadsheet 2): ")

    # Load the base roster (spreadsheet 1)
    base_roster_df = pd.read_excel(base_roster_file)

    # Load the sign-up form data (spreadsheet 2)
    signup_df = pd.read_excel(signup_file)

    # Prompt user for the start and end time
    start_time_input = input("Enter the start time (e.g., 8:00): ")
    end_time_input = input("Enter the end time (e.g., 9:00): ")

    # Convert user input to datetime objects
    start_time = datetime.strptime(f"8/13/2024 {start_time_input} PM", "%m/%d/%Y %I:%M %p")
    end_time = datetime.strptime(f"8/13/2024 {end_time_input} PM", "%m/%d/%Y %I:%M %p")

    # Extract the relevant columns ("Name First", "Student Number", and "Timestamp")
    base_roster_ids = base_roster_df[['Name First', 'Student Number']]
    signup_ids = signup_df[['Name First', 'Student Number', 'Timestamp']]

    # Convert the "Timestamp" column to datetime format
    signup_ids['Timestamp'] = pd.to_datetime(signup_ids['Timestamp'], format="%m/%d/%Y %I:%M:%S %p")

    # Find IDs in the base roster that are not in the signup list
    missing_signups = base_roster_ids.loc[~base_roster_ids['Student Number'].isin(signup_ids['Student Number'])]
    missing_signups.loc[:, 'Reason'] = "Did not register"

    # Flag students who signed up outside the specified time frame
    signup_outside_timeframe = signup_ids.loc[
        ~((signup_ids['Timestamp'] >= start_time) & (signup_ids['Timestamp'] <= end_time))
    ]
    signup_outside_timeframe.loc[:, 'Reason'] = signup_outside_timeframe['Timestamp'].dt.strftime('%H:%M')

    # Combine the missing sign-ups with those who signed up outside the time frame
    flagged_students = pd.concat([
        missing_signups,
        signup_outside_timeframe[['Name First', 'Student Number', 'Reason']]
    ]).drop_duplicates()

    # Print the results
    print("Residents who did not sign up or signed up outside the specified time frame:")
    print(flagged_students)

    # Optionally, if you want to find people who signed up but are not on the base roster
    extra_signups = signup_ids.loc[~signup_ids['Student Number'].isin(base_roster_ids['Student Number'])]

    print("\nSign-ups not found on the base roster:")
    print(extra_signups)

    # Prompt user for the output file path with validation
    output_file = get_valid_directory_path("Enter the path to save the results (including the filename, e.g., C:\\path\\to\\output.xlsx): ")

    # Ensure the output file path includes an Excel file extension
    if not output_file.endswith('.xlsx'):
        output_file = os.path.join(output_file, 'flagged_students.xlsx')

    # Write the results to an Excel file
    with pd.ExcelWriter(output_file) as writer:
        flagged_students.to_excel(writer, sheet_name="Flagged Students", index=False)
        extra_signups.to_excel(writer, sheet_name="Extra Sign-ups", index=False)

    print(f"\nResults have been saved to {output_file}")

# Run the function
process_signup_roster()
