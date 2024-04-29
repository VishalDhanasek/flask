from flask import Flask,request
from flask_cors import CORS
import pandas as pd
import re
from datetime import datetime
import calendar
import os
import openpyxl
from secrets import compare_digest
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
import time
import threading



app = Flask(__name__)
cors = CORS(app)

# Define the questions and their respective validation patterns/messages
questions_admin = [
    ("Month", "Enter Month", "^\\d{2}$", "Invalid input. Please use DD format."),
    ("Year", "Enter Year: ", "^\\d{4}$", "Invalid input. Please use YYYY format")
]

shifts = ["Morning", "Afternoon", "Night"]



# Function to save responses to Excel file
def save_to_excel_admin(responses):
    try:
        # Read existing data from Excel file if it exists
        existing_data = pd.read_excel("C:\\Users\\veeru\\Downloads\\roster\\admin.xlsx")
        # Append new responses to existing data
        new_data = pd.concat([existing_data, responses], ignore_index=True)
        # Write the combined data to the Excel file
        new_data.to_excel("C:\\Users\\veeru\\Downloads\\roster\\admin.xlsx", index=False)
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame with the new responses and write it to the Excel file
        df = pd.DataFrame(responses)
        df.to_excel("C:\\Users\\veeru\\Downloads\\roster\\admin.xlsx", index=False)
    return {"message":'successfully saved'}


# Function to validate if the entered date is valid within a month
def is_valid_date(day):
    try:
        day = int(day)
        year = datetime.now().year
        month = datetime.now().month
        num_days = calendar.monthrange(year, month)[1]
        return 1 <= day <= num_days
    except ValueError:
        return False




# Print the available dates
def available_dates():
    leave_request = {}
    existing_data = pd.read_excel("C:\\Users\\veeru\\Downloads\\roster\\employee_chatbot.xlsx",
                                  engine='openpyxl')
    dates = existing_data[['Planned_Leave_1', 'Planned_Leave_2']].stack().dropna().astype(int)

    # Get the current year and month
    current_year = datetime.now().year
    current_month = datetime.now().month

    # Get the number of days in the current month
    num_days = calendar.monthrange(current_year, current_month)[1]

    # Generate all dates of the current month
    dates_of_current_month = [datetime(current_year, current_month, day) for day in range(1, num_days + 1)]

    print("Dates Available")
    date_list = []
    for date in dates_of_current_month:
        day_of_month = date.day
        leave_request[day_of_month] = 0
        for iterator in dates:
            if iterator == day_of_month:
                leave_request[day_of_month] += 1


        if leave_request[day_of_month] <= 1:
            # print(f"{date},")
            print(f"{day_of_month}", end="   ")
            date_list.append(day_of_month)

    return {"available_dates":date_list}


# Function to save responses to Excel file
def save_to_excel_employees(responses):
    leave_request = {}
    try:
        # Read existing data from Excel file if it exists
        existing_data = pd.read_excel("C:\\Users\\veeru\\Downloads\\roster\\employee_chatbot.xlsx",
                                      engine='openpyxl')

        # Convert existing employee_id values to integers for comparison
        existing_data['employee_id'] = existing_data['employee_id'].astype(int)

        # Convert responses employee_id to integer for comparison
        responses['employee_id'] = int(responses['employee_id'])

        # Check if the preferred dates are already taken
        preferred_dates = [responses['Planned_Leave_1'], responses['Planned_Leave_2']]
        if all(date is None for date in preferred_dates):
            print("Both preferred dates are None")
        elif any(date is None for date in preferred_dates):
            print("At least one preferred date is None")
        else:
            # Convert string representations to actual values using eval
            # Convert string representations to actual integers
            preferred_dates = [int(date) for date in preferred_dates]

        # print(preferred_dates, preferred_dates[0], type(preferred_dates[0]), existing_data['Preferred_Date_1'],
        # type(existing_data['Preferred_Date_1']))
        dates = existing_data[['Planned_Leave_1', 'Planned_Leave_2']].stack().dropna().astype(int)

        # Print all dates of the current month
        for date in preferred_dates:
            for iterator in dates:
                if iterator == date:
                    if iterator not in leave_request:
                        leave_request[iterator] = 0

                    leave_request[iterator] += 1

                    if leave_request[iterator] >= 2:
                        print(f"Preferred date {date} is already taken. Please choose another date.")
                        return available_dates()

        if responses['employee_id'] in existing_data['employee_id'].values:
            existing_data.loc[existing_data['employee_id'] == responses['employee_id'], 'Planned_Leave_1'] = responses[
                'Planned_Leave_1']
            existing_data.loc[existing_data['employee_id'] == responses['employee_id'], 'Planned_Leave_2'] = responses[
                'Planned_Leave_2']
            print("Employee data updated.")
        else:
            # Append new responses to existing data
            existing_data = pd.concat([existing_data, pd.DataFrame(responses, index=[0])], ignore_index=True)
            print("New employee added.")

        # Write the combined data to the Excel file
        existing_data.to_excel("C:\\Users\\veeru\\Downloads\\roster\\employee_chatbot.xlsx", index=False,
                               engine='openpyxl')
        return{"message":"Thank you! Your responses have been saved."}
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame with the new responses and write it to the Excel file
        df = pd.DataFrame(responses, index=[0])
        df.to_excel("C:\\Users\\veeru\\Downloads\\roster\\employee_chatbot.xlsx", index=False,
                    engine='openpyxl')
        return{"message":"Thank you! Your responses have been saved."}



# Define the output path where the modified Excel file should be saved
output_path = "C:\\Users\\veeru\\Downloads\\roster\\modified_roster.xlsx"

def swap_dates(employee_id_1, employee_id_2, date_employee_1):
    # Read Excel file into a DataFrame
    df = pd.read_excel("C:\\Users\\veeru\\Downloads\\roster\\modified_roster.xlsx")
    df['Employee ID'] = pd.to_numeric(df['Employee ID'], errors='coerce')


    # Convert input employee IDs to integers
    employee_id_1 = float(employee_id_1)
    print("--",employee_id_1, type(employee_id_1), "--", df["Employee ID"])
    employee_id_2 = float(employee_id_2)
    print("--",employee_id_2, type(employee_id_2), "--", df["Employee ID"])

    # Locate rows corresponding to the provided employee IDs
    row_employee_1 = df[df['Employee ID'] == employee_id_1].index
    row_employee_2 = df[df['Employee ID'] == employee_id_2].index

    # Ensure both employee IDs exist in the DataFrame
    if len(row_employee_1) == 0 or len(row_employee_2) == 0:
        print("One or both of the provided employee IDs do not exist.")
        return {"message":"One or both of the provided employee IDs do not exist."}

    # Get column indices corresponding to the provided dates
    col_date_employee_1 = "Mar " + str(date_employee_1)
    col_date_employee_2 = "Mar " + str(date_employee_1)

    # Check if the provided dates exist in the DataFrame
    if col_date_employee_1 not in df.columns or col_date_employee_2 not in df.columns:
        print("One or both of the provided dates do not exist.")
        return {"message":"One or both of the provided dates do not exist."}

    # Swap values for the specified dates
    temp_value = df.at[row_employee_1[0], col_date_employee_1]
    df.at[row_employee_1[0], col_date_employee_1] = df.at[row_employee_2[0], col_date_employee_2]
    df.at[row_employee_2[0], col_date_employee_2] = temp_value

    # Write the modified DataFrame to the specified output Excel file
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    return {"message":f"Values swapped successfully. Modified data saved to {output_path}"}

@app.route('/main', methods=['POST'])
def main():
    item = request.get_json()
    print("item -- ",item)
    if 'employee_id' in item:
        if 'Planned_Leave_1' in item:
            return save_to_excel_employees(item)
        else:
            return swap_dates(item['employee_id'],item['swap_id'],item['swap_date_1'])
    else:
        return save_to_excel_admin(item)
    

if __name__ == '__main__':
    app.run(debug=True, port=5002, host='0.0.0.0')