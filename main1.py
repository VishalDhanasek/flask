from flask import Flask, jsonify,request
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
from aiohttp import ClientSession
import asyncio
import aiofiles
import aiohttp



app = Flask(__name__)
cors = CORS(app)

# Define the questions and their respective validation patterns/messages
questions_admin = [
    ("Month", "Enter Month", "^\\d{2}$", "Invalid input. Please use DD format."),
    ("Year", "Enter Year: ", "^\\d{4}$", "Invalid input. Please use YYYY format")
]

shifts = ["Morning", "Afternoon", "Night"]


def warning_mail_receiver(employee_id_2,employee_name_2,total_consecutive_days,col_date_employee_1):
    admin_df = pd.read_excel("admin.xlsx")
    year = admin_df.loc[0, "Year"]
    email = "rotavrts@gmail.com"
    password = "rhdd gtal zuso gwnc"
    manager_email = "madheshns57@gmail.com"
    subject = "Shift Swap Notification"
    msg = f'We wanted to notify you that {employee_name_2} (Employee ID: {employee_id_2}) has received a shift swap on {col_date_employee_1}, {year}, which could result in them working for {total_consecutive_days} consecutive days.\n\nRegards,\nVirtusa'

    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = email
    message["To"] = manager_email
    message["Subject"] = subject
    # Add body to email
    message.attach(MIMEText(msg, "plain"))
    # Connect to the server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    # Login to the server
    server.login(email, password)
    # Send email
    server.sendmail(email, manager_email, message.as_string())
    server.quit()

def warning_mail_requester(employee_id_1,total_consecutive_days,employee_name_1,col_date_employee_1):
    admin_df = pd.read_excel("admin.xlsx")
    year = admin_df.loc[0, "Year"]
    email = "rotavrts@gmail.com"
    password = "rhdd gtal zuso gwnc"
    manager_email = "madheshns57@gmail.com"
    subject = "Shift Swap Notification"
    msg = f'We wanted to notify you that {employee_name_1} (Employee ID: {employee_id_1}) has requested a shift swap on {col_date_employee_1}, {year}, which could result in them working for {total_consecutive_days} consecutive days.\n\nRegards,\nVirtusa'

    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = email
    message["To"] = manager_email
    message["Subject"] = subject
    # Add body to email
    message.attach(MIMEText(msg, "plain"))
    # Connect to the server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    # Login to the server
    server.login(email, password)
    # Send email
    server.sendmail(email, manager_email, message.as_string())
    server.quit()

approval=None

async def send_email_and_wait(employee_id_1,employee_id_2,date_employee_1):
    admin_df = pd.read_excel("admin.xlsx")
    month = admin_df.loc[0, "Month"]
    
    # Convert month number to three-letter abbreviation
    month_abbr = calendar.month_abbr[int(month)]

    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"
    # Read Excel file into a DataFrame
    m_df = pd.read_excel("modified_roster.xlsx")
    shift_employee_2 = m_df.loc[m_df['Employee ID'] == int(employee_id_2), col_date_employee_1].iloc[0]
    employee_name_1 = m_df.loc[m_df['Employee ID'] == int(employee_id_1), "Employee Name"].iloc[0]
    if(shift_employee_2!='O'):
        # Initialize variables to store consecutive working days count and previous and next shifts
        consecutive_days_before = 0
        consecutive_days_after = 0
        previous_shift = None
        next_shift = None
        next_shift_index = None

        # Find the index of the selected date
        selected_date_index = m_df.columns.get_loc(month_abbr+" "+date_employee_1)
        print(selected_date_index)
        # Locate the row corresponding to the employee ID
        employee_row = m_df[m_df['Employee ID'] == int(employee_id_1)]
        print("employee_row", employee_row)
        # Get the index of the row for the employee
        employee_index = employee_row.index[0]


        # Iterate over the roster to count consecutive working days before the selected date
        for i in range(selected_date_index - 1,1,-1):
            shift = m_df.iloc[employee_index, i]
            print(shift)
            if shift != 'O' and m_df.iloc[employee_index,i] in ['M', 'A', 'N', 'G']:
                consecutive_days_before += 1
            else:
                break

        # Iterate over the roster to count consecutive working days after the selected date
        for i in range(selected_date_index + 1, len(m_df.columns)-2):
            shift = m_df.iloc[employee_index, i]
            print(shift)
            if shift != 'O':
                consecutive_days_after += 1
            else:
                break

        # Determine the previous shift
        previous_shift_index = selected_date_index - 1
        if previous_shift_index < len(m_df.columns) and m_df.iloc[employee_index, previous_shift_index] in ['M', 'A', 'N', 'G']:
            previous_shift = m_df.iloc[employee_index, previous_shift_index]
        else:
            previous_shift_index=None

        # Determine the next shift
        next_shift_index = selected_date_index + 1
        if next_shift_index < len(m_df.columns) and m_df.iloc[employee_index, next_shift_index] in ['M', 'A', 'N', 'G']:
            next_shift = m_df.iloc[employee_index, next_shift_index]
        else:
            next_shift=None

        # Display the count and shifts to the employee
        total_consecutive_days = consecutive_days_before + 1 + consecutive_days_after
        if(total_consecutive_days>5):
            warning_mail_requester(employee_id_1,total_consecutive_days,employee_name_1,col_date_employee_1)

    timestamp = datetime.now()  # Get current timestamp
    await asyncio.sleep(30)
    approval_status=None

    async with aiohttp.ClientSession() as session:
        async with session.get("http://localhost:5002/approval") as response:
            if response.status == 200:
                json_response = await response.json()
                approval_status = json_response.get("approval")
            else:
                print("Error fetching approval status:", response.status)
            
    # Read Excel file into a DataFrame
    df = pd.read_excel("modified_roster.xlsx")

    emp_df = pd.read_excel("employee_chatbot.xlsx")

    swap_requests_df =pd.read_excel("swap_request_log.xlsx")

    admin_df = pd.read_excel("admin.xlsx")
    month = admin_df.loc[0, "Month"]

    employee_name_1 = emp_df.loc[emp_df['employee_id'] == int(employee_id_1), "employee_name"].iloc[0]
    employee_name_2 = emp_df.loc[emp_df['employee_id'] == int(employee_id_2), "employee_name"].iloc[0]

    # Convert month number to three-letter abbreviation
    month_abbr = calendar.month_abbr[int(month)]

    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"

    shift_employee_1 = df.loc[df['Employee ID'] == int(employee_id_1), col_date_employee_1].iloc[0]
    shift_employee_2 = df.loc[df['Employee ID'] == int(employee_id_2), col_date_employee_1].iloc[0]
    # Record the time when the email was sent
    email_sent_time = time.time()

    # For demonstration purposes, assume a timeout period of 1 minute
    timeout_duration = 1  # 1 minute in seconds

    while time.time() - email_sent_time < timeout_duration:
        # Check for the user's response
        print("approval status --> ",approval_status)
        if approval_status is not None:
            
            if approval_status == "Accepted":
                swap_dates(employee_id_1, employee_id_2, date_employee_1)
                swap_requests_df = swap_requests_df._append({
                    "Requester_ID": employee_id_1,
                    "Requester_Name": employee_name_1,
                    "Recipient_ID": employee_id_2,
                    "Recipient_Name": employee_name_2,
                    "Date": col_date_employee_1,
                    "Shift_Of_Requester" : shift_employee_1,
                    "Requested_Shift" : shift_employee_2,
                    "Status": "Accepted",
                    "Timestamp": timestamp.strftime("%Y-%m-%d %H:%M:%S")  # Add timestamp to the DataFrame
                }, ignore_index=True)
            else:
                swap_requests_df = swap_requests_df._append({
                    "Requester_ID": employee_id_1,
                    "Requester_Name": employee_name_1,
                    "Recipient_ID": employee_id_2,
                    "Recipient_Name": employee_name_2,
                    "Date": col_date_employee_1,
                    "Shift_Of_Requester": shift_employee_1,
                    "Requested_Shift": shift_employee_2,
                    "Status": "Declined",
                    "Timestamp": timestamp.strftime("%Y-%m-%d %H:%M:%S")   # Add timestamp to the DataFrame
                }, ignore_index=True)

                matching_employees = df[(df[col_date_employee_1] == shift_employee_2)]
               
                # Get email addresses corresponding to employee IDs
                requested_Employee_email = emp_df.loc[emp_df['employee_id'] == int(employee_id_1), "Email ID"].iloc[0]
                email = "rotavrts@gmail.com"
                password = "rhdd gtal zuso gwnc"
                rec_email = requested_Employee_email
                subject = "Shift Swap Notification"
                # Get other employees with the same shift
                other_employees = matching_employees[(matching_employees['Employee ID'] != int(employee_id_1)) &
                                                     (matching_employees['Employee ID'] != int(employee_id_2))]
                # Prepare the HTML message
                msg = f'''
                <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <title>Shift Request Declined</title>
                </head>
                <body>
                    <p>Dear {employee_name_1},</p>
                    <p>This is to inform you that your shift request has been declined on {month_abbr} - {date_employee_1} by {employee_name_2} (Employee ID: {employee_id_2}).</p>
                    <p>Other employees with the same shift:</p>
                    <ul>
                '''

                # Iterate over other employees and add them to the HTML message
                for index, row in other_employees.iterrows():
                    employee_id = int(row['Employee ID'])  # Convert to integer
                    msg += f"        <li>Employee ID: {employee_id} | Employee Name: {row['Employee Name']}</li>\n"

                # Complete the HTML message
                msg += '''
                    </ul>
                    <p>You can contact the above mentioned employees for shift swap.</p>
                    <p>Regards,<br>Virtusa</p>
                </body>
                </html>
            '''

                # Create a multipart message
                message = MIMEMultipart()
                message["From"] = email
                message["To"] = rec_email
                message["Subject"] = subject

                # Add body to email
                message.attach(MIMEText(msg, "html"))

                # Connect to the server
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()

                # Login to the server
                server.login(email, password)

                # Send email
                server.sendmail(email, rec_email, message.as_string())
                # print("Sent email to " + rec_email)

                # print("Declined")
                # Save the updated DataFrame to an Excel file
            swap_requests_df.to_excel("swap_request_log.xlsx", index=False)

            break
        await asyncio.sleep(1)


def send_email_with_buttons(sender_email, receiver_email, sender_password, accept_link, decline_link,employee_id_1,employee_id_2,date_employee_1,month):
    # Read Excel file into a DataFrame
    df = pd.read_excel("modified_roster.xlsx")

    emp_df = pd.read_excel("employee_chatbot.xlsx")
    admin_df = pd.read_excel("admin.xlsx")
    year = admin_df.loc[0, "Year"]
    manager_mail="madheshns57@gmail.com"

    # Convert input employee IDs to integers
    employee_id_1 = int(employee_id_1)
    employee_id_2 = int(employee_id_2)

    # Locate rows corresponding to the provided employee IDs
    row_employee_1 = df[df['Employee ID'] == employee_id_1].index
    row_employee_2 = df[df['Employee ID'] == employee_id_2].index

    # Ensure both employee IDs exist in the DataFrame
    if len(row_employee_1) == 0 or len(row_employee_2) == 0:
        print("One or both of the provided employee IDs do not exist.")
        return

    # Convert month number to three-letter abbreviation
    month_abbr = calendar.month_abbr[int(month)]

    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"

    # Check if the provided dates exist in the DataFrame
    if col_date_employee_1 not in df.columns or col_date_employee_2 not in df.columns:
        print("One or both of the provided dates do not exist.")
        return

    employee_name_1 = emp_df.loc[emp_df['employee_id'] == int(employee_id_1), "employee_name"].iloc[0]
    employee_name_2 = emp_df.loc[emp_df['employee_id'] == int(employee_id_2), "employee_name"].iloc[0]

    shift_employee_2 = df.loc[df['Employee ID'] == int(employee_id_2), col_date_employee_1].iloc[0]


    consecutive_days_before = 0
    consecutive_days_after = 0
    previous_shift = None
    next_shift = None
    next_shift_index = None

    # Find the index of the selected date
    selected_date_index = df.columns.get_loc(month_abbr+" "+date_employee_1)
    print(selected_date_index)
    # Locate the row corresponding to the employee ID
    employee_row = df[df['Employee ID'] == int(employee_id_2)]
    print("employee_row", employee_row)
    # Get the index of the row for the employee
    employee_index = employee_row.index[0]


    # Iterate over the roster to count consecutive working days before the selected date
    for i in range(selected_date_index - 1,1,-1):
        shift = df.iloc[employee_index, i]
        print(shift)
        if shift != 'O' and df.iloc[employee_index,i] in ['M', 'A', 'N', 'G']:
            consecutive_days_before += 1
        else:
            break

    # Iterate over the roster to count consecutive working days after the selected date
    for i in range(selected_date_index + 1, len(df.columns)-2):
        shift = df.iloc[employee_index, i]
        print(shift)
        if shift != 'O':
            consecutive_days_after += 1
        else:
            break

    # Determine the previous shift
    previous_shift_index = selected_date_index - 1
    if previous_shift_index < len(df.columns) and df.iloc[employee_index, previous_shift_index] in ['M', 'A', 'N', 'G']:
        previous_shift = df.iloc[employee_index, previous_shift_index]
    else:
        previous_shift_index=None

    # Determine the next shift
    next_shift_index = selected_date_index + 1
    if next_shift_index < len(df.columns) and df.iloc[employee_index, next_shift_index] in ['M', 'A', 'N', 'G']:
        next_shift = df.iloc[employee_index, next_shift_index]
    else:
        next_shift=None

    if next_shift == 'M':
        next_shift = "Morning"
    elif next_shift == 'A':
        next_shift = "Afternoon"
    elif next_shift  == 'N':
        next_shift = "Night"
    elif next_shift == 'G':
        next_shift = "General"
    else:
        next_shift = "Off"

    if previous_shift == 'M':
        previous_shift = "Morning"
    elif previous_shift == 'A':
        previous_shift = "Afternoon"
    elif previous_shift  == 'N':
        previous_shift = "Night"
    elif previous_shift == 'G':
        previous_shift = "General"
    else:
        previous_shift = "Off"

    # Display the count and shifts to the employee
    total_consecutive_days = consecutive_days_before + 1 + consecutive_days_after
    msg=""
    if(total_consecutive_days>5):
        warning_mail_receiver(employee_id_2,employee_name_2,total_consecutive_days,col_date_employee_1)
        msg = f"""
        <html><body><p>You will be working consecutively for {total_consecutive_days} days if you accept this request.</p>
          <p>Previous day shift: {previous_shift if previous_shift else 'None'} </p>
          <p>Next day shift: {next_shift if next_shift else 'None'} </p>
        </body>
        </html>
        """

    # Email content
    message = MIMEMultipart("alternative")
    message["Subject"] = "Shift Swap Notification"
    message["From"] = sender_email
    message["To"] = receiver_email
    

    # HTML content with accept and decline buttons
    html_content = f"""
    <html>
      <body>
        <p>Dear {employee_name_2}</p>
        <p>You have a shift swap request on {col_date_employee_1}, {year}, by {employee_name_1} (Employee ID: {employee_id_1}).</p>
        <p> {msg} </p>
        <p> Kindly respond:</p>
        <a href="{accept_link}"><button onclick="this.disabled=true" style="background-color: #4CAF50; color: white; padding: 15px 32px; text-align: center; display: inline-block; font-size: 16px;">Accept</button></a>
        <a href="{decline_link}"><button onclick="this.disabled=true" style="background-color: #f44336; color: white; padding: 15px 32px; text-align: center; display: inline-block; font-size: 16px;">Decline</button></a>
      </body>
    </html>
"""

    # Attach HTML content to the email
    message.attach(MIMEText(html_content, "html"))

    # Connect to SMTP server and send email
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        



async def swap_pre_process(employee_id_1,employee_id_2,date_employee_1):
    sender_email = "rotavrts@gmail.com"
    sender_password = "rhdd gtal zuso gwnc"
    accept_link = "http://localhost:5002/accept"
    decline_link = "http://localhost:5002/decline"
    # Read Excel file into a DataFrame
    df = pd.read_excel("modified_roster.xlsx")

    emp_df = pd.read_excel("employee_chatbot.xlsx")

    admin_df = pd.read_excel("admin.xlsx")
    month = admin_df.loc[0, "Month"]

    # Convert input employee IDs to integers
    employee_id_1 = int(employee_id_1)
    employee_id_2 = int(employee_id_2)

    # Locate rows corresponding to the provided employee IDs
    row_employee_1 = df[df['Employee ID'] == employee_id_1].index
    row_employee_2 = df[df['Employee ID'] == employee_id_2].index

    # Ensure both employee IDs exist in the DataFrame
    if len(row_employee_1) == 0 or len(row_employee_2) == 0:
        print("One or both of the provided employee IDs do not exist.")
        return

    # Convert month number to three-letter abbreviation
    month_abbr = calendar.month_abbr[int(month)]

    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"
    print("col_date_employee_1",col_date_employee_1)
    print("col_date_employee_2",col_date_employee_2)

    # Check if the provided dates exist in the DataFrame
    if col_date_employee_1 not in df.columns or col_date_employee_2 not in df.columns:
        print("One or both of the provided dates do not exist.")
        return
    

    # Get email addresses corresponding to employee IDs
    receiver_email = emp_df.loc[emp_df['employee_id'] == int(employee_id_2), "Email ID"].iloc[0]
    employee_name_1 = emp_df.loc[emp_df['employee_id'] == int(employee_id_1), "employee_name"].iloc[0]
    employee_name_2 = emp_df.loc[emp_df['employee_id'] == int(employee_id_2), "employee_name"].iloc[0]

    send_email_with_buttons(sender_email, receiver_email,sender_password, accept_link, decline_link,employee_id_1,employee_id_2,date_employee_1,month)

    # Start a new thread to send email and wait for response
    await (send_email_and_wait(employee_id_1, employee_id_2, date_employee_1))
    # Continue running the chatbot without waiting for the email response
    return{"Message":"Email sent. Waiting for response in the background."}



# Function to save responses to Excel file
def save_to_excel_admin(responses):
    try:
        # Read existing data from Excel file if it exists
        existing_data = pd.read_excel("admin.xlsx")
        # Append new responses to existing data
        morning_list,afternoon_list,night_list = [],[],[]
        for key,value in responses.items():
            if 'morning' in key.lower():
                morning_list.append(value)
            elif 'afternoon' in key.lower():
                afternoon_list.append(value)
            elif 'night' in key.lower():
                night_list.append(value)
        new_response = []
        for idx,data in enumerate(morning_list):
            new_response.append({"Month":responses['month'],"Year":responses['year'],"Morning":morning_list[idx],"Afternoon":afternoon_list[idx],"Night":night_list[idx],"General": 0,"Project Code":"AAS2"})
        # new_data = pd.concat([existing_data, pd.DataFrame(new_response)], ignore_index=True)
        # Write the combined data to the Excel file
        new_data = pd.DataFrame(new_response)
        new_data.to_excel("admin.xlsx", index=False)
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame with the new responses and write it to the Excel file
        df = pd.DataFrame(responses)
        df.to_excel("admin.xlsx", index=False)
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
    existing_data = pd.read_excel("employee_chatbot.xlsx",
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
        existing_data = pd.read_excel("employee_chatbot.xlsx",
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
           preferred_dates = [int(date) for date in preferred_dates if date != ""]

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
        existing_data.to_excel("employee_chatbot.xlsx", index=False,
                               engine='openpyxl')
        return{"message":"Thank you! Your responses have been saved."}
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame with the new responses and write it to the Excel file
        df = pd.DataFrame(responses, index=[0])
        df.to_excel("employee_chatbot.xlsx", index=False,
                    engine='openpyxl')
        return{"message":"Thank you! Your responses have been saved."}

# Define the output path where the modified Excel file should be saved
output_path = "modified_roster.xlsx"

def swap_dates(employee_id_1, employee_id_2, date_employee_1):
    # Read Excel file into a DataFrame
    df = pd.read_excel("modified_roster.xlsx")
    df['Employee ID'] = pd.to_numeric(df['Employee ID'], errors='coerce')
    admin_df = pd.read_excel("admin.xlsx")
    emp_df=pd.read_excel("employee_chatbot.xlsx")
    month = admin_df.loc[0, "Month"]
    year = admin_df.loc[0,"Year"]
    sender_email = emp_df.loc[emp_df['employee_id'] == int(employee_id_1), "Email ID"].iloc[0]
    email_2=emp_df.loc[emp_df['employee_id'] == int(employee_id_2), "Email ID"].iloc[0]
    employee_name_1 =emp_df.loc[emp_df['employee_id'] == int(employee_id_1), "employee_name"].iloc[0]
    employee_name_2 = emp_df.loc[emp_df['employee_id'] == int(employee_id_2), "employee_name"].iloc[0]
    # Convert input employee IDs to integers
    employee_id_1 = int(employee_id_1)
    print("--",employee_id_1, type(employee_id_1), "--", df["Employee ID"])
    employee_id_2 = int(employee_id_2)
    print("--",employee_id_2, type(employee_id_2), "--", df["Employee ID"])

    # Locate rows corresponding to the provided employee IDs
    row_employee_1 = df[df['Employee ID'] == employee_id_1].index
    row_employee_2 = df[df['Employee ID'] == employee_id_2].index

    # Ensure both employee IDs exist in the DataFrame
    if len(row_employee_1) == 0 or len(row_employee_2) == 0:
        print("One or both of the provided employee IDs do not exist.")
        return {"message":"One or both of the provided employee IDs do not exist."}

    month_abbr = calendar.month_abbr[int(month)]
    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"

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
    send_email_to_cab_service_about_requester(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year,sender_email)
    send_email_to_cab_service_about_receiver(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year,sender_email)
    send_email_to_manager(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year)
    send_mail_to_requester(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year,sender_email)
    send_confirmation_mail_to_receiver(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year,email_2)
    return {"message":f"Values swapped successfully. Modified data saved to {output_path}"}


def send_confirmation_mail_to_receiver(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year,email_2):
    df = pd.read_excel("modified_roster.xlsx")
    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"
    shift_employee_1 = df.loc[df['Employee ID'] == int(employee_id_1), col_date_employee_1].iloc[0]
    shift_employee_2 = df.loc[df['Employee ID'] == int(employee_id_2), col_date_employee_1].iloc[0]
    if shift_employee_2 == 'M':
        shift_employee_2 = "Morning"
    elif shift_employee_2 == 'A':
        shift_employee_2 = "Afternoon"
    elif shift_employee_2 == 'N':
        shift_employee_2 = "Night"
    elif shift_employee_2 == 'G':
        shift_employee_2 = "General"
    else:
        shift_employee_2 = "Off"
    email = "rotavrts@gmail.com"
    password = "rhdd gtal zuso gwnc"

    subject = "Shift Swap Notification"
    msg = f'Dear {employee_name_2}, \n\nThis mail serves as a confirmation that you have accepted to a shift swap with {employee_name_1} (Employee ID:{employee_id_1}). Your updated schedule will be {shift_employee_2} shift on {col_date_employee_1},{year}. \n\nRegards,\nVirtusa'

    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = email
    message["To"] = email_2
    message["Subject"] = subject
    # Add body to email
    message.attach(MIMEText(msg, "plain"))
    # Connect to the server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    # Login to the server
    server.login(email, password)
    # Send email
    server.sendmail(email,email_2, message.as_string())
    server.quit()

def consecutive_days():
    df = pd.read_excel("modified_roster.xlsx")
    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"
    shift_employee_2 = df.loc[df['Employee ID'] == int(employee_id_2), col_date_employee_1].iloc[0]
    consecutive_days_before = 0
    consecutive_days_after = 0
    previous_shift = None
    next_shift = None
    next_shift_index = None

    # Find the index of the selected date
    selected_date_index = df.columns.get_loc(month_abbr+" "+date_employee_1)
    print(selected_date_index)
    # Locate the row corresponding to the employee ID
    employee_row = df[df['Employee ID'] == int(employee_id_2)]
    print("employee_row", employee_row)
    # Get the index of the row for the employee
    employee_index = employee_row.index[0]


    # Iterate over the roster to count consecutive working days before the selected date
    for i in range(selected_date_index - 1,1,-1):
        shift = df.iloc[employee_index, i]
        print(shift)
        if shift != 'O' and df.iloc[employee_index,i] in ['M', 'A', 'N', 'G']:
            consecutive_days_before += 1
        else:
            break

    # Iterate over the roster to count consecutive working days after the selected date
    for i in range(selected_date_index + 1, len(df.columns)-2):
        shift = df.iloc[employee_index, i]
        print(shift)
        if shift != 'O':
            consecutive_days_after += 1
        else:
            break

    # Determine the previous shift
    previous_shift_index = selected_date_index - 1
    if previous_shift_index < len(df.columns) and df.iloc[employee_index, previous_shift_index] in ['M', 'A', 'N', 'G']:
        previous_shift = df.iloc[employee_index, previous_shift_index]
    else:
        previous_shift_index=None

    # Determine the next shift
    next_shift_index = selected_date_index + 1
    if next_shift_index < len(df.columns) and df.iloc[employee_index, next_shift_index] in ['M', 'A', 'N', 'G']:
        next_shift = df.iloc[employee_index, next_shift_index]
    else:
        next_shift=None

    if next_shift == 'M':
        next_shift = "Morning"
    elif next_shift == 'A':
        next_shift = "Afternoon"
    elif next_shift  == 'N':
        next_shift = "Night"
    elif next_shift == 'G':
        next_shift = "General"
    else:
        next_shift = "Off"

    if previous_shift == 'M':
        previous_shift = "Morning"
    elif previous_shift == 'A':
        previous_shift = "Afternoon"
    elif previous_shift  == 'N':
        previous_shift = "Night"
    elif previous_shift == 'G':
        previous_shift = "General"
    else:
        previous_shift = "Off"

    # Display the count and shifts to the employee
    total_consecutive_days = consecutive_days_before + 1 + consecutive_days_after
    msg=""
    if(total_consecutive_days>5):
        msg = f"You will be working consecutively for {total_consecutive_days} days if you accept this request.\n" \
          f"Your previous shift is {previous_shift if previous_shift else 'None'}\n" \
          f"Your next shift will be {next_shift if next_shift else 'None'}"

def send_email_to_manager(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year):
    df = pd.read_excel("modified_roster.xlsx")
    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"
    shift_employee_2 = df.loc[df['Employee ID'] == int(employee_id_2), col_date_employee_1].iloc[0]
    if shift_employee_2 == 'M':
        shift_employee_2 = "Morning"
    elif shift_employee_2 == 'A':
        shift_employee_2 = "Afternoon"
    elif shift_employee_2 == 'N':
        shift_employee_2 = "Night"
    elif shift_employee_2 == 'G':
        shift_employee_2 = "General"
    else:
        shift_employee_2 = "Off"

    
    email = "rotavrts@gmail.com"
    password = "rhdd gtal zuso gwnc"
    manager_email = "madheshns57@gmail.com"
    subject = "Shift Swap Notification"
    msg = f'Dear Madhesh, \n\nThis is to inform you that {employee_name_1} (Employee ID: {employee_id_1}) has swapped the shift with {employee_name_2} (Employee ID: {employee_id_2}) on {col_date_employee_1}, {year}.\n\nRegards,\nVirtusa'

    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = email
    message["To"] = manager_email
    message["Subject"] = subject
    # Add body to email
    message.attach(MIMEText(msg, "plain"))
    # Connect to the server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    # Login to the server
    server.login(email, password)
    # Send email
    server.sendmail(email, manager_email, message.as_string())
    server.quit()
def send_mail_to_requester(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year,sender_email):
    df = pd.read_excel("modified_roster.xlsx")
    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"
    shift_employee_1 = df.loc[df['Employee ID'] == int(employee_id_1), col_date_employee_1].iloc[0]
    if shift_employee_1 == 'M':
        shift_employee_1 = "Morning"
    elif shift_employee_1 == 'A':
        shift_employee_1 = "Afternoon"
    elif shift_employee_1 == 'N':
        shift_employee_1 = "Night"
    elif shift_employee_1 == 'G':
        shift_employee_1 = "General"
    else:
        shift_employee_1 = "Off"
    email = "rotavrts@gmail.com"
    password = "rhdd gtal zuso gwnc"
    subject = "Shift Swap Notification"
    msg = f'Dear {employee_name_1}, \n\nThis is to inform you that the {shift_employee_1} shift you have requested on {month_abbr} {date_employee_1} {year}, has been approved by {employee_name_2} (Employee id: {employee_id_2}).\n\nRegards,\nVirtusa'

    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = email
    message["To"] = sender_email
    message["Subject"] = subject
    # Add body to email
    message.attach(MIMEText(msg, "plain"))
    # Connect to the server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    # Login to the server
    server.login(email, password)
    # Send email
    server.sendmail(email, sender_email, message.as_string())
    server.quit()

def send_email_to_cab_service_about_receiver(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year,sender_email):
    df = pd.read_excel("modified_roster.xlsx")
    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"
    shift_employee_2 = df.loc[df['Employee ID'] == int(employee_id_2), col_date_employee_1].iloc[0]
    if shift_employee_2=='M':
        shift_employee_2="Morning"
    elif shift_employee_2=='A':
        shift_employee_2="Afternoon"
    elif shift_employee_2 =='N':
        shift_employee_2="Night"
    elif shift_employee_2 =='G':
        shift_employee_2="General"
    else:
        shift_employee_2="Off"
    email = "rotavrts@gmail.com"
    password = "rhdd gtal zuso gwnc"
    cab_email="cabvrts@gmail.com"
    subject = "Shift Swap Notification for Cab Service"
    if shift_employee_2=="Off":
        msg=f'Dear Cab Service Provider, \n\nThis is to inform you that {employee_name_2} with Employee ID {employee_id_2} will be on leave on {month_abbr} {date_employee_1},{year}. Kindly cancel the cab facility for the above mentioned day. \n\nRegards,\nVirtusa'
    else:
        msg = f'Dear Cab Service Provider, \n\nThis is to inform you that the shift has been updated to {shift_employee_2} shift for the employee {employee_name_2} (Employee ID: {employee_id_2}) on {month_abbr} {date_employee_1},{year}. Kindly make the necessary cab arrangements accordingly.\n\nRegards,\nVirtusa'

    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = email
    message["To"] = cab_email
    message["Subject"] = subject
    # Add body to email
    message.attach(MIMEText(msg, "plain"))
    # Connect to the server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    # Login to the server
    server.login(email, password)
    # Send email
    server.sendmail(email, cab_email, message.as_string())
    server.quit()

def send_email_to_cab_service_about_requester(month_abbr,date_employee_1,employee_name_1,employee_id_1,employee_id_2,employee_name_2,year,sender_email):
    df = pd.read_excel("modified_roster.xlsx")
    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"
    shift_employee_1 = df.loc[df['Employee ID'] == int(employee_id_1), col_date_employee_1].iloc[0]
    shift_employee_2 = df.loc[df['Employee ID'] == int(employee_id_2), col_date_employee_1].iloc[0]
    if shift_employee_1=='M':
        shift_employee_1="Morning"
    elif shift_employee_1=='A':
        shift_employee_1="Afternoon"
    elif shift_employee_1 =='N':
        shift_employee_1="Night"
    elif shift_employee_1 =='G':
        shift_employee_1="General"
    else:
        shift_employee_1="Off"
    email = "rotavrts@gmail.com"
    password = "rhdd gtal zuso gwnc"
    cab_email="cabvrts@gmail.com"
    subject = "Shift Swap Notification for Cab Service"
    if shift_employee_1=="Off":
        msg=f'Dear Cab Service Provider, \n\nThis is to inform you that {employee_name_1} with Employee ID {employee_id_1} will be on leave on {month_abbr} {date_employee_1},{year}. Kindly cancel the cab facility for the above mentioned day. \n\nRegards,\nVirtusa'
    else:
        msg = f'Dear Cab Service Provider, \n\nThis is to inform you that the shift has been updated to {shift_employee_1} shift for the employee {employee_name_1} (Employee ID: {employee_id_1}) on {month_abbr} {date_employee_1},{year}. Kindly make the necessary cab arrangements accordingly.\n\nRegards,\nVirtusa'

    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = email
    message["To"] = cab_email
    message["Subject"] = subject
    # Add body to email
    message.attach(MIMEText(msg, "plain"))
    # Connect to the server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    # Login to the server
    server.login(email, password)
    # Send email
    server.sendmail(email, cab_email, message.as_string())
    server.quit()

@app.route("/accept", methods=["GET"])
def accept_request():
    global approval
    approval = "Accepted"
    print(approval)
    return "Response saved: Accepted"

@app.route("/decline", methods=["GET"])
def decline_request():
    global approval
    approval = "Declined"
    print(approval)
    return "Response saved: Declined"

@app.route("/approval", methods=["GET"])
def get_approval_status():
    return {"approval": approval}

@app.route('/main', methods=['POST'])
def main():
    item = request.get_json()
    print("item -- ",item)
    if 'employee_id' in item:
        if 'Planned_Leave_1' in item:
            return save_to_excel_employees(item)
        else:
            return asyncio.run(swap_pre_process(item['employee_id'], item['swap_id'], item['swap_date_1']))
            # return jsonify({"Message": "Email sent. Waiting for response in the background."})
            
    else:
        return save_to_excel_admin(item)

@app.route('/employee_login', methods=['POST'])
def employee_login():
    item = request.get_json()
    print("item -- ",item)
    employee_id = item['id']
    print("employee_id -- ",employee_id)
    entered_password = item['password']
    print("entered_password -- ",entered_password)

    try:
        # Load employee credentials from Excel file
        credentials_df = pd.read_excel("credentials.xlsx")

        # Check if the employee ID exists in the DataFrame
        if str(employee_id) in credentials_df['Id'].astype(str).values:
            # Retrieve the corresponding password as string
            matching_rows = credentials_df[credentials_df['Id'].astype(str) == str(employee_id)]

            if not matching_rows.empty:
                password = str(matching_rows.iloc[0]['Password'])

                # Compare entered password with the stored password
                if entered_password == password:
                    print("Password matching")
                    return {"Message":"true"}
                else:
                    print("Invalid password.")
                    return {"Message":"false"}
            else:
                print("No matching rows found.")
                return {"Message":"false"}
        else:
            print("Employee ID not found.")
            return {"Message":"false"}
    except FileNotFoundError:
        print("Credentials file not found.")
        return {"Message":"false"}
    except Exception as e:
        print("Error:", e)
        return {"Message":"false"}
    
@app.route('/employee_available_for_swap_shift',methods=['POST'])
def employee_available_for_swap_shift():
    item = request.get_json()
    print("item -- ",item)
    date_employee_1 = item['swap_date_1'] 
    shift_to_swap = item['shift_to_swap']

    df = pd.read_excel("modified_roster.xlsx")
    admin_df = pd.read_excel("admin.xlsx")
    month = admin_df.loc[0, "Month"]
    print("month -- ", month)

    # Convert month number to three-letter abbreviation
    month_abbr = calendar.month_abbr[int(month)]

    # Get column indices corresponding to the provided dates
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    col_date_employee_2 = f"{month_abbr} {date_employee_1}"

    # Check if the provided dates exist in the DataFrame
    if col_date_employee_1 not in df.columns or col_date_employee_2 not in df.columns:
        print("One or both of the provided dates do not exist.")
        return{"message":"One or both of the provided dates do not exist."}


    # Filter the DataFrame based on the provided date and shift
    matching_employees = df[(df[col_date_employee_1] == shift_to_swap)]
    print("matching Employees -- ", matching_employees)
    # Print the names and IDs of matching employees
    if not matching_employees.empty:
        employee_list = []
        print("Employees with the specified shift on the given date:")
        for index, row in matching_employees.iterrows():
            employee_id = int(row['Employee ID'])
            employee_list.append(f"{employee_id} : {row['Employee Name']}")
        return({"message":'\n\n'.join(employee_list)})
    else:
        return{"message":"null"}

@app.route('/check_consecutive_days',methods=['POST'])
def check_consecutive_days():
    item = request.get_json()
    print("item -- ",item)
    date_employee_1 = item['swap_date_1'] 
    shift_to_swap = item['shift_to_swap']
    employee_id_1 = item['employee_id']
    flag=False
    df = pd.read_excel("modified_roster.xlsx")
    admin_df = pd.read_excel("admin.xlsx")
    month = admin_df.loc[0, "Month"]
    month_abbr = calendar.month_abbr[int(month)]
    col_date_employee_1 = f"{month_abbr} {date_employee_1}"
    employee_name_1 = df.loc[df['Employee ID'] == int(employee_id_1), "Employee Name"].iloc[0]


    if(shift_to_swap!='O'):
        # Initialize variables to store consecutive working days count and previous and next shifts
        consecutive_days_before = 0
        consecutive_days_after = 0
        previous_shift = None
        next_shift = None
        next_shift_index = None

        # Find the index of the selected date
        selected_date_index = df.columns.get_loc(month_abbr+" "+date_employee_1)
        print(selected_date_index)
        # Locate the row corresponding to the employee ID
        employee_row = df[df['Employee ID'] == int(employee_id_1)]
        print("employee_row", employee_row)
        # Get the index of the row for the employee
        employee_index = employee_row.index[0]


        # Iterate over the roster to count consecutive working days before the selected date
        for i in range(selected_date_index - 1,1,-1):
            shift = df.iloc[employee_index, i]
            print(shift)
            if shift != 'O' and df.iloc[employee_index,i] in ['M', 'A', 'N', 'G']:
                consecutive_days_before += 1
            else:
                break

        # Iterate over the roster to count consecutive working days after the selected date
        for i in range(selected_date_index + 1, len(df.columns)-2):
            shift = df.iloc[employee_index, i]
            print(shift)
            if shift != 'O':
                consecutive_days_after += 1
            else:
                break

        # Determine the previous shift
        previous_shift_index = selected_date_index - 1
        if previous_shift_index < len(df.columns) and df.iloc[employee_index, previous_shift_index] in ['M', 'A', 'N', 'G']:
            previous_shift = df.iloc[employee_index, previous_shift_index]
        else:
            previous_shift_index=None

        # Determine the next shift
        next_shift_index = selected_date_index + 1
        if next_shift_index < len(df.columns) and df.iloc[employee_index, next_shift_index] in ['M', 'A', 'N', 'G']:
            next_shift = df.iloc[employee_index, next_shift_index]
        else:
            next_shift=None

        # Display the count and shifts to the employee
        total_consecutive_days = consecutive_days_before + 1 + consecutive_days_after
        if(total_consecutive_days>5):
            print({"message": f"You will be working consecutively for {total_consecutive_days} days."})
            return {
    "message": f"You will be working consecutively for {total_consecutive_days} days.\n"
               f"Previous shift: {previous_shift if previous_shift else 'None'}\n                    "
               f"\n"
               f"Next shift: {next_shift if next_shift else 'None'}"
}

        return({"message": "null"})
    else:
        return({"message": "null"})


if __name__ == '__main__':
    app.run(debug=True, port=5002, host='0.0.0.0')