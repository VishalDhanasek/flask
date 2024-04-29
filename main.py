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

approval=None



def send_email_and_wait(employee_id_1,employee_id_2,date_employee_1):
    time.sleep(30)
    approval_status=None
    response = requests.get("https://b2411a61-a517-4ae4-9b30-5cbd4e3a793d-00-xlmxsla4v0nm.worf.replit.dev:5000/approval")
    if response.status_code == 200:
        approval_status= response.json().get("approval")
    else:
        print("Error fetching approval status:", response.status_code)
        
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
                    "Status": "Accepted"
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
                    "Status": "Declined"
                }, ignore_index=True)


                # Get email addresses corresponding to employee IDs
                requested_Employee_email = emp_df.loc[emp_df['employee_id'] == int(employee_id_1), "Email ID"].iloc[0]
                email = "vishal.d2019cse@sece.ac.in"
                password = "Dkvm@2016"
                rec_email = requested_Employee_email
                subject = "Shift Swap Notification"
                msg = f'Dear {employee_name_1}, \n\nThis is to inform you that your shift request has been declined on {month_abbr} - {date_employee_1} by {employee_name_2} with \nEmployee ID: {employee_id_2}\n\nRegards,\nVirtusa'

                # Create a multipart message
                message = MIMEMultipart()
                message["From"] = email
                message["To"] = rec_email
                message["Subject"] = subject

                # Add body to email
                message.attach(MIMEText(msg, "plain"))

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


def send_email_with_buttons(sender_email, receiver_email, sender_password, accept_link, decline_link,employee_id_1,employee_id_2,date_employee_1,month):
    # Read Excel file into a DataFrame
    df = pd.read_excel("modified_roster.xlsx")

    emp_df = pd.read_excel("employee_chatbot.xlsx")

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
        <p>You have a shift swap request on {col_date_employee_1} by {employee_name_1}. Please respond:</p>
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



def swap_pre_process(employee_id_1,employee_id_2,date_employee_1):
    sender_email = "vishal.d2019cse@sece.ac.in"
    sender_password = "Dkvm@2016"
    accept_link = "https://b2411a61-a517-4ae4-9b30-5cbd4e3a793d-00-xlmxsla4v0nm.worf.replit.dev:5000/accept"
    decline_link = "https://b2411a61-a517-4ae4-9b30-5cbd4e3a793d-00-xlmxsla4v0nm.worf.replit.dev:5000/decline"
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
    email_thread = threading.Thread(target=send_email_and_wait,
                                    args=(employee_id_1, employee_id_2, date_employee_1))
    email_thread.start()

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
    month = admin_df.loc[0, "Month"]



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

    return {"message":f"Values swapped successfully. Modified data saved to {output_path}"}

@app.route('/')
def index():
    return 'Hello from Vishal!'

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
            return swap_pre_process(item['employee_id'],item['swap_id'],item['swap_date_1'])
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

    

if __name__ == '__main__':
    app.run(debug=True, port=5002, host='0.0.0.0')