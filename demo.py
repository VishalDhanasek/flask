from flask import Flask,request
from flask_cors import CORS
import pandas as pd
import re


app = Flask(__name__)
cors = CORS(app)
# from datetime import datetime
# Define the questions and their respective validation patterns/messages
questions = [
    (0,"employee_name", "Enter employee name: ", "^[a-zA-Z .]+$", "Invalid input. Please enter text only."),
    (1,"employee_id", "Enter employee ID (numbers only): ", "^\\d+$", "Invalid input. Please enter numbers only."),
    (2,"Planned_Leave_1", "Enter first planned leave (DD format) (press Enter to skip): ", r"^(0[1-9]|[12][0-9]|3[01])$",
     "Invalid date format. Please use DD."),
    (3,"Planned_Leave_2", "Enter second planned leave (DD format) (press Enter to skip): ",
     r"^(0[1-9]|[12][0-9]|3[01])$|^$",
     "Invalid date format. Please use DD or leave it empty.")
]

# Function to validate if the entered date is valid within a month
def is_valid_date(day):
    try:
        day = int(day)
        return 1 <= day <= 31
    except ValueError:
        return False

# Function to ask questions and validate responses
def ask_questions(response_list):
    responses = {}
    for idx, question, prompt, pattern, error_message in questions:
        # while True:
            response = response_list[idx]
            if pattern:
                if response.strip():
                    if re.match(pattern, response):
                        if question.startswith("Preferred_Date") and not is_valid_date(response):
                            print("Invalid date. Please enter a valid day of the month.")
                        else:
                            responses[question] = response
                            # break
                    else:
                        print(error_message)
                else:
                    # Set response to None if input is empty
                    responses[question] = None
                    # break
            else:
                responses[question] = response
                # break
    return responses

# Function to save responses to Excel file
def save_to_excel(responses):
    leave_request = {}
    try:
        # Read existing data from Excel file if it exists
        existing_data = pd.read_excel("C:\\Users\\veeru\\Downloads\\project.xlsx", engine='openpyxl')
        # Convert existing employee_id values to integers for comparison
        existing_data['employee_id'] = existing_data['employee_id'].astype(int)
        # Convert responses employee_id to integer for comparison
        responses['employee_id'] = int(responses['employee_id'])
        # Check if the preferred dates are already taken
        preferred_dates = [responses['Planned_Leave_1'], responses['Planned_Leave_2']]
        if all(date is None for date in preferred_dates):
            return("Both preferred dates are None")
        elif any(date is None for date in preferred_dates):
            return("At least one preferred date is None")
        else:
            # Convert string representations to actual values using eval
            # Convert string representations to actual integers
            preferred_dates = [int(date) for date in preferred_dates]
        # print(preferred_dates, preferred_dates[0], type(preferred_dates[0]), existing_data['Preferred_Date_1'],
        # type(existing_data['Preferred_Date_1']))
        dates = existing_data[['Planned_Leave_1', 'Planned_Leave_2']].stack().dropna().astype(int)
        for date in preferred_dates:
            for iterator in dates:
                if iterator == date:
                    if iterator not in leave_request:
                        leave_request[iterator] = 0
                        
                    leave_request[iterator]+=1
                    if leave_request[iterator]>=2:
                        return(f"Preferred date {date} is already taken. Please choose another date.")
                        
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
        existing_data.to_excel("C:\\Users\\veeru\\Downloads\\project.xlsx", index=False, engine='openpyxl')
        return ("Thank you! Your responses have been saved.")
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame with the new responses and write it to the Excel file
        df = pd.DataFrame(responses, index=[0])
        df.to_excel("C:\\Users\\veeru\\Downloads\\project.xlsx", index=False, engine='openpyxl')
        return ("Thank you! Your responses have been saved.")

# Main function
@app.route('/main', methods=['POST'])
def main():
    item = request.get_json()
    print("item -- ",item)
    employee_name = item['employee_name']
    print("employee_name -- ",employee_name)
    employee_id = item['employee_id']
    print("employee_id -- ",employee_id)
    Preferred_Date_1 = item['Planned_Leave_1']
    print("Preferred_Date_1 -- ",Preferred_Date_1)
    Preferred_Date_2 = item['Planned_Leave_2']
    print("Preferred_Date_2 -- ",Preferred_Date_2)
    responses = ask_questions([employee_name, employee_id, Preferred_Date_1, Preferred_Date_2])
    print("responses -- ",responses)
    if 'error' in responses:
        return responses
    else:
        return save_to_excel(responses)

if __name__ == "__main__":
    app.run(debug=True, port=5002, host='0.0.0.0')