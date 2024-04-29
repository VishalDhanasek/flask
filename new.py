import pandas as pd
import re
# from datetime import datetime
# Define the questions and their respective validation patterns/messages
questions = [
    ("employee_name", "Enter employee name: ", "^[a-zA-Z .]+$", "Invalid input. Please enter text only."),
    ("employee_id", "Enter employee ID (numbers only): ", "^\\d+$", "Invalid input. Please enter numbers only."),
    ("Preferred_Date_1", "Enter first preferred date (DD): ", r"^(0[0-9]|[12][0-9]|3[01])$",
     "Invalid date format. Please use DD."),
    ("Preferred_Date_2", "Enter second preferred date (DD) (press Enter to skip): ",
     r"^(0[0-9]|[12][0-9]|3[01])$|^$",
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
def ask_questions():
    responses = {}
    for question, prompt, pattern, error_message in questions:
        while True:
            response = input(prompt)
            if pattern:
                if response.strip():
                    if re.match(pattern, response):
                        if question.startswith("Preferred_Date") and not is_valid_date(response):
                            print("Invalid date. Please enter a valid day of the month.")
                        else:
                            responses[question] = response
                            break
                    else:
                        print(error_message)
                else:
                    # Set response to None if input is empty
                    responses[question] = None
                    break
            else:
                responses[question] = response
                break
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
        preferred_dates = [responses['Preferred_Date_1'], responses['Preferred_Date_2']]
        if all(date is None for date in preferred_dates):
            print("Both preferred dates are None")
        elif any(date is None for date in preferred_dates):
            print("At least one preferred date is None")
        else:
            # Convert string representations to actual values using eval
            preferred_dates = [eval(i) for i in preferred_dates]
            print(preferred_dates)
        # print(preferred_dates, preferred_dates[0], type(preferred_dates[0]), existing_data['Preferred_Date_1'],
        # type(existing_data['Preferred_Date_1']))
        dates = existing_data[['Preferred_Date_1', 'Preferred_Date_2']].stack().dropna().astype(int)
        print(dates)
        print(type(dates))
        for date in preferred_dates:
            for iterator in dates:
                if iterator == date:
                    if iterator not in leave_request:
                        leave_request[iterator] = 0
                        print(leave_request[iterator])
                    leave_request[iterator]+=1
                    print(leave_request[iterator])
                    if leave_request[iterator]>=2:
                        print(f"Preferred date {date} is already taken. Please choose another date.")
                        return
        if responses['employee_id'] in existing_data['employee_id'].values:
            existing_data.loc[existing_data['employee_id'] == responses['employee_id'], 'Preferred_Date_1'] = responses[
                'Preferred_Date_1']
            existing_data.loc[existing_data['employee_id'] == responses['employee_id'], 'Preferred_Date_2'] = responses[
                'Preferred_Date_2']
            print("Employee data updated.")
        else:
            # Append new responses to existing data
            existing_data = pd.concat([existing_data, pd.DataFrame(responses, index=[0])], ignore_index=True)
            print("New employee added.")
        # Write the combined data to the Excel file
        existing_data.to_excel("C:\\Users\\veeru\\Downloads\\project.xlsx", index=False, engine='openpyxl')
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame with the new responses and write it to the Excel file
        df = pd.DataFrame(responses, index=[0])
        df.to_excel("C:\\Users\\veeru\\Downloads\\project.xlsx", index=False, engine='openpyxl')

# Main function
def main():
    print("Welcome to the Chatbot!")
    responses = ask_questions()
    save_to_excel(responses)
    print("Thank you! Your responses have been saved.")

if __name__ == "__main__":
    main()

