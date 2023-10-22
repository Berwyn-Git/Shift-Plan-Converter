import openpyxl

# Load the workbook
workbook = openpyxl.load_workbook('beisen_shift_plan.xlsx')

# Select the worksheet
sheet = workbook['sheet1']

# Set the email address list
email_address_list = []

# Loop through each row in the sheet and append the name and email to the list
for row in sheet.iter_rows(min_row=2, values_only=True):
    name, email, *_ = row[:2]
    email_address_list.append((name, email))

email_dict = {row[0]: row[1] for row in email_address_list}
# print(email_dict)