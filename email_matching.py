import pandas as pd

# Load the employee data
employee_data = pd.read_excel('beisen.xlsx')
employee_data['Full Name'] = employee_data['Last Name'] + ' ' + employee_data['First Name']
employee_data['Email'] = ""

# Load the email data
email_dict = {"张三": "zhangsan@example.com",
              "李四": "lisi@example.com",
              "王五": "wangwu@example.com"}

# Match the employee names with the email dictionary
for i, employee_name in enumerate(employee_data['Full Name']):
    for email_name, email_address in email_dict.items():
        if email_name in employee_name or employee_name.split()[0] in email_name or employee_name.split()[1] in email_name:
            employee_data.loc[i, 'Email'] = email_address
            break

# Print the employee data with email addresses
print(employee_data)
