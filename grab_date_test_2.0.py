import pandas as pd
import re

# read source data and extract employee names and match dates
# source_df = pd.read_excel('PD shiftplan March.xlsx')
source_df = pd.read_excel('PD shiftplan March.xlsx', parse_dates=[])

source_df.iloc[:, 0] = source_df.iloc[:, 0].astype(str)
employee_names = source_df.iloc[:, 0].tolist()  # selects all rows in the first column of the source
pattern = r'^[\u4e00-\u9fa5]+\s[a-zA-Z]+$|^[a-zA-Z]+\s[\u4e00-\u9fa5]+$'
employee_names = [name for name in source_df.iloc[:, 0].tolist() if re.match(pattern, name)]
match_dates = source_df.columns[1:].tolist()    # extracts the column headers for the columns starting from the third column (index 2) in the source

# read finished file and initialize with employee names and match dates
finished_df = pd.read_excel('finished_file.xlsx')
finished_df = finished_df.iloc[:, :2]   # selecting all rows and only the first two columns of the source
finished_df.columns = ['*姓名', '*邮箱']
finished_df = finished_df.reindex(columns=['*姓名', '*邮箱'] + match_dates)  # to conform the DataFrame to a new index with specified column labels

# loop over each cell in source data and update corresponding cell in finished file
for i, employee_name in enumerate(employee_names):
    # find matching row in finished file (accounting for different name formats)
    matching_rows = finished_df[finished_df['*姓名'].str.contains(employee_name.split()[0]) &
                                 finished_df['*姓名'].str.contains(employee_name.split()[-1])]
    if len(matching_rows) == 0:
        print(f"No matching row found for {employee_name}.")
        continue
    elif len(matching_rows) > 1:
        print(f"Multiple matching rows found for {employee_name}. Using first match.")
    row_index = matching_rows.index[0]

    # loop over each date column and update cell value if necessary
    for date in match_dates:
        shift_value = source_df.iloc[i, source_df.columns.get_loc(date)]
        if shift_value == 'E':
            finished_df.at[row_index, date] = 'E-测试'
        elif shift_value == 'M':
            finished_df.at[row_index, date] = 'M-测试'
        elif shift_value == 'N':
            finished_df.at[row_index, date] = 'N-测试'

# save updated finished file
finished_df.to_excel('finished_file_updated.xlsx', index=False)
