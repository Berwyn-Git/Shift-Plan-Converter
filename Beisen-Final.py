import pandas as pd
import re
from datetime import datetime

from email_address import email_address_list

# Read the original data sheet
df = pd.read_excel('Beisen_April.xlsx')

# Copy the data from the 19th row to a new worksheet
df_new = df.iloc[:].copy()

# Rename the first column to "*姓名" and insert a new column as the second column with the title "*邮箱"
df_new = df_new.rename(columns={df_new.columns[0]: '*姓名'})
df_new.insert(1, '*邮箱', '')

# create dictionary from email_address_list
email_dict = {row[0]: row[1] for row in email_address_list}
# print(email_dict)


def match_name(name, email_dict):
    for email_name in email_dict.keys():
        # Split the name into Chinese and English parts
        chinese_part = re.findall(r'[\u4e00-\u9fff]+', name)
        print(chinese_part)
        email_parts = re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', email_name)
        # Check if either the Chinese or English part matches
        if any(part in email_parts for part in chinese_part):
            return email_dict[email_name]
    return None


# map names to email addresses using email_dict
df_new['*邮箱'] = df_new['*姓名'].astype(str).apply(lambda x: match_name(x, email_dict))

# Replace all the values in the worksheet
for col in df_new.columns[2:]:
    for i, val in df[col].items():
        if pd.isna(val):
            continue
        if "-" in str(val):
            # print("this is the ok val:", val)
            start, end = str(val).split("-", 1)
            test_pattern = re.compile(r'\d{1,2}:\d{2}')
            matches = test_pattern.findall(str(val))
            if len(matches) == 2:
                start_time = pd.to_datetime(matches[0], format="%H:%M")
                start = start_time.strftime('%H:%M')
                end_time = pd.to_datetime(matches[1], format="%H:%M")
                end = end_time.strftime('%H:%M')
                # print(type(end))
                if end_time == '24:00':
                    end = pd.to_datetime('00:00', format="%H:%M")
                else:
                    end_time = pd.to_datetime(end.strip(), format="%H:%M")
                    end = end_time.strftime('%H:%M')
                    # print(type(end))
                duration = (end_time - start_time).seconds / 3600

                weekday = pd.to_datetime(col).weekday()
                # print(duration)
                if duration >= 8.5:
                    if start_time == pd.to_datetime("07:00", format="%H:%M") and weekday < 4.5:
                        df_new.loc[i, col] = "E-测试"
                    elif start_time == pd.to_datetime("15:30", format="%H:%M") and weekday < 4.5:
                        df_new.loc[i, col] = "M-测试"
                    elif start_time == pd.to_datetime("07:00", format="%H:%M") and weekday > 4.5:
                        df_new.loc[i, col] = "周末E-测试"
                    elif start_time == pd.to_datetime("15:30", format="%H:%M") and weekday > 4.5:
                        df_new.loc[i, col] = "周末M-测试"
                    else:
                        print("in duration >= 8.5, find this:", start, end, duration, "and row:", i, "and col:", col)
                if duration > 7:
                    if end_time == pd.to_datetime("07:00", format="%H:%M") and weekday < 4.5:
                        df_new.loc[i, col] = "N-测试"
                    elif end_time == pd.to_datetime("07:00", format="%H:%M") and weekday > 4.5:
                        df_new.loc[i, col] = "周末N-非7小时加班-测试"
                    else:
                        print("in duration > 7, find this:", start, end, duration, "and row:", i, "and col:", col)
                if duration == 7:
                    if end_time == pd.to_datetime("07:00", format="%H:%M") and weekday > 4.5:
                        df_new.loc[i, col] = "周末N-7小时加班-测试"
                    if end_time == pd.to_datetime("07:00", format="%H:%M") and weekday < 4.5:
                        df_new.loc[i, col] = "N-测试"
                    else:
                        print("in duration = 7, find this:", start, end, duration, "and row:", i, "and col:", col)
                else:
                    print(start, end, duration)
            else:
                print(matches)
        elif str(val) == 'Rest':
            df_new.loc[i, col] = '普通休息日'
        # else:
        #     print(str(val))

        # Add check for selected names and apply replacement rules
        if df_new.loc[i, '*姓名'] in ['Xia Feng夏锋', 'Zhai Ming翟明', 'Huang Changyin黄昌银', 'Meng Xiangying孟祥营']:
            # print("find it!")
            if df_new.loc[i, col] == '周末E-测试':
                df_new.loc[i, col] = 'E-测试'
            elif df_new.loc[i, col] == '周末M-测试':
                df_new.loc[i, col] = 'M-测试'
            elif df_new.loc[i, col] == '周末N-7小时加班-测试':
                df_new.loc[i, col] = 'N-测试'
            elif df_new.loc[i, col] == '周末N-非7小时加班-测试':
                df_new.loc[i, col] = 'N-测试'

# drop rows containing NaN values in '*姓名' column
df_new.dropna(subset=['*姓名'], inplace=True)

# filter for rows containing Chinese characters in '*姓名' column
df_new = df_new[df_new['*姓名'].str.contains('[\u4e00-\u9fff]')]

# Save the new worksheet to
df_new.to_excel('beisen_new.xlsx', index=False)
