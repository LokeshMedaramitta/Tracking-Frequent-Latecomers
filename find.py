import pandas as pd
import datetime

excel_file = pd.ExcelFile('filename.xlsx')

student_data = {}

for i, sheet_name in enumerate(excel_file.sheet_names):

    df = pd.read_excel(excel_file, sheet_name)
    df.columns = df.columns.str.strip()


    year = (i // 3) + 1
    for index, row in df.iterrows():
        if row['ID'] not in student_data:
            student_data[row['ID']] = {}
        if year not in student_data[row['ID']]:
            student_data[row['ID']][year] = {
                'Name': row['Name'],
                'latecomer': 0,
                'latecomer_count': {}
            }

while True:
    date_input = input("Enter the date (dd-mm-yyyy): ")
    try:
        date = datetime.datetime.strptime(date_input, '%d-%m-%Y').date()
    except ValueError:
        print("Invalid date format. Please enter date in the format dd-mm-yyyy.")
        continue

    year_input = input(f"Enter year (1-4) for {date.strftime('%d-%m-%Y')}: ")
    if year_input.lower() == 'exit':
        break

    try:
        year = int(year_input)
        if year not in [1, 2, 3, 4]:
            print("Invalid year input. Please enter a number between 1 and 4.")
            continue
    except ValueError:
        print("Invalid year input. Please enter a number between 1 and 4.")
        continue

    latecomers_input = input("Enter latecomers as a comma-separated list of ID suffixes: ")
    for id in student_data.keys():
        for y in range(1, 4):
            if id in student_data and y in student_data[id]:
                if id[-3:] in latecomers_input.split(',') and y == year:
                    student_data[id][year]['latecomer'] += 1
                student_data[id][y]['latecomer_count'] = student_data[id][y]['latecomer']
    
    data_groups = {}
    for id, years in student_data.items():
        for y, data in years.items():
            if y not in data_groups:
                data_groups[y] = []
            data_groups[y].append((data['Name'], id, data['latecomer_count']))
    
    year_latecomers = [data for data in data_groups.get(year, []) if data[2] > 0]
    if year_latecomers:
        print(f"Year {year}:")
        for data in year_latecomers:
            if 1 <= data[2] <= 3:
                print(f"Name: {data[0]}, ID: {data[1]}, Latecomer Count: {data[2]} (Warning: 1-3 Latecomers)")
            elif 3 < data[2] <= 10:
                print(f"Name: {data[0]}, ID: {data[1]}, Latecomer Count: {data[2]} (Attention: 4-10 Latecomers)")
            elif data[2] > 10:
                print(f"Name: {data[0]}, ID: {data[1]}, Latecomer Count: {data[2]} (Attention:>10 Latecomers)")