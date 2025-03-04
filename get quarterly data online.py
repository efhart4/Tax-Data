import requests
import pandas as pd

# Data comes from here: 
# https://www.census.gov/programs-surveys/qtax/data/tables.All.html



# Loop to Download Quarterly Data from 1995 to 2024
#------------------------------------
for years in range(1995, 2025):
    for myquarter in range(1, 5):
        
        # Break the loop after 2024 Q3
        if years == 2024 and myquarter > 3:
            break
        # URL of the Excel file
        if years == 1999:
            url = f'https://www2.census.gov/programs-surveys/qtax/tables/{years}/qtx99{myquarter}t3.xls'
            file_path = f'collected quarterly data/q{myquarter}t3_{years}.xls'
        elif years < 2021 or (years == 2021 and myquarter < 3):
            url = f'https://www2.census.gov/programs-surveys/qtax/tables/{years}/q{myquarter}t3.xls'
            file_path = f'collected quarterly data/q{myquarter}t3_{years}.xls'
        else:
            url = f'https://www2.census.gov/programs-surveys/qtax/tables/{years}/q{myquarter}t3.xlsx'
            file_path = f'collected quarterly data/q{myquarter}t3_{years}.xlsx'


        # Download the file
        try:
            response = requests.get(url)
            response.raise_for_status()  # Raise an HTTPError for bad responses (4xx and 5xx)
            with open(file_path, 'wb') as file:
                file.write(response.content)
            print(f'File downloaded successfully for {years} Q{myquarter}.')
        except requests.exceptions.RequestException as e:
            print(f'Failed to download file for {years} Q{myquarter}. Error: {e}')
            exit()

        # Load the Excel file into a pandas DataFrame
        try:
            if years < 2011 or (years == 2011 and myquarter == 1):
                quarterly_report = pd.read_excel(file_path, skiprows=7)
            else:
                quarterly_report = pd.read_excel(file_path, skiprows=5)
            print('Excel file loaded successfully.')
            # Preview the data
            print(quarterly_report.head())

        except Exception as e:
            print(f'Failed to load Excel file for {years} Q{myquarter}. Error: {e}')

