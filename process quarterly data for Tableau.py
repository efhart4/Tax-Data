
# Import necessary libraries
import pandas as pd

#First function to process data. I think this can be used from 2010 Q2 and forward
# function to process quarterly report for most recent type
#------------------------------------
def process_post_2010_quarterly_report(quarterly_report, year, quarter):

    # Rename the total row to 'T00' before filtering
    total_tax_descriptions = ['Total Taxes', 'Total Tax', 'Total', 'Total Of All Taxes', 'TOTAL']
    quarterly_report.loc[quarterly_report.iloc[:, 0].isin(total_tax_descriptions), quarterly_report.columns[1]] = 'T00'

    # in the second column select only rows with TXX code for tax type
    #------------------------------------
    filtered_df = quarterly_report[quarterly_report.iloc[:, 1].astype(str).str.startswith('T') & 
                                   (quarterly_report.iloc[:, 1] != 'T18') & 
                                   (quarterly_report.iloc[:, 1] != 'T14')]

    # ALSO KEEP TOTAL TAXES ROW
    filtered_df = pd.concat([filtered_df, quarterly_report[quarterly_report.iloc[:, 1] == 'Total Taxes']], ignore_index=True)

    # rename the second column to 'new_code'
    filtered_df = filtered_df.rename(columns={filtered_df.columns[1]: 'new_code'})
    
    # rename various versions of Washington, D.C. to a consistent format
    dc_variations = ['Ex. Wash. DC', 'Wash. DC', 'Washington DC', 'Washington, DC', 'D.C.']
    for variation in dc_variations:
        filtered_df = filtered_df.rename(columns={variation: 'Washington, D.C.'})


    # remove first column with verbal description
    #------------------------------------
    filtered_df = filtered_df.drop(columns=[filtered_df.columns[0]])

    # remove columns that provide no data
    #------------------------------------
    # Get the column names of the DataFrame
    shifted_columns = filtered_df.columns.tolist()
    shifted_columns = ['new_code']+ ['NA'] + ['U.S. State Total'] + shifted_columns[2:-1]

    # Assign the shifted column names back to the DataFrame
    filtered_df.columns = shifted_columns

    # Keep every other column starting from the first
    columns_to_keep = list(filtered_df.columns[::2])

    # Use .loc to select columns by their names
    filtered_df = filtered_df.loc[:, columns_to_keep]


    # Remove trailing stars from state names in column names
    #------------------------------------
    filtered_df.columns = filtered_df.columns.str.rstrip('*')


    # add column for date
    #------------------------------------
    # create a date object for the quarter
    date = pd.Period(year=year, quarter=quarter, freq='Q')
    filtered_df['quarter_collected'] = date


    # turn X and x into 0
    #------------------------------------
    filtered_df = filtered_df.replace(['X', 'x'], 0)

    # Ensure the replacement results in numeric 0
    filtered_df = filtered_df.apply(pd.to_numeric, errors='ignore')


    # restructure data so that there is one row for the date and a column for ech states and tax type
    #------------------------------------
    # Pivot the data 
    pivoted_data = filtered_df.pivot_table(index='quarter_collected', columns='new_code', values=filtered_df.columns[1:])

    # Flatten the columns
    pivoted_data.columns = ['_'.join(col).strip() for col in pivoted_data.columns.values]
    
    return pivoted_data


def process_pre_2010_quarterly_report(quarterly_report, year, quarter):
    # Rename the total row to 'T00' before filtering
    total_tax_descriptions = ['Total Taxes', 'Total Tax', 'Total', 'Total Of All Taxes', 'TOTAL']
    quarterly_report.loc[quarterly_report.iloc[:, 0].isin(total_tax_descriptions), quarterly_report.columns[1]] = 'T00'


    # Filter the data
    filtered_df = quarterly_report[quarterly_report.iloc[:, 1].astype(str).str.startswith('T') & 
                                (quarterly_report.iloc[:, 1] != 'T18') & 
                                (quarterly_report.iloc[:, 1] != 'T14') &
                                (quarterly_report.iloc[:, 1] != 'T02')]
    


    # Rename the second column to 'new_code'
    filtered_df = filtered_df.rename(columns={filtered_df.columns[1]: 'new_code'})


    # rename the third column to 'US State Total'
    filtered_df = filtered_df.rename(columns={filtered_df.columns[2]: 'U.S. State Total' })

    # rename various versions of Washington, D.C. to a consistent format
    dc_variations = ['Ex. Wash. DC', 'Wash. DC', 'Washington DC', 'Washington, DC', 'D.C.']
    for variation in dc_variations:
        filtered_df = filtered_df.rename(columns={variation: 'Washington, D.C.'})
    
    # remove first column with verbal description
    #------------------------------------
    filtered_df = filtered_df.drop(columns=[filtered_df.columns[0]])
    
    # add column for date
    #------------------------------------
    # create a date object for the quarter
    date = pd.Period(year=year, quarter=quarter, freq='Q')
    filtered_df['quarter_collected'] = date
    
    # turn X and x into 0
    #------------------------------------
    filtered_df = filtered_df.replace(['X', 'x'], 0)
    
    # Ensure the replacement results in numeric 0
    filtered_df = filtered_df.apply(pd.to_numeric, errors='ignore')
    
    # restructure data so that there is one row for the date and a column for each state X tax type
    #------------------------------------
    # Pivot the data 
    pivoted_data = filtered_df.pivot_table(index='quarter_collected', columns='new_code', values=filtered_df.columns[1:])
    # Flatten the columns
    pivoted_data.columns = ['_'.join(col).strip() for col in pivoted_data.columns.values]
    
    
    #print(filtered_df.head())
    return pivoted_data


# Loop to Download Quarterly Data from 1995 to 2024
#------------------------------------
historical_data = pd.DataFrame()
# Initialize a list to store mismatched reports
mismatched_reports = []

for years in range(1995, 2025):
    for myquarter in range(1, 5):
        
        # Break the loop after 2024 Q3
        if years == 2024 and myquarter > 2:
            break

        # Get path for saved file
        if years == 1999:
            file_path = f'collected quarterly data/q{myquarter}t3_{years}.xls'
        elif years < 2021 or (years == 2021 and myquarter < 3):
            file_path = f'collected quarterly data/q{myquarter}t3_{years}.xls'
        else:
            file_path = f'collected quarterly data/q{myquarter}t3_{years}.xlsx'

        # Load the Excel file into a pandas DataFrame, while skipping the correct number of rows. 
        try:
            if years < 1997 or (years == 1997 and myquarter <= 2): 
                quarterly_report = pd.read_excel(file_path, skiprows=8)
            
            elif  (years == 1997 and myquarter >= 3) or (years > 1997 and years <2004) or (years ==2004 and myquarter <=2): 
                quarterly_report = pd.read_excel(file_path, skiprows=7) 
            
            elif (years >=2005 and years<= 2010) or (years ==2004 and myquarter >2):
                quarterly_report = pd.read_excel(file_path, skiprows=6) 
    
            else:
                quarterly_report = pd.read_excel(file_path, skiprows=5)
            
            print(f'Excel file loaded successfully from {file_path}.')

        except Exception as e:
            print(f'Failed to load Excel file for {years} Q{myquarter}. Error: {e}')


        # now using the loaded data to create the observation
        #------------------------------------
        if (years< 2010) or (years == 2010 and myquarter == 1):
            quarterly_report = process_pre_2010_quarterly_report(quarterly_report, years, myquarter)
        else :
            quarterly_report = process_post_2010_quarterly_report(quarterly_report, years, myquarter)


        # Identify and print mismatched columns IF WE HAVE PASSSED THE FIRST YEAR
        if not historical_data.empty:
            # Get the columns of both DataFrames
            historical_columns = set(historical_data.columns)
            current_columns = set(quarterly_report.columns)

            # Find mismatched columns
            mismatched_columns = historical_columns ^ current_columns

            if mismatched_columns:
                mismatched_values = {col: historical_data[col].iloc[-1] if col in historical_data.columns else 'N/A' for col in mismatched_columns}
                mismatched_reports.append((years, myquarter, mismatched_columns, mismatched_values))



        # Concatenate the processed data to historical_data 
        historical_data = pd.concat([historical_data, quarterly_report])

# Define the dictionary to map tax codes to descriptions
tax_code_to_description = {
    'T00': 'Total Taxes',
    'T01': 'Property taxes T01',
    'T09': 'General sales and gross receipts T09',
    'T13': 'Motor fuels T13',
    'T10': 'Alcoholic beverages T10',
    'T15': 'Public utilities T15',
    'T12': 'Insurance premiums T12',
    'T16': 'Tobacco products T16',
    'T14': 'Pari-mutuels T14',
    'T11': 'Amusements T11',
    'T19': 'Other selective sales and gross receipts T19',
    'T20': 'Alcoholic beverages T20',
    'T27': 'Public utilities T27',
    'T24': 'Motor vehicles T24',
    'T25': 'Motor vehicle operators T25',
    'T22': 'Corporations in general T22',
    'T23': 'Hunting and fishing T23',
    'T21': 'Amusements T21',
    'T28': 'Occupation and businesses T28',
    'T29': 'Other license taxes T29',
    'T40': 'Individual income T40',
    'T41': 'Corporation net income T41',
    'T50': 'Death and gift T50',
    'T53': 'Severance T53',
    'T51': 'Documentary and stock transfer T51',
    'T99': 'Other taxes, NEC T99'
}

# Define the dictionary to map tax codes to descriptions
tax_code_to_description = {
    'T00': 'Total Taxes',
    'T01': 'Property taxes T01',
    'T09': 'General sales and gross receipts T09',
    'T13': 'Motor fuels T13',
    'T10': 'Alcoholic beverages T10',
    'T15': 'Public utilities T15',
    'T12': 'Insurance premiums T12',
    'T16': 'Tobacco products T16',
    'T14': 'Pari-mutuels T14',
    'T11': 'Amusements T11',
    'T19': 'Other selective sales and gross receipts T19',
    'T20': 'Alcoholic beverages T20',
    'T27': 'Public utilities T27',
    'T24': 'Motor vehicles T24',
    'T25': 'Motor vehicle operators T25',
    'T22': 'Corporations in general T22',
    'T23': 'Hunting and fishing T23',
    'T21': 'Amusements T21',
    'T28': 'Occupation and businesses T28',
    'T29': 'Other license taxes T29',
    'T40': 'Individual income T40',
    'T41': 'Corporation net income T41',
    'T50': 'Death and gift T50',
    'T53': 'Severance T53',
    'T51': 'Documentary and stock transfer T51',
    'T99': 'Other taxes, NEC T99'
}

# Convert tax codes to descriptions in the historical_data DataFrame
historical_data.columns = pd.MultiIndex.from_tuples(
    [tuple(col.rsplit('_', 1)) for col in historical_data.columns]
)
historical_data.columns = historical_data.columns.set_levels(
    [historical_data.columns.levels[0], [tax_code_to_description.get(code, code) for code in historical_data.columns.levels[1]]]
)

restructured_data = pd.DataFrame()
# Loop through all top-level indices (state names) and add a column with the state name for each row
state_names = historical_data.columns.get_level_values(0).unique()
for state in state_names:
    temp_db = historical_data[state].copy()
    #print(temp_db.head())
    temp_db['State'] = state
    restructured_data = pd.concat([restructured_data, temp_db.reset_index()], ignore_index=True)

# Separate the 'quarter_collected' column
quarter_collected = restructured_data['quarter_collected']

# Replace negative numbers with 0, excluding the 'quarter_collected' column
restructured_data = restructured_data.applymap(lambda x: 0 if isinstance(x, (int, float)) and x < 0 else x)

# Reassign the 'quarter_collected' column back to the DataFrame
restructured_data['quarter_collected'] = quarter_collected

print(restructured_data)

# Save the final DataFrame to an Excel file
restructured_data.to_excel('processed quarterly data/processed_quarterly_data_Tableau.xlsx', index=False)