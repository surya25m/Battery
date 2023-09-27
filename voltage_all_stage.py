import pandas as pd

# Read the cation and anion data from Excel files
cation_data = pd.read_excel('All_cation_BE_cation_num.xlsx')
anion_data = pd.read_excel('All_anion_BE_num.xlsx')

# Create an empty DataFrame to store the results
results = pd.DataFrame(columns=['Cation', 'Cation_Stage', 'No. of cations',
                                'Anion', 'Anion_Stage', 'No. of anions', 'Voltage'])

# Iterate through each cation and anion combination
for cation_index, cation_row in cation_data.iterrows():
    for anion_index, anion_row in anion_data.iterrows():
        # Check if cation_stage and anion_stage are the same
#        if cation_row['Cation_Stage'] == anion_row['Anion_Stage']:
            # Check if the number of cations and anions are equal
            if cation_row['No. of cations'] == anion_row['No. of anions']:
                # Calculate the voltage
                voltage = abs((cation_row['No. of cations'] * cation_row['Cation_BE']) +
                           (anion_row['No. of anions'] * anion_row['Anion_BE'])) / cation_row['No. of cations']

                # Store the results in the DataFrame
                results = results.append({
                    'Cation': cation_row['Cation'],
                    'Cation_Stage': cation_row['Cation_Stage'],
                    'No. of cations': cation_row['No. of cations'],
                    'Anion': anion_row['Anion'],
                    'Anion_Stage': anion_row['Anion_Stage'],
                    'No. of anions': anion_row['No. of anions'],
                    'Voltage': voltage
                }, ignore_index=True)

# Write the results to a new Excel file
results.to_excel('Voltage_all_stage_by_num.xlsx', index=False)
