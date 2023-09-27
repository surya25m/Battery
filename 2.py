import pandas as pd

# Read the cation data from Excel
cation_df = pd.read_excel('All_cation_BE.xlsx')

# Read the anion data from Excel
anion_df = pd.read_excel('All_anion_BE.xlsx')

# Initialize an empty list to store the voltage data
voltage_data = []

# Iterate over each combination of cation and anion
for _, cation_row in cation_df.iterrows():
    for _, anion_row in anion_df.iterrows():
        cation_stage = cation_row['Cation_Stage']
        anion_stage = anion_row['Anion_Stage']
        cation_be = cation_row['Cation_BE']
        anion_be = anion_row['Anion_BE']
        no_of_cations = cation_row['No. of cations']
        no_of_anions = anion_row['No. of anions']

        # Check conditions for calculating voltage
        if cation_stage == anion_stage and no_of_cations == no_of_anions:
            voltage = ((no_of_cations * cation_be) + (no_of_anions * anion_be)) / no_of_cations

            # Store the voltage data with cation and anion details
            voltage_data.append({
                'Cation': cation_row['Cation'],
                'Cation_Stage': cation_stage,
                'No. of Cations': no_of_cations,
                'Anion': anion_row['Anion'],
                'Anion_Stage': anion_stage,
                'No. of Anions': no_of_anions,
                'Voltage': voltage
            })

# Create a new DataFrame from the voltage data
voltage_df = pd.DataFrame(voltage_data)

# Write the voltage data to a new Excel file
voltage_df.to_excel('Voltage_Data.xlsx', index=False)
