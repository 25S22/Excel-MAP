import pandas as pd

# Load the Excel files
sheet1 = pd.read_excel('sheet1.xlsx')  # Columns: Instance Name, Private Ip, os(updated), environment, owner, application, department
sheet2 = pd.read_excel('sheet2.xlsx')  # Columns: Hostname, IP Address, Physical/Virtual, OS, Environment, Primary Owner, Application Name, Business entity, Location
sheet3 = pd.read_excel('sheet3.xlsx')  # Columns: SRNO, LOCATION, HOSTNAME, IP ADDRESS, DEVICE TYPE, STATUS, MAKE, MODEL

# Define final structure
final_columns = [
    'SR. NO.', 'Asset Category', 'DEVICE TYPE', 'LOCATION', 'HOST NAME', 'IP ADDRESS',
    'OS', 'Environment', 'STATUS', 'MAKE', 'MODEL', 'OWNER', 'Hosted Application',
    'Physical/Virtual', 'Business Unit'
]

# Sheet 1 Mapping (Cloud)
mapped1 = pd.DataFrame({
    'SR. NO.': range(1, len(sheet1) + 1),
    'Asset Category': ['Cloud'] * len(sheet1),
    'DEVICE TYPE': [None] * len(sheet1),
    'LOCATION': [None] * len(sheet1),
    'HOST NAME': sheet1['Instance Name'],
    'IP ADDRESS': sheet1['Private Ip'],
    'OS': sheet1['os(updated)'],
    'Environment': sheet1['environment'],
    'STATUS': [None] * len(sheet1),
    'MAKE': [None] * len(sheet1),
    'MODEL': [None] * len(sheet1),
    'OWNER': sheet1['owner'],
    'Hosted Application': sheet1['application'],
    'Physical/Virtual': ['Virtual'] * len(sheet1),
    'Business Unit': sheet1['department']
})

# Sheet 2 Mapping (On-Prem)
mapped2 = pd.DataFrame({
    'SR. NO.': range(len(sheet1) + 1, len(sheet1) + len(sheet2) + 1),
    'Asset Category': ['On-Prem'] * len(sheet2),
    'DEVICE TYPE': [None] * len(sheet2),
    'LOCATION': sheet2['Location'],
    'HOST NAME': sheet2['Hostname'],
    'IP ADDRESS': sheet2['IP Address'],
    'OS': sheet2['OS'],
    'Environment': sheet2['Environment'],
    'STATUS': [None] * len(sheet2),
    'MAKE': [None] * len(sheet2),
    'MODEL': [None] * len(sheet2),
    'OWNER': sheet2['Primary Owner'],
    'Hosted Application': sheet2['Application Name'],
    'Physical/Virtual': sheet2['Physical/Virtual'],
    'Business Unit': sheet2['Business entity']
})

# Sheet 3 Mapping (Network)
mapped3 = pd.DataFrame({
    'SR. NO.': range(len(sheet1) + len(sheet2) + 1, len(sheet1) + len(sheet2) + len(sheet3) + 1),
    'Asset Category': ['Network'] * len(sheet3),
    'DEVICE TYPE': sheet3['DEVICE TYPE'],
    'LOCATION': sheet3['LOCATION'],
    'HOST NAME': sheet3['HOSTNAME'],
    'IP ADDRESS': sheet3['IP ADDRESS'],
    'OS': [None] * len(sheet3),
    'Environment': [None] * len(sheet3),
    'STATUS': sheet3['STATUS'],
    'MAKE': sheet3['MAKE'],
    'MODEL': sheet3['MODEL'],
    'OWNER': [None] * len(sheet3),
    'Hosted Application': [None] * len(sheet3),
    'Physical/Virtual': ['Physical'] * len(sheet3),
    'Business Unit': [None] * len(sheet3)
})

# Combine all data
final_df = pd.concat([mapped1, mapped2, mapped3], ignore_index=True)[final_columns]

# Export to Excel
final_df.to_excel("merged_assets_with_business_unit.xlsx", index=False)
