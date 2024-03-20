# -*- coding: utf-8 -*-
"""
Created on Wed Aug  9 11:43:06 2017
@author: thomas.dunn
"""

import pandas as pd

# Define file paths
file1 = 'AL_GE_20180203.txt'
file2 = 'EOB_GE_20180203.txt'
file3 = 'MonthlyReport_GE_20180203.txt'
file4 = 'Accum_GE_20180203.txt'

# Read data from text files into DataFrames
df1 = pd.read_table(file1, dtype=object, names=['NEC #', 'Client Name', 'Pharmacy NPI', 'Pharmacy Name', 'Billing Date', 'TP & Copays', 'Dispense Fees', 'Admin Fees'])
df2 = pd.read_table(file2, dtype=object, names=['NEC #', 'Client Name', 'PharmacyName', 'PharmacyNPI', 'ClaimClassificationName', 'BIN', 'PCN', 'GroupID', 'PrescriberName', 'PrescriberNPI', 'PatientName', 'PatientDOB', 'DateOfDispense', 'PrescriptionNumber', 'FillNumber', 'NDCNumber', 'DrugName', 'Quantity', 'PercentageReplenished', 'TotalPaid', 'DispenseFee', 'PlanReceipts', 'InvoiceDate', 'InvoiceNumber', 'DateOfPurchase', '340BCost'])
df3 = pd.read_table(file3, dtype=object, names=['NEC #', 'Client Name', 'Pharmacy Name', 'Facility Name', 'PO Number', 'Invoice Number', 'Date of Purchase', 'NDC Number', 'NDC Description', 'Quantity Ordered', 'Bottle Size', 'Extended Cost'])
df4 = pd.read_table(file4, dtype=object, names=['NEC #', 'Client Name', 'Pharmacy Name', 'Pharmacy NPI', 'NDC Number', 'NDC Description', 'Pills Accumulated', 'Bottle Size', 'Packages Available To Order'])

# Data processing and formatting
for df in [df1, df2, df3, df4]:
    df[['TP & Copays', 'Dispense Fees', 'Admin Fees']] = df[['TP & Copays', 'Dispense Fees', 'Admin Fees']].apply(pd.to_numeric)
    df[['TP & Copays', 'Dispense Fees', 'Admin Fees']] = df[['TP & Copays', 'Dispense Fees', 'Admin Fees']].round(2)

# Additional processing for specific DataFrames
df1['340B Savings'] = df1['TP & Copays'] + df1['Dispense Fees'] + df1['Admin Fees']
df3['Extended Cost'] = df3['Extended Cost'].apply(pd.to_numeric)
df3['Extended Cost'] = df3['Extended Cost'].round(2)

# Extract unique client names
df1clients = pd.DataFrame(df1['Client Name']).drop_duplicates()
df2clients = pd.DataFrame(df2['Client Name']).drop_duplicates()
df3clients = pd.DataFrame(df3['Client Name']).drop_duplicates()
df4clients = pd.DataFrame(df4['Client Name']).drop_duplicates()

# Concatenate client names into a single DataFrame
dftotalclients = df1clients.append(df2clients).append(df3clients).append(df4clients).drop_duplicates()
listofnames = dftotalclients['Client Name'].tolist()

# Generate Excel files for each client
for name in listofnames:
    excelname = ''.join([str(name), '.xlsx'])
    writer = pd.ExcelWriter(excelname, engine='xlsxwriter')
    workbook = writer.book
    for i, df in enumerate([df1, df2, df3, df4], start=1):
        dfprint = df.loc[df['Client Name'] == name]
        dfprint.set_index('NEC #', inplace=True)
        dfprint.to_excel(writer, sheet_name=['AL', 'EOB', 'Order', 'Accum'][i - 1])
        money_fmt = workbook.add_format({'num_format': '$###,###,##0.00', 'bold': False})
        pct_fmt = workbook.add_format({'num_format': '0%'})
        worksheet = writer.sheets[['AL', 'EOB', 'Order', 'Accum'][i - 1]]
        worksheet.set_column(0, 30, 30)
        if i == 2:
            worksheet.set_column('S:S', 20, pct_fmt)
        else:
            worksheet.set_column('F:I', 12, money_fmt)
    writer.save()
    print("Generated", excelname)

# Generate a summary Excel file
excelname = 'totals.xlsx'
writer = pd.ExcelWriter(excelname, engine='xlsxwriter')
workbook = writer.book
for i, df in enumerate([df1, df2, df3, df4], start=1):
    df.set_index('NEC #', inplace=True)
    df.to_excel(writer, sheet_name=['AL', 'EOB', 'Order', 'Accum'][i - 1])
    money_fmt = workbook.add_format({'num_format': '$###,###,##0.00', 'bold': False})
    pct_fmt = workbook.add_format({'num_format': '0%'})
    worksheet = writer.sheets[['AL', 'EOB', 'Order', 'Accum'][i - 1]]
    worksheet.set_column(0, 30, 30)
    if i == 2:
        worksheet.set_column('S:S', 20, pct_fmt)
    else:
        worksheet.set_column('F:I', 12, money_fmt)
writer.save()
print("Generated", excelname)
print("Complete")
