import pandas as pd
import os
excel_file = "data_list.xlsx"
csv_file = "data_list.csv"
#df = pd.read_excel(excel_file)
df = pd.read_csv(csv_file, delimiter=";")
print(df)
column_names = list(df.columns)
#print(list(df.columns))
with open('mail_template.txt') as f:
    template = f.read()

for ind in df.index:
    lines = template
    for col in column_names:
        value = df[col][ind]
        if pd.isna(value):
            value = 'error'
        else:
            value = str(value)

        print(value)

        lines = lines.replace(col, value)
        print("test")

    print(lines)
    mail_address = (df["mail-address"][ind]).replace(".","_")
    filename = f'mail_output/{mail_address}_Bestellnummer_.txt'

    with open(filename, 'w') as f:
        f.write(lines)

