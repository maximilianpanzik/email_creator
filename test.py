import pandas as pd
excel_file = "data_list.xlsx"
#df = pd.read_excel(excel_file)
pd.read_excel(excel_file, index_col=None, header=1)