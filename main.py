from datetime import datetime

import pandas as pd

import Repository
import View

path = r"C:\Users\laeti\OneDrive\Desktop\Master 2 BF\Autre\projet maquette\account.invoice (1).xlsx"

# get data from excel_file
df_bq = pd.read_excel(path, sheet_name="bq")
df_erp = pd.read_excel(path, sheet_name="ERP")
excel = pd.read_excel(path, sheet_name="tableau")
excel['Date'] = excel['Date'].apply(lambda x: datetime.strftime(x, "%B"))

column_status = 'Status'
column_amount = 'Montant1'

# format and merge dataframes
df_bq = Repository.reformat_file(df_bq)
df_erp = Repository.reformat_file(df_erp)
inner_join = pd.merge(df_erp, df_bq, on='Numéro', how='inner')
Repository.get_status(df=inner_join, column_status=column_status)

# Open excel workbook and print the status in sheet BQ
View.print_data(path=path, df=inner_join, status=column_status)

# get invoices and amounts for which a reminder e-mail must be sent
final_list = Repository.get_list_reminder(df_erp)

# sent reminder email
View.mail("oceane.guiovanna@gmail.com", "lvlk xsbt zjkt dfjh", "Laetitia_sfeir@hotmail.com", "Mail de relance",
          final_list)

# creating the necessary dataframe for graphic and table
table1 = Repository.create_tables(excel, column_amount, column_status)[0]
table2 = Repository.create_tables(excel, column_amount, column_status)[1]

# generate chart and pdf
View.graph(table1, table2)
