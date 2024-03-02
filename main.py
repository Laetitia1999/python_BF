from datetime import datetime

import pandas as pd

import repository
import view
import toml


def main():
    # Chargement de la configuration depuis config.toml
    config = toml.load('config.toml')
    # Obtenez le chemin à partir de la configuration
    input_path = config["input"]["input_path"]
    bq_sheet = config['input']['sheet_bank']
    erp_sheet = config['input']['sheet_server']
    tableau_sheet = config['input']['sheet_data']
    column_status = config['input']['column_status']
    column_amount = config['input']['column_amount']

    # get data from excel_file
    df_bq = pd.read_excel(input_path, sheet_name=config["input"]["sheet_bank"])
    df_erp = pd.read_excel(input_path, sheet_name=config["input"]["sheet_server"])
    excel = pd.read_excel(input_path, sheet_name=config["input"]["sheet_data"])
    excel['Date'] = excel['Date'].apply(lambda x: datetime.strftime(x, "%B"))

    column_status = 'Status'
    column_amount = 'Montant1'

    # format and merge dataframes
    df_bq = repository.reformat_file(df_bq)
    df_erp = repository.reformat_file(df_erp)
    inner_join = pd.merge(df_erp, df_bq, on='Numéro', how='inner')
    repository.get_status(df=inner_join, column_status=column_status)

    # Open excel workbook and print the status in sheet BQ
    view.print_data(input_path=input_path, df=inner_join, column_status=column_status)

    # get invoices and amounts for which a reminder e-mail must be sent
    final_list = repository.get_list_reminder(df_erp)

    # sent reminder email
    view.mail(config['mail']['smtp_server'], config['mail']['sender_email'], config['mail']['sender_password'], config['mail']['recipient_email'], config['mail']['mail_subject'], config['mail']['mail_body'], final_list)

    # creating the necessary dataframe for graphic and table
    table1 = repository.create_tables(excel, column_amount, column_status)[0]
    table2 = repository.create_tables(excel, column_amount, column_status)[1]

    # generate chart and pdf
    view.graph(config, table1, table2)

    # j'ai ajouté ça pour la partie graph
    for graph_config in config['graphics']:
        if graph_config['id'] == 1:
            view.graph(table1, graph_config['options'])
        elif graph_config['id'] == 2:
            view.graph(table2, graph_config['options'])



