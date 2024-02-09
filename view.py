import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import matplotlib.pyplot as plt
import xlwings as xw
from matplotlib.backends.backend_pdf import PdfPages


def print_data(config, df):
    """
    Print the status of the dataFrame in the Excel sheet
    :param path: path of the wordbook
    :param df: dataFrame
    :param column_status: column status of the dataframe that need to be printed
    """
    excel_file_path = config['input']['input_path']
    sheet_bank = config['input']['sheet_bank']
    column_status = config['input']['column_status']

    wb = xw.Book(excel_file_path)  # Open the workbook
    sheet = wb.sheets[sheet_bank]  # select sheet
    sheet['G1'].options(index=False).value = df[column_status]  # write status in the file


def mail(config, invoice_list):
    """
    send a reminder to client that didn't pay their invoice
    :param sender: sender email
    :param password: application password
    :param mail_receiver: receiver email
    :param mail_subject: subject of the email
    :param invoice_list: list of tuple with invoice and amount of each client
    :return: sent an email
    """
    smtp_server = config['mail']['smtp_server']
    smtp_port = config['mail']['smtp_port']
    sender_email = config['mail']['sender_email']
    sender_password = config['mail']['sender_password']
    recipient_email = config['mail']['recipient_email']
    mail_subject = config['mail']['mail_subject']
    mail_body = config['mail']['mail_body']

    for invoice, amount in invoice_list:
        # Corps de l'e-mail
        body = f"""\
        Bonjour,

        Je vous informe que votre facture {invoice}, pour un montant de {amount} euros reste à ce jour impayée.

        Je vous pris de bien vouloir vous acquitter du montant de la facture dans un délai de 8 jours, dans le cas \
        contraire
        votre dossier sera transmis à notre service de recouvrement.

        Merci pour votre compréhension.
        """

        # Créez un objet MIMEMultipart
        message = MIMEMultipart()
        message["From"] = str(sender_email)
        message["To"] = str(recipient_email)
        message["Subject"] = str(mail_subject)

        # Attachez le corps au message
        message.attach(MIMEText(body, "plain"))

        # Initialisez la connexion SMTP
        try:
            server = smtplib.SMTP(smtp_server,smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            print("successful connection")
        except Exception as e:
            print("Erreur lors de la connexion au serveur SMTP : ", e)
            exit()

        # Envoyer l'e-mail
        try:
            server.sendmail(sender_email, recipient_email, message.as_string())
            print("E-mail envoyé")
        except Exception as e:
            print("Erreur lors de l'envoi de l'e-mail : ", e)

        # Fermez la connexion SMTP
        server.quit()


def graph(config, table, tab):
    """
    create plots, stack them and save into a pdf file
    :param table: first dataframe to appear in the pdf
    :param tab:  second dataframe to appear in the pdf
    :return: a pdf with the two dataframe and their plots
    """

    pdf_file_path = config['pdf']['title']

    with PdfPages(pdf_file_path) as pdf:
        fig, axes = plt.subplots(4, 1, figsize=(8, 13))
        fig.tight_layout(pad=7.0)  # Adjust layout to prevent clipping of titles

        axes[0].axis('tight')
        axes[0].axis('off')
        axes[0].set_title('Tableau récapitulatif des données')
        axes[0].table(cellText=table.values, colLabels=table.columns, loc='center')

        axes[1].bar(table['Status'], table['count'], color=['tab:blue', 'tab:orange', 'tab:green'])
        axes[1].set_xlabel('Status')
        axes[1].set_ylabel('Nombre de facture')
        axes[1].set_title('Bar chart')

        axes[2].axis('tight')
        axes[2].axis('off')
        axes[2].set_title('Tableau récapitulatif des données par mois')
        axes[2].table(cellText=tab.values, colLabels=tab.columns, loc='center')

        axes[3].bar(tab['Date'], tab['Total'])
        axes[3].set_title("Représentation du CA")

        # Save the current figure to the PDF
        pdf.savefig()
        plt.close()
