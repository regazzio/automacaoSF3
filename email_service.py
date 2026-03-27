import win32com.client as win32
from pathlib import Path
import datetime
from dotenv import load_dotenv
import os
load_dotenv()


def generate_html(df):


    print("DEBUG DF:")
    print(df)
    print("Quantidade de linhas:", len(df))
    print("Está vazio?", df.empty)

    if df.empty:
        print("Nenhum contrato encontrado para o dia de hoje.")

        return """
        <p><b>Não foram encontrados contratos que devam ser alterados.</b></p>
        """

    tabela = df.to_html(index=False, border=1, justify="center")

    estilo = """
    <style>
        table {
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 12px;
        }
        th {
            background-color: #f2f2f2;
            padding: 8px;
            white-space: nowrap;
        }
        td {
            padding: 6px;
            white-space: nowrap;
        }
    </style>
    """

    print("Contratos encontrados e tabela HTML gerada.")
    return estilo + tabela


def send_email(html, anexo):
    load_dotenv()
    hoje = datetime.date.today()

    EMAIL_TO = os.getenv("EMAIL_TO")
    EMAIL_CC = os.getenv("EMAIL_CC")

    if not EMAIL_TO:
        raise ValueError("Variável EMAIL_TO não definida")

    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)

    mail.Subject = f"Ocorrências SAC SF3 - {hoje.strftime('%d/%m/%Y')}"
    mail.Display()

    mail.To = EMAIL_TO
    mail.CC = EMAIL_CC or ""

    mail.HTMLBody = f"""
    <p>Boa tarde Camila,</p>

    <p>Segue abaixo o resultado do monitoramento do dia
    <b>{hoje.strftime('%d/%m/%Y')}</b>:</p>

    {html}

    <p>Qualquer dúvida, fico à disposição.</p>
    """ + mail.HTMLBody

    if anexo and Path(anexo).exists():
        mail.Attachments.Add(str(anexo))

    print("Enviando email...")
    mail.Send()
    print("Email enviado com sucesso!")
