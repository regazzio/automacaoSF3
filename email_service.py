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

    estilo = """
    <style>
        .resumo {
            font-family: Arial, sans-serif;
            font-size: 13px;
            margin-bottom: 12px;
        }
        .contrato-header {
            font-family: Arial, sans-serif;
            margin-top: 18px;
            margin-bottom: 4px;
        }
        .contrato-badge {
            background-color: #2c3e50;
            color: #ffffff;
            padding: 4px 10px;
            border-radius: 4px;
            font-weight: bold;
            font-size: 13px;
        }
        .contrato-qtd {
            color: #555555;
            font-size: 12px;
            margin-left: 6px;
        }
        table.ocorrencias {
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 12px;
            width: 100%;
            margin-bottom: 4px;
        }
        table.ocorrencias th {
            background-color: #f2f2f2;
            padding: 6px 8px;
            text-align: left;
            border: 1px solid #ddd;
        }
        table.ocorrencias td {
            padding: 6px 8px;
            border: 1px solid #ddd;
            vertical-align: top;
        }
    </style>
    """

    total_contratos = df["CONTRATO"].nunique()
    total_ocorrencias = len(df)

    resumo = f"""
    <p class="resumo">
        Total de <b>{total_contratos}</b> contrato(s) com <b>{total_ocorrencias}</b> ocorrência(s) no período.
    </p>
    """

    colunas_exibir = [
        c for c in ["DATA", "COD.OCORRENCIA", "DESCR.OCORRENCIA", "COMPLEMENTO 2", "USUARIO"]
        if c in df.columns
    ]

    blocos_html = []

    # mantém a ordem de aparição no arquivo (sort=False), agrupando por contrato
    for contrato, grupo in df.groupby("CONTRATO", sort=False):
        qtd = len(grupo)
        plural = "ocorrência" if qtd == 1 else "ocorrências"

        cabecalho = f"""
        <div class="contrato-header">
            <span class="contrato-badge">Contrato {contrato}</span>
            <span class="contrato-qtd">{qtd} {plural}</span>
        </div>
        """

        tabela = grupo[colunas_exibir].to_html(
            index=False, border=0, justify="left", classes="ocorrencias", escape=False
        )

        blocos_html.append(cabecalho + tabela)

    print("Contratos encontrados e tabela HTML agrupada gerada.")
    return estilo + resumo + "".join(blocos_html)


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