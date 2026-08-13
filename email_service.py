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
            margin-bottom: 10px;
        }
        table.ocorrencias {
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 12px;
            width: 100%;
        }
        table.ocorrencias th {
            background-color: #2c3e50;
            color: #ffffff;
            padding: 6px 8px;
            text-align: left;
            border: 1px solid #1a252f;
        }
        table.ocorrencias td {
            padding: 6px 8px;
            border: 1px solid #ddd;
            vertical-align: top;
        }
        td.contrato-cell {
            font-weight: bold;
            background-color: #eef2f5;
            border-right: 2px solid #2c3e50;
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

    cabecalho_colunas = "".join(f"<th>{c}</th>" for c in colunas_exibir)
    linhas_html = []
    cores_fundo = ["#ffffff", "#f7f9fa"]  # alterna cor por grupo de contrato

    for i, (contrato, grupo) in enumerate(df.groupby("CONTRATO", sort=False)):
        qtd = len(grupo)
        cor_fundo = cores_fundo[i % 2]

        for idx, (_, linha) in enumerate(grupo.iterrows()):
            linhas_html.append('<tr style="background-color:%s;">' % cor_fundo)

            if idx == 0:
                linhas_html.append(
                    f'<td class="contrato-cell" rowspan="{qtd}">{contrato}</td>'
                )

            for col in colunas_exibir:
                linhas_html.append(f"<td>{linha[col]}</td>")

            linhas_html.append("</tr>")

    tabela = f"""
    <table class="ocorrencias">
        <thead>
            <tr><th>CONTRATO</th>{cabecalho_colunas}</tr>
        </thead>
        <tbody>
            {''.join(linhas_html)}
        </tbody>
    </table>
    """

    print("Contratos encontrados e tabela HTML agrupada gerada.")
    return estilo + resumo + tabela


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