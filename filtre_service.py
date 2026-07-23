import pandas as pd
from pathlib import Path
import datetime
from data_utils import obter_mes_ano
mes, mes_upper, ano = obter_mes_ano()

def filter_ocorrencias():
    hoje = datetime.date.today()

    PATH_CWD = Path.cwd()
    PATH_DOWNLOADS = PATH_CWD / 'src' / 'downloads'
    PATH_DOWNLOADS.mkdir(parents=True, exist_ok=True)

    PATH_OCORRENCIAS = PATH_DOWNLOADS / 'ocorrenciasSF3.xls'
    PATH_HISTORICO = PATH_DOWNLOADS / "historico_processados.csv"

    meses_pt = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO",
    4: "ABRIL", 5: "MAIO", 6: "JUNHO",
    7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO",
    10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
    }

    mes_atual = meses_pt[hoje.month]
    ano_atual = hoje.year

    NOME_DA_ABA_ALVO = f"{mes_upper} {ano}"
    COLUNA_DATA_REAL = 'DATA'
    COLUNA_CONTRATO = 'CONTRATO'

    # histórico
    if PATH_HISTORICO.exists() and PATH_HISTORICO.stat().st_size > 0:
        df_historico = pd.read_csv(PATH_HISTORICO)
    else:
        df_historico = pd.DataFrame(
            columns=["Contrato", "Data_Ocorrencia", "Data_Envio"]
        )

    print("Linhas no histórico:", len(df_historico))

    #le excel
    df_ocorrencias = pd.read_excel(
        PATH_OCORRENCIAS,
        sheet_name=NOME_DA_ABA_ALVO,
        header=0
    )

    #filtra data
    df_ocorrencias.dropna(subset=[COLUNA_DATA_REAL], inplace=True)
    df_ocorrencias[COLUNA_DATA_REAL] = pd.to_datetime(
        df_ocorrencias[COLUNA_DATA_REAL],
        errors='coerce'
    )
    print("Data filtrada com sucesso.")

    #criando a chave
    df_ocorrencias[COLUNA_CONTRATO] = df_ocorrencias[COLUNA_CONTRATO].astype(str)

    df_ocorrencias["Data_Key"] = (
        df_ocorrencias[COLUNA_DATA_REAL].dt.date.astype(str)
    )

    df_ocorrencias["CHAVE"] = (
        df_ocorrencias[COLUNA_CONTRATO] + "_" + df_ocorrencias["Data_Key"]
    )

    if not df_historico.empty:
        df_historico["CHAVE"] = (
            df_historico["Contrato"].astype(str)
            + "_"
            + df_historico["Data_Ocorrencia"].astype(str)
        )

    #filtro finalizado
    ocorrencias_novas = df_ocorrencias[
        ~df_ocorrencias["CHAVE"].isin(df_historico.get("CHAVE", []))
    ]

    contagem = (ocorrencias_novas.groupby([COLUNA_CONTRATO, "Data_Key"]).size())

    chaves_validas = contagem[contagem == 1].index

    ocorrencias_filtradas = ocorrencias_novas[
        ocorrencias_novas.set_index([COLUNA_CONTRATO, "Data_Key"])
        .index.isin(chaves_validas)
    ].reset_index(drop=True)


    print("Contratos filtrados com sucesso.")

    ocorrencias_para_historico = ocorrencias_novas.copy()

    #gera excel
    path_excel = PATH_DOWNLOADS / "ocorrencias_filtradas.xlsx"
    ocorrencias_filtradas.to_excel(path_excel, index=False)
    print("Arquivo Excel gerado com sucesso.")

    #formata data p chata da sabrina
    df_html = ocorrencias_filtradas.copy()
    df_html['DATA'] = df_html['DATA'].dt.strftime('%d/%m/%Y')
    print("Data formatada para a Sabrina.")

    #salva no historico os contratos enviados
    if not ocorrencias_para_historico.empty:
        df_novos = pd.DataFrame({
            "Contrato": ocorrencias_para_historico[COLUNA_CONTRATO].astype(str),
            "Data_Ocorrencia": ocorrencias_para_historico[COLUNA_DATA_REAL].dt.date.astype(str),
            "Data_Envio": hoje.strftime("%Y-%m-%d")
        })

        df_novos = df_novos.drop_duplicates(subset=["Contrato", "Data_Ocorrencia"])

        df_historico = pd.concat([df_historico, df_novos], ignore_index=True)
        df_historico.to_csv(PATH_HISTORICO, index=False)

        print("Histórico atualizado com sucesso.")
    else:
        print("Nenhuma ocorrência nova para salvar no histórico.")


    return df_html, path_excel
