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

    PATH_OCORRENCIAS = PATH_DOWNLOADS / 'ocorrenciasSF3.xlsx'
    PATH_HISTORICO = PATH_DOWNLOADS / "historico_processados.csv"

    NOME_DA_ABA_ALVO = f"{mes_upper} {ano}"
    COLUNA_DATA_REAL = 'DATA'
    COLUNA_CONTRATO = 'CONTRATO'
    COLUNA_COD_OCORRENCIA = 'COD.OCORRENCIA'
    COLUNA_COMPLEMENTO = 'COMPLEMENTO 2'

    # histórico
    if PATH_HISTORICO.exists() and PATH_HISTORICO.stat().st_size > 0:
        df_historico = pd.read_csv(PATH_HISTORICO, dtype=str)
    else:
        df_historico = pd.DataFrame(
            columns=["CHAVE", "Contrato", "Data_Ocorrencia", "Cod_Ocorrencia", "Complemento", "Data_Envio"]
        )

    print("Linhas no histórico:", len(df_historico))

    # le excel
    df_ocorrencias = pd.read_excel(
        PATH_OCORRENCIAS,
        sheet_name=NOME_DA_ABA_ALVO,
        header=0,
        engine="openpyxl"
    )

    print("Linhas lidas do excel (antes de remover duplicatas exatas):", len(df_ocorrencias))
    df_ocorrencias = df_ocorrencias.drop_duplicates(keep='first').reset_index(drop=True)
    print("Linhas após remover duplicatas exatas:", len(df_ocorrencias))

    # filtra data
    df_ocorrencias.dropna(subset=[COLUNA_DATA_REAL], inplace=True)
    df_ocorrencias[COLUNA_DATA_REAL] = pd.to_datetime(
        df_ocorrencias[COLUNA_DATA_REAL],
        errors='coerce'
    )
    print("Data filtrada com sucesso.")

    # normaliza colunas usadas na chave (evita erro se vier NaN)
    df_ocorrencias[COLUNA_CONTRATO] = df_ocorrencias[COLUNA_CONTRATO].astype(str)
    df_ocorrencias[COLUNA_COD_OCORRENCIA] = df_ocorrencias[COLUNA_COD_OCORRENCIA].astype(str)
    df_ocorrencias[COLUNA_COMPLEMENTO] = df_ocorrencias[COLUNA_COMPLEMENTO].fillna('').astype(str)

    # Data_Key inclui data E hora (se a coluna DATA tiver hora, isso já diferencia
    # ocorrências do mesmo contrato/dia automaticamente)
    df_ocorrencias["Data_Key"] = df_ocorrencias[COLUNA_DATA_REAL].astype(str)

    # CHAVE única por OCORRÊNCIA, não só por contrato+dia
    df_ocorrencias["CHAVE"] = (
        df_ocorrencias[COLUNA_CONTRATO] + "_"
        + df_ocorrencias["Data_Key"] + "_"
        + df_ocorrencias[COLUNA_COD_OCORRENCIA] + "_"
        + df_ocorrencias[COLUNA_COMPLEMENTO]
    )

    # remove qualquer coisa que já esteja no histórico (mesma CHAVE completa)
    ocorrencias_novas = df_ocorrencias[
        ~df_ocorrencias["CHAVE"].isin(df_historico.get("CHAVE", []))
    ].reset_index(drop=True)

    print("Ocorrências novas encontradas:", len(ocorrencias_novas))

    # gera excel
    path_excel = PATH_DOWNLOADS / "ocorrencias_filtradas.xlsx"
    ocorrencias_novas.drop(columns=["Data_Key", "CHAVE"]).to_excel(path_excel, index=False)
    print("Arquivo Excel gerado com sucesso.")

    # formata data p chata da sabrina
    df_html = ocorrencias_novas.drop(columns=["Data_Key", "CHAVE"]).copy()
    df_html['DATA'] = df_html['DATA'].dt.strftime('%d/%m/%Y')
    print("Data formatada para a Sabrina.")

    # salva no historico exatamente o que foi enviado, chave completa
    if not ocorrencias_novas.empty:
        df_novos = pd.DataFrame({
            "CHAVE": ocorrencias_novas["CHAVE"],
            "Contrato": ocorrencias_novas[COLUNA_CONTRATO].astype(str),
            "Data_Ocorrencia": ocorrencias_novas[COLUNA_DATA_REAL].dt.date.astype(str),
            "Cod_Ocorrencia": ocorrencias_novas[COLUNA_COD_OCORRENCIA].astype(str),
            "Complemento": ocorrencias_novas[COLUNA_COMPLEMENTO].astype(str),
            "Data_Envio": hoje.strftime("%Y-%m-%d")
        })

        # aqui o drop_duplicates é seguro: só remove se a OCORRÊNCIA inteira for igual
        df_novos = df_novos.drop_duplicates(subset=["CHAVE"])

        df_historico = pd.concat([df_historico, df_novos], ignore_index=True)
        df_historico.to_csv(PATH_HISTORICO, index=False)

        print("Histórico atualizado com sucesso. Novas linhas gravadas:", len(df_novos))
    else:
        print("Nenhuma ocorrência nova para salvar no histórico.")

    return df_html, path_excel