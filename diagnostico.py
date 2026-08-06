import pandas as pd
from pathlib import Path
from data_utils import obter_mes_ano

mes, mes_upper, ano = obter_mes_ano()

PATH_OCORRENCIAS = Path.cwd() / "src" / "downloads" / "ocorrenciasSF3.xlsx"
NOME_DA_ABA_ALVO = f"{mes_upper} {ano}"

df = pd.read_excel(PATH_OCORRENCIAS, sheet_name=NOME_DA_ABA_ALVO, header=0, engine="openpyxl")

print("Total de linhas lidas:", len(df))

# marca TODAS as ocorrências de duplicata exata (não só a repetida, a original também)
duplicadas_exatas = df[df.duplicated(keep=False)]
print("\nLinhas 100% idênticas a outra (duplicata exata):", len(duplicadas_exatas))

# agora olha só por CONTRATO + DATA, pra achar "quase duplicatas"
duplicadas_por_contrato_data = df[df.duplicated(subset=["CONTRATO", "DATA"], keep=False)]
print("Linhas com mesmo CONTRATO + DATA (podem ou não ser duplicata real):", len(duplicadas_por_contrato_data))

if len(duplicadas_por_contrato_data) > 0:
    print("\n--- Exemplo de grupo com mesmo CONTRATO + DATA ---")
    exemplo_contrato = duplicadas_por_contrato_data["CONTRATO"].iloc[0]
    print(df[df["CONTRATO"] == exemplo_contrato].to_string())

# checa se tem espaços/whitespace escondido nas colunas de texto
for col in ["CONTRATO", "COD.OCORRENCIA", "DESCR.OCORRENCIA", "COMPLEMENTO 2", "USUARIO", "COD.USUARIO"]:
    if col in df.columns:
        tem_espaco = df[col].astype(str).apply(lambda x: x != x.strip()).sum()
        if tem_espaco > 0:
            print(f"\nATENÇÃO: coluna '{col}' tem {tem_espaco} valores com espaços extras no início/fim.")

print("\nSalvando amostra de duplicatas em duplicatas_encontradas.xlsx para inspeção visual...")
if len(duplicadas_por_contrato_data) > 0:
    duplicadas_por_contrato_data.to_excel(Path.cwd() / "src" / "downloads" / "duplicatas_encontradas.xlsx", index=False)
    print("Arquivo salvo.")
else:
    print("Nenhuma duplicata por CONTRATO+DATA encontrada.")