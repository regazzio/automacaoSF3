from pathlib import Path
from filtre_service import filter_ocorrencias
from ftp_service import down_ocorrencias
from email_service import generate_html, send_email
import schedule
import time
import pandas as pd
from datetime import datetime


def main():
    print(f"Executando em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    down_ocorrencias()

    downloads = Path.cwd() / "src" / "downloads"
    path_historico = downloads / "historico_processados.csv"

    # linhas do histórico ANTES de processar (para auditoria)
    if path_historico.exists() and path_historico.stat().st_size > 0:
        linhas_historico_antes = len(pd.read_csv(path_historico, dtype=str))
    else:
        linhas_historico_antes = 0

    df, anexo = filter_ocorrencias()

    # linhas do histórico DEPOIS de processar
    if path_historico.exists() and path_historico.stat().st_size > 0:
        linhas_historico_depois = len(pd.read_csv(path_historico, dtype=str))
    else:
        linhas_historico_depois = 0

    print("\n===== RESUMO DA EXECUÇÃO =====")
    print(f"Ocorrências enviadas no e-mail hoje: {len(df)}")
    print(f"Linhas no histórico antes: {linhas_historico_antes}")
    print(f"Linhas no histórico depois: {linhas_historico_depois}")
    print(f"Novas linhas gravadas no histórico: {linhas_historico_depois - linhas_historico_antes}")

    if len(df) != (linhas_historico_depois - linhas_historico_antes):
        print("ATENÇÃO: quantidade enviada no e-mail é diferente da quantidade gravada no histórico. Verificar manualmente.")
    print("================================\n")

    if df.empty:
        html = "<p><strong>Não há ocorrências novas para o dia de hoje.</strong></p>"
    else:
        html = generate_html(df)

    print("\nPré-visualização pronta.")

    send_email(html, anexo)

    # limpando arquivos downloads
    for arquivo in downloads.glob("*"):
        if arquivo.name != "historico_processados.csv":
            try:
                arquivo.unlink()
                print(f"Removido: {arquivo.name}")
            except Exception as e:
                print(f"Erro ao remover {arquivo.name}: {e}")

    print("Execução finalizada com sucesso.\n"
    "Aguardando para próxima execução.")


if __name__ == "__main__":
    schedule.every().monday.at("12:00").do(main)
    schedule.every().tuesday.at("19:56").do(main)
    schedule.every().wednesday.at("13:14").do(main)
    schedule.every().thursday.at("13:31").do(main)
    schedule.every().friday.at("12:00").do(main)

    print("Agendador iniciado. Aguardando execução às 12:00...")
    
    while True:
        schedule.run_pending()
        time.sleep(5)