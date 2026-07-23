import datetime

def obter_mes_ano():
    hoje = datetime.date.today()

    meses = {
        1: ("Janeiro", "JANEIRO"),
        2: ("Fevereiro", "FEVEREIRO"),
        3: ("Março", "MARÇO"),
        4: ("Abril", "ABRIL"),
        5: ("Maio", "MAIO"),
        6: ("Junho", "JUNHO"),
        7: ("Julho", "JULHO"),
        8: ("Agosto", "AGOSTO"),
        9: ("Setembro", "SETEMBRO"),
        10: ("Outubro", "OUTUBRO"),
        11: ("Novembro", "NOVEMBRO"),
        12: ("Dezembro", "DEZEMBRO"),
    }

    mes, mes_upper = meses[hoje.month]

    return mes, mes_upper, hoje.year