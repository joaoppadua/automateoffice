#Script with function to calculate deadlines

from datetime import date, timedelta

def prazo(ano, mes, dia, num_days):
    """Soma os dias do prazo e dá o termo final
        ano, mes, dia e num_days = integers dados pelo usuário"""
    termo_inicial = date(ano, mes, dia)
    termo_final = termo_inicial + timedelta(num_days)
    return termo_final

ano = int(input("Diga o ano: "))
mes = int(input('Diga o mês: '))
dia = int(input('Diga o dia: '))
num_days = int(input('Diga o prazo em dias: '))

print(prazo(ano, mes, dia, num_days))



