#! python3
# Simple script to calculate "prescricao administrativa"

from datetime import date, datetime

def presc_interval(d1, d2):
    date_format = '%d/%m/%Y'
    marco1 = datetime.strptime(d1, date_format)
    if d2 != '':
        marco2 = datetime.strptime(d2, date_format)
    else: 
        marco2 = datetime.now()
    delta = marco2 - marco1
    perguntaTeste = input('Entre inicio do PAD e decisao? ')
    if perguntaTeste == 'sim' or perguntaTeste == 'Sim':
        return (delta.days-140)
    elif perguntaTeste == 'não' or perguntaTeste == 'Não':
        return delta.days
    else: 
        raise ValueError('Resposta incorreta')

def days_to(number_of_days):
    year = int(number_of_days/365)
    plusMonths = int((number_of_days%365)/30)
    plusDays = int((number_of_days%365)%30)
    months = int(number_of_days/30)
    plusDays2 = int(number_of_days%30)
    print('Em anos: {} ano(s), {} mes(es) e {} dia(s).\nEm meses: {} mes(es) e {} dia(s).'.format(year, plusMonths, plusDays, months, plusDays2)) 

#TODO: check against prescricao rates

day1 = input('Digite termo inicial do prazo no formato DD/MM/AAA: ')
day2 = input('Digite termo final do prazo no formato DD/MM/AAA: ')

days_to(presc_interval(day1, day2))


