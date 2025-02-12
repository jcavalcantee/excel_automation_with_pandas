from datetime import datetime
from dateutil.relativedelta import relativedelta
import locale

def retornaMesAtual():
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    mes = datetime.now().strftime('%B')
    return mes.capitalize()

def retornaNumeroMes():
    return datetime.now().month

def retornaMesAnterior():
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    mes_atual = datetime.now()
    mes_anterior = mes_atual - relativedelta(months=1)
    return mes_anterior.strftime('%B').capitalize()
