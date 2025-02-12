import time  

def formatar_tempo(segundos):
    horas = int(segundos // 3600)
    minutos = int((segundos % 3600) // 60)
    segundos = int(segundos % 60)
    return f"{horas:02}:{minutos:02}:{segundos:02}"

def medir_tempo(func):
    def wrapper(*args, **kwargs):
        inicio = time.perf_counter()
        resultado = func(*args, **kwargs)
        fim = time.perf_counter()
        tempo_execucao = fim - inicio
        tempo_formatado = formatar_tempo(tempo_execucao)
        print(f"Tempo de execução da função {func.__name__}: {tempo_formatado}")
        return resultado
    return wrapper


