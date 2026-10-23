# ex05: Calcula as horas acumuladas por semana

horas = int(input("Horas por semana: "))
semanas = int(input("Quantidade de semanas: "))

for i in range(1, semanas + 1):
    print(f"Semana {i}: {horas * i} horas acumuladas")