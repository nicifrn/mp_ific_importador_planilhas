# ex02: Calcula as semanas do curso

horas = int(input("Carga horária total: "))
semana = int(input("Horas por semana: "))

print(f"{horas // semana} semanas completas e {horas % semana} horas na última semana.")