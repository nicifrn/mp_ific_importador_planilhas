# ex04: Verifica a quantidade de vagas

texto = input("Quantidade de vagas: ")

if not texto.isdigit():
    print("Vagas precisa ser um número inteiro")
elif int(texto) == 0:
    print("Curso sem vagas")
elif int(texto) > 40:
    print("Turma grande: confirme o número de vagas")
else:
    print(f"OK: {texto} vagas")