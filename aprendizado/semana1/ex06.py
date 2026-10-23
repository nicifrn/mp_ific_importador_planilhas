# ex06: Organiza cursos em ordem alfabética

cursos = []

while True:
    nome = input("Nome do curso (Enter para terminar): ")

    if nome == "":
        break

    cursos.append(nome)

if len(cursos) == 0:
    print("Nenhum curso digitado")
else:
    print(f"Total de cursos: {len(cursos)}")

    for i, curso in enumerate(sorted(cursos), start=1):
        print(f"{i}. {curso}")