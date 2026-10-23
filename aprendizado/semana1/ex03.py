# ex03: Verifica se o turno é válido

turno = input("Digite o turno: ")
limpo = turno.strip().lower()

turnos = ["matutino", "vespertino", "noturno", "integral"]

if limpo in turnos:
    print(f"Turno válido: {limpo}")
else:
    print(f"Turno inválido: '{turno}'")