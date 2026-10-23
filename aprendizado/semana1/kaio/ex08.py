# ex08: Conta turnos válidos e inválidos

turnos = [
    "Noturno", "noturno ", "MATUTINO", "vespertino",
    "Noite", "integral", " Matutino", "noturno"
]

qtd_noturno = 0
qtd_matutino = 0
qtd_vespertino = 0
qtd_integral = 0
invalidos = []

for turno in turnos:
    limpo = turno.strip().lower()

    if limpo == "noturno":
        qtd_noturno += 1
    elif limpo == "matutino":
        qtd_matutino += 1
    elif limpo == "vespertino":
        qtd_vespertino += 1
    elif limpo == "integral":
        qtd_integral += 1
    else:
        invalidos.append(turno)

print("Noturno:", qtd_noturno)
print("Matutino:", qtd_matutino)
print("Vespertino:", qtd_vespertino)
print("Integral:", qtd_integral)
print("Inválidos:", invalidos)