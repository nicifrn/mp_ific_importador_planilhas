# ex07: Calcula maior, menor, soma e média

cargas = [160, 40, 200, 80, 20, 160]

maior = cargas[0]
menor = cargas[0]
soma = 0

for carga in cargas:
    if carga > maior:
        maior = carga

    if carga < menor:
        menor = carga

    soma += carga

media = soma / len(cargas)

print("Maior:", maior)
print("Menor:", menor)
print("Soma:", soma)
print("Média:", media)

print(max(cargas), min(cargas), sum(cargas))