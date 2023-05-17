valorEmprestimo=72000 #PV
valorParcela=0 #P
nroParcelas=60 #N
taxaJuros=1.6
vTotalJurosAcumulado = 0
vTotalCapitalAmortizado = 0

i = (taxaJuros/100) #I
tx = (1 + i) ** nroParcelas #I Calculado

txA = tx * i #Parte de cima da formula
txB = tx - 1 #Parte de baixo da formula

txR = txA / txB #Formula Calculada

valorParcela = valorEmprestimo * txR #PV * Formula Calculada = P

print(f'Valor da Parcela {valorParcela}')

for nroP in range(nroParcelas):
    vJuros = valorEmprestimo * i #72000 * 0,016 (Cada Aoperação Diminuir/Somar desse valor aqui)
    vTotalJurosAcumulado = vTotalJurosAcumulado + vJuros #Saber o total Pago de Juros no final do Caluclo
    vCapitalAmortizado = valorParcela - vJuros #Saber quanto do capital foi amortizado Valor da Parcela menos o Juros
    vTotalCapitalAmortizado = vTotalCapitalAmortizado + vCapitalAmortizado #Saber o quanto foi pago do capital no final do calculo
    valorEmprestimo = valorEmprestimo - vCapitalAmortizado #Saber o valor do saldo devedor para calcular o Juros novamente


    print('#########-------############')
    print(f'Juros: {vJuros}')
    print(f'Total Parcela: {valorParcela}')
    print(f'Total Capital Amortizado: {vCapitalAmortizado}')
    print(f'Valor da Dívida atualizado {valorEmprestimo}')
    print('############################')

print('\n\n------------Totais--------------------')
print(f'Total Juros {vTotalJurosAcumulado}')
print(f'Total Amortizado {vTotalCapitalAmortizado}')
print(f'Total Devedor Emprestimo {valorEmprestimo}')
print(f'Total pago {vTotalJurosAcumulado + vTotalCapitalAmortizado}')




