#!/usr/bin/env python3.8

from openpyxl import load_workbook
caminho = '/media/jeremias/STORAGE/Documentos/Editando Tabelas com Python/Envio de Mensagens via Robo/cancelados.xlsx'
arquivo_excel = load_workbook(caminho)
planilha1 = arquivo_excel.active

numero = int(input("Digite um Número: "))

max_linha = planilha1.max_row
max_coluna = planilha1.max_column
for i in range(1, max_linha + 1):
    for j in range(1, max_coluna + 1):
        if(planilha1.cell(row=i, column=j).value == numero):
            linha = i 
            coluna = j
        
        #print(planilha1.cell(row=i, column=j).value, end=" - ")

#a1 = planilha1['A1']#Nome
#b1 = planilha1['B1']#Telefone
## Imprime o valor da célula A1
#print("Nome: " + a1.value)
#print("Telefone: " + str(b1.value))
##print("A: " + str(b1.value)) 63992332727
        
print("A linha: {}, e a coluna: {}, é onde se encontra o número: {}".format(linha,coluna,numero))
