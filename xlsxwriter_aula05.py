import xlsxwriter
from math import*


arquivo = xlsxwriter.Workbook('Aula_06.xlsx')
planilha = arquivo.add_worksheet('Planilha 1')

# Criando o Roll de dados 
entrada = [1.60,1.69,1.72,1.73,1.73,1.74,1.75,1.75,1.75,1.75,1.75,1.76,1.78,1.80,1.82,1.82,1.84,1.88]
#entrada = input('Digite o valor: ')
#dados.append(entrada)
#
#while True:
#    if entrada == 'pare':
#        break
#    else:
#        entrada = input('Digite o valor: ')
#        dados.append(entrada)
#
#dados.pop()
#entrada = [float(i) for i in dados]
entrada.sort()
print(entrada)

#Tamanho da lista
n = len(entrada)
#Regra de Sturges
k = 1 + 3.3*log10(n)
#Critério de arredondamento
decimal = k - int(k)
if decimal <= 0.5:
    k = int(k)
else:
    k = ceil(k)

# NO EXCEL

#Roll de dados
border1 = arquivo.add_format({'border': 1, 'center_across':True})
planilha.write('A1', 'Roll',border1)

linha = 1
coluna = 0
for numeros in entrada:
    planilha.write(linha, coluna, numeros,border1)
    linha += 1

# Nomes 
border = arquivo.add_format({'border': 1})
planilha.set_column('C:C',20)
planilha.write('C5', 'Total de Dados',border)
planilha.write('C6', 'Valor Mínimo',border)
planilha.write('C7', 'Valor Máximo',border)
planilha.write('C8', 'Número de Classes',border)
planilha.write('C9', 'Amplitude do Intervalo',border)
planilha.write('C10','Amplitude da Classe',border)

# Resultado
planilha.write('D5', n,border)
planilha.write('D6', '=MIN(A2:A'+str(n+1)+')',border)
planilha.write('D7', '=MAX(A2:A'+str(n+1)+')',border)
planilha.write('D8', k,border)
planilha.write('D9', '=D7-D6',border)
planilha.write_formula('D10','=ROUNDUP(D9/D8,2)',border)

#Construindo a Tabela
border1 = arquivo.add_format({'bottom':6, 'top':6, 'center_across':True})
planilha.merge_range('F1:H1', 'Classe',border1)
planilha.set_column('I:I',10)
planilha.write('I1', 'Frequência',border1)
planilha.set_column('J:J',15)
planilha.write('J1', 'Freq. Acumulada', border1)

#Calculando os limites inferiores e superiores
linha = 2 #a partir de F3
coluna = 5 #coluna F
planilha.write('F2', '=MIN(A2:A'+str(n+1)+')-0.01')
for i in range(linha, k+1):
    planilha.write_formula(linha, coluna, '=F'+str(i)+'+$D$10')
    planilha.write_formula(linha-1, coluna+2, '=F'+str(i)+'+$D$10')
    planilha.write(linha-1, coluna+1, '|--')
    linha +=1
planilha.write_formula(linha-1, coluna+2, '=F'+str(i+1)+'+$D$10')
planilha.write(linha-1, coluna+1, '|--')
planilha.set_column('G:G',5)

planilha.merge_range('F'+str(linha+1)+':H'+str(linha+1), 'Total', border1)

#Frequência Absoluta
linha = 1
planilha.write_array_formula(linha, coluna+3, linha+k-1, coluna+3, '=FREQUENCY(A2:A'+str(n+1)+',H'+str(linha+1)+':H'+str(k)+')')
planilha.write_formula(linha+k, coluna+3, '=SUM(I'+str(linha+1)+':I'+str(k+1)+')')

#Frequênica Acumulada
for i in range(linha, k+1):
    planilha.write_formula(linha, coluna+4, '=SUM($I$2:I'+str(i+1)+')')
    linha+=1




arquivo.close()