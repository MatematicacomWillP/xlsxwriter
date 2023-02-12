import xlsxwriter

#== CRIANDO A NOSSA PRIMEIRA TABELA ==#

arquivo = xlsxwriter.Workbook('Aula_04.xlsx')
planilha = arquivo.add_worksheet('Planilha 1')

formatacao = arquivo.add_format({'bold':True, 'center_across':True, 'border':1})
borda = arquivo.add_format({'border':1})

planilha.write('A1', 'Nome', formatacao)
planilha.write('B1', 'Idade', formatacao)

dados = (['Nome1', 18], ['Nome2', 20], ['Nome3', 25])

linha = 1
coluna = 0

for nome, idade in dados:
    planilha.write(linha, coluna, nome, borda)
    planilha.write(linha, coluna + 1, idade, borda)
    linha += 1

planilha.write(linha, 0, "Total", borda)
planilha.write(linha, 1, '=SUM(B1:B'+str(linha)+')', borda)


arquivo.close()