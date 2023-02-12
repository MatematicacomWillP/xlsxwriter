import xlsxwriter

arquivo = xlsxwriter.Workbook('Aula_03.xlsx')
planilha = arquivo.add_worksheet('Planilha 1')

formatacao = arquivo.add_format({'bottom':6,})
planilha.write('B3', None, formatacao)


arquivo.close()