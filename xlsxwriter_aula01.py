import xlsxwriter

arquivo = xlsxwriter.Workbook('Aula.xlsx')
planilha = arquivo.add_worksheet('Planilha 1')

formatacao = arquivo.add_format({'bg_color':'black', 'color':'blue','font':'Consolas','font_size':20})

planilha.write(0,0, 'Olá Mundo', formatacao)
planilha.set_column(0,0,20)
arquivo.close()