#FORMATACAO CONDICIONAL COM ICONES
import xlsxwriter as opcoesDoXlsxWriter
import os

#1 - indicando onde será criado o arquivo, seu nome e sua extensão. Importante a questão das barras duplas (testar).
nomeCaminhoArquivo = 'C:\\Users\\Windows\\Desktop\\Python Projetos\\xlsxwriter\FormatacaoCondicionalIcones.xlsx'
planilhaExcel = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)
sheetDados = planilhaExcel.add_worksheet("Dados") #Para renomear o nome da Sheet1 para Dados.

#Aqui iremos inserir as colunas e os valores nas células
inserirDados = [
    ["Coluna 1", "Coluna 2", "Coluna 3", "Coluna 4"],
    [34, 50 ,12, 34],
    [23, 43, 76, 51],
    [43, 29, 34, 12],
    [29, 58 ,73, 19],
    [18, 30, 45, 12],
]

#
sheetDados.write('A1',"Exemplo de formatação condicional com conjunto de ícones")


for linha, range in enumerate(inserirDados):
    sheetDados.write_row(linha + 2, 1, range)

#formatação condicional com ícones
sheetDados.conditional_format('B4:E4', {'type': 'icon_set',
                                        'icon_style': '3_traffic_lights'})

sheetDados.conditional_format('B5:E5', {'type': 'icon_set',
                                        'icon_style': '3_traffic_lights',
                                        'reverse_icons': True})

sheetDados.conditional_format('B6:E6', {'type': 'icon_set',
                                        'icon_style': '3_arrows'})

sheetDados.conditional_format('B7:E7', {'type': 'icon_set',
                                        'icon_style': '4_arrows'})

sheetDados.conditional_format('B8:E8', {'type': 'icon_set',
                                        'icon_style': '5_ratings'})


#3 - Para fechar e salvar as informações
planilhaExcel.close()

#4 - Abrir o arquivo para verificar o resultado
os.startfile(nomeCaminhoArquivo)
