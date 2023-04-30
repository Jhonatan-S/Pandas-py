import pandas as pd
from openpyxl import Workbook
import openpyxl


df = pd.read_excel('table.xlsx')

# monthBirth = input('Digite o mês de aniversário: Ex: 05 ')

workbook = openpyxl.load_workbook('table.xlsx')
worksheet = workbook.active

# Define o mês desejado
monthBirth = 4

# Cria um novo workbook e seleciona a planilha ativa
new_workbook = openpyxl.Workbook()
new_worksheet = new_workbook.active

# Define o cabeçalho da tabela
header = ['Name', 'Birthday', 'Email', 'Tel']

# Adiciona o cabeçalho à primeira linha da planilha
new_worksheet.append(header)

# def loopCheck():
#     global monthBirth

#     while monthBirth[0] == ' ':
#         monthBirth = input('Digite o mês de aniversário sem espaço: Ex: 05 ')
#         return loopCheck()

#     while len(monthBirth) > 2 :
#         if monthBirth[0] == ' ' or  monthBirth[1] == ' ' or monthBirth[2] or len(monthBirth) > 2:
#             monthBirth = input('Digite o mês de aniversário sem espaço: Ex: 05 ')
#             return loopCheck()

#     while len(monthBirth) < 2:
#         monthBirth = input('Digite o mês de aniversário no formado exemplificado: Ex: 05 ')
#         return loopCheck()

# loopCheck()

# for i, birthDay in enumerate(df['Data de Nascimento']):
#     birthDay = str(birthDay)
#     name = df.loc[i, 'Nome']
#     tel = df.loc[i, 'Telefone']
#     email = df.loc[i, 'Email']

#     if len(birthDay) > 5 and birthDay[3:5] == str(monthBirth):    


#         col = 1
#         row = 1
#         ws.cell(row= row, column= col).value = name
#         ws.cell(row= row, column= col + 1).value = birthDay
#         ws.cell(row= row, column= col + 2).value = email
#         ws.cell(row= row, column= col + 3).value = tel 

#         row += row

for i, birthDay in enumerate(worksheet.iter_rows(min_row=2, values_only=True)):
    # Extrai o mês de cada data de aniversário
    birthMonth = birthDay[2].month
    if birthMonth == monthBirth:
        name = worksheet.cell(row=i+2, column=1).value
        tel = worksheet.cell(row=i+2, column=3).value
        email = worksheet.cell(row=i+2, column=4).value
        new_worksheet.append([name, birthDay[2], email, tel])

        
new_workbook.save('aniversariantes.xlsx')



    


        






        
