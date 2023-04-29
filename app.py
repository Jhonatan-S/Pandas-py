
import pandas as pd
import time


df = pd.read_excel('table.xlsx')

monthBirth = input('Digite o mês de animversário: Ex: 05 ')


def loopCheck():
    global monthBirth

    while monthBirth[0] == ' ':
        monthBirth = input('Digite o mês de animversário sem espaço: Ex: 05 ')
        return loopCheck()

    while len(monthBirth) > 2 :
        if monthBirth[0] == ' ' or  monthBirth[1] == ' ' or monthBirth[2] or len(monthBirth) > 2:
            monthBirth = input('Digite o mês de animversário sem espaço: Ex: 05 ')
            return loopCheck()

    while len(monthBirth) < 2:
        monthBirth = input('Digite o mês de aniversário no formado exemplificado: Ex: 05 ')
        return loopCheck()

loopCheck()

for i, birthDay in enumerate(df['Data de Nascimento']):
    birthDay = str(birthDay)
    name = df.loc[i, 'Nome']
    tell = df.loc[i, 'Telefone']
    email = df.loc[i, 'Email']


    if len(birthDay) > 5 and birthDay[3:5] == str(monthBirth):
        print(f'Nome: {name} \nTel: {tell} \nEmail: {email} \nData de aniversário: {birthDay}' )
        print()
        print()

        writer = pd.ExcelWriter('Área de Trabalho/tabela.xlsx')

        dates = [[name, birthDay, tell, email]]

        

loop = 0
while True:
    loop = loop + 1

    outPut = pd.DataFrame({'Loop': [loop]})

    df.to_excel(writer, dates, index= False)
    writer.save()
    if loop>=32:
        break

writer.save()