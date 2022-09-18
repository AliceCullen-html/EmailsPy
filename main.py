# Primeiro importamos as libs
from win32com import client
import pandas as pd


# Lendo o Excel
tabela = pd.read_excel("emails.xlsx")
print(tabela)

# Convertendo a coluna para lista
listando_apenas_1_coluna = list(tabela['EMAILS'])
print(listando_apenas_1_coluna)

# Convertendo a lista para string depois tratando
texto01 = str(listando_apenas_1_coluna)
texto01 = texto01.replace('[', '')
texto01 = texto01.replace(',', ';')
texto01 = texto01.replace(']', '')
print(texto01)


# iniciando a lib win32com
outlook = win32com.client.Dispatch("Outlook.Application")

conta = ''
# aqui colocamos a var com todos os emails que criamos acima
mailto = texto01

sub = ('')

for item in outlook.Session.Accounts:
    # coloque o seu email no lugar remetente@email.com
    if item.SmtpAddress == 'remetente@email.com':
        conta = item
        break


mail = outlook.CreateItem(0)


if conta:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, conta))


mail.To = mailto

mail.Subject = 'Test'

mail.HTMLBody = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Maecenas ultricies lacus orci, sed ultrices ligula pretium eget. Vestibulum quis pharetra ipsum. Vivamus tristique eleifend orci, eu facilisis velit auctor at. Integer iaculis dolor quis massa mollis congue. Mauris porttitor, felis id vestibulum consequat, mi lectus euismod turpis, non euismod purus dolor eu purus. Phasellus semper lorem eu consequat eleifend. Fusce tristique, felis vitae auctor venenatis, tellus sapien facilisis ligula, accumsan finibus urna erat id nibh.'


# Anexos (pode colocar quantos quiser =D):
attachment = r'C:\Users\vulco\Desktop\HASHTAG\python\Email02\emails.xlsx'
mail.Attachments.Add(attachment)


mail.Send()
