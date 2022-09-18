import win32com.client

mailto = input('Para quem deseja enviar?')
sub = input('CC;')
ht = input('Digite sua mensagem normalmente ou em HTML')
outlook = win32com.client.Dispatch("Outlook.Application")
conta = ''


for item in outlook.Session.Accounts:
    # coloque o seu email no lugar remetente@email.com
    if item.SmtpAddress == 'marcusvcunha0800@hotmail.com':
        conta = item
        break


mail = outlook.CreateItem(0)


if conta:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, conta))


mail.To = mailto


mail.Subject = 'Email vindo do Outlook'

#mail.CC = 'email@gmail.com'

#mail.BCC = 'email@gmail.com'

# se preferir pode mandar um email mais simples usando o mail.Body
mail.HTMLBody = ht

# Anexos (pode colocar quantos quiser =D):

mail.Send()
