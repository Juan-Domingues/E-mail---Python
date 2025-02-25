import os
import win32com.client as win32

# Caminho do arquivo
caminho_arquivo = r"C:\Users\juan.pablo\Downloads\teste python\blabla.jpg"

# Verificar se o arquivo existe
if not os.path.isfile(caminho_arquivo):
    print(f"Erro: Arquivo não encontrado: {caminho_arquivo}")
else:
    # Criar conexão com o Outlook
    outlook = win32.Dispatch('outlook.application')

    # Criar um novo e-mail
    email = outlook.CreateItem(0)

    # Configurar os detalhes do e-mail
    email.To = "jp291103@gmail.com"
    email.Subject = "Assunto do E-mail com Anexo"
    email.Body = "Olá,\n\nSegue em anexo o arquivo solicitado.\n\nAtenciosamente,\n[Seu Nome]"

    # Anexar o arquivo
    email.Attachments.Add(caminho_arquivo)

    # Enviar o e-mail
    email.Send()
    print("E-mail enviado com sucesso com o anexo!")


