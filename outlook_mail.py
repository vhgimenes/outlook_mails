"""
Author: Victor Gimenes
Date: 13/10/2021
Módulo criado para armazenar funções relacionadas da criação e envio de emails via outlook app utilizando python e html.
"""

# Importando módulos externos
import win32com.client
import os 

# Funções locais
def send_email(mail_to:str, mail_subject:str, mail_html:str, send:int, mail_cc:str=None, path: str=None):
    """
    Função responsável automtizar o envio de emails via outlook app.
    
    Obs.: outlook precisar estar logado e aberto no computador.

    Args:
        mail_to (str): remetentes para o email (; como delimitador).
        mail_subject (str): assunto do email.
        mail_html (str): corpo do email, em html.
        send (int): deve ser enviado (1) ou somento plotado no outlook (0).
        path (list): lista contendo o path de todas os arquivos que devem ser colocados em anexo.
    """
    # Criando a conexão com o app do outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    # Criando o e-mail
    mail = outlook.CreateItem(0) 
    # Adicionando os e-maisl dos destinatários
    mail.To = mail_to # you can add multiple emails with the ; as delimiter. E.g. test@test.com; test2@test.com;
    # Adicionando os e-mails dos destinatário que devem estar em cópia (se necessário)
    if not isinstance(mail_cc,type(None)):
        # Msg.CC = "test@test.com"
        mail.CC = mail_cc
    # Adicionando assunto no e-mail
    mail.Subject = mail_subject
    # Hack para adicionar a assinuta do remetente junto com o corpo do e-mail
    mail.GetInspector 
    signature = mail.HTMLBody 
    mail.HTMLBody = mail_html + signature
    # Adicionando os arquivos que devem estar em anexo 
    if path!=None:
        for i in path: mail.Attachments.Add(i)
    # Enviando E-mail
    if send == 1:
        # Envio automático
        mail.Send()
    elif send == 0:
        # Display no app sem envio automático
        mail.Display()
    
