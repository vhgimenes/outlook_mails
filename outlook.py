"""
Módulo criado para conter funções relacionadas a automatização do envio de emails com python.
"""

# # Funções Auxiliares
import win32com.client
import os 

def send_email(mail_to:str,mail_subject:str,mail_html:str,send:int,path: str=None):
    """
    Função responsável automtizar o envio de emails via outlook.
    
    Obs.: outlook precisar estar logado e aberto.

    Args:
        mail_to (str): remetentes para o email (; como delimitador)
        mail_subject (str): assunto do email
        mail_html (str): corpo do email
        send (int): deve ser enviado ou somento plopado no outlook
        path (list): lista contendo o path de todas os arquivos que devem ser colocados em anexo
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0) # Email
    mail.To = mail_to # you can add multiple emails with the ; as delimiter. E.g. test@test.com; test2@test.com;
    # Msg.CC = "test@test.com"
    mail.Subject = mail_subject
    mail.GetInspector 
    signature = mail.HTMLBody 
    mail.HTMLBody = mail_html + signature

    # Para colocarmos imagens dentro de HTMLs precisamos colocar essas imagens em anexo 
    if path!=None:
        for i in path: mail.Attachments.Add(i)
    
    #Enviando E-mail
    if send == 1:
        mail.Send()
    elif send == 0:
        mail.Display()
    