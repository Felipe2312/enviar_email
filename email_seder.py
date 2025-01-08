import smtplib
from email.message import EmailMessage
from time import sleep
import imghdr
from time import sleep
import streamlit as st
import pandas as pd


class Emailer:
    def __init__(self, email_origem, senha_email, servidor='smtp.gmail.com', porta=465):
        self.email_origem = email_origem
        self.senha_email = senha_email
        self.servidor = servidor
        self.porta = porta


    def enviar_email(self, topico, conteudo_email, df, intervalo_em_segundos=60):
        # Filtra os emails não enviados
        df_nao_enviados = df.query('enviado != enviado')
        quantidade_emails = len(df_nao_enviados['email'])
        
        if quantidade_emails == 0:
            print("Todos os emails já foram enviados!")
            return
        cont = 0
        for index, row in df.iterrows():
            if row['enviado'] != 'enviado':
                email_destino = row['email']
                self.mail = EmailMessage()
                self.mail['Subject'] = topico
                self.mail['From'] = self.email_origem
                self.mail['To'] = email_destino
                self.mail.add_header('Content-Type', 'text/html')
                self.mail.set_payload(conteudo_email.encode('utf-8'))
                
                cont +=1 
                porcentagem_envio = f'{cont / quantidade_emails * 100:.2f}%'

                try:
                    with smtplib.SMTP_SSL(self.servidor, self.porta) as smtp:
                        smtp.login(user=self.email_origem, password=self.senha_email)
                        smtp.send_message(self.mail)
                    
                    df.at[index, 'enviado'] = 'enviado'
                    print(f"E-mail enviado para {email_destino}")
                    print(f"Progresso: {porcentagem_envio}")
                except Exception as e:
                    df.at[index, 'enviado'] = 'erro'
                    print(f"Erro ao enviar e-mail para {email_destino}: {e}")
                    print(f"Progresso: {porcentagem_envio}")

                # Aguarda o intervalo especificado
                sleep(intervalo_em_segundos)
    
    def anexar_imagem(self, lista_imagens):
        for imagem in lista_imagens:
            with open(imagem, 'rb') as arquivo:
                dados = arquivo.read()
                extensao_imagem = imghdr.what(arquivo.name)
                nome_arquivo = arquivo.name
            self.mail.add_attachment(dados, maintype='image',
                                     subtype=extensao_imagem, filename=nome_arquivo)

    def anexar_arquivos(self, lista_arquivos):
        for arquivo in lista_arquivos:
            with open(arquivo, 'rb') as a:
                dados = a.read()
                nome_arquivo = a.name
            self.mail.add_attachment(dados, maintype='application',
                                     subtype='octet-stream', 
                                     filename=nome_arquivo)
