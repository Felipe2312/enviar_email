import smtplib
from email.message import EmailMessage
from time import sleep
import imghdr
import streamlit as st
import pandas as pd
import io


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def disparar_emails(email_origem, senha_email, servidor, porta, topico, conteudo_email, df, intervalo_em_segundos):
    # Filtra os emails não enviados
    df_nao_enviados = df.query('enviado != enviado')
    quantidade_emails = len(df_nao_enviados['email'])
    if quantidade_emails == 0:
        print("Todos os emails já foram enviados!")
        return
    
    placeholder = st.empty()
    progress_container = st.empty()
    for index, row in df.iterrows():
        if row['enviado'] != 'enviado':
            email_destino = row['email']
            mail = EmailMessage()
            mail['Subject'] = topico
            mail['From'] = email_origem
            mail['To'] = email_destino
            mail.add_header('Content-Type', 'text/html')
            mail.set_payload(conteudo_email.encode('utf-8'))
            
            try:
                with smtplib.SMTP_SSL(servidor, porta) as smtp:
                    smtp.login(user=email_origem, password=senha_email)
                    smtp.send_message(mail)
                df.at[index, 'enviado'] = 'enviado'

                #BARRA DE PROGRESSO
                total_items = len(df['email'])
                enviados = len(df.query('enviado == "enviado" or enviado == "erro"')) 
                progresso = (enviados / total_items) if total_items > 0 else 0
                progress_container.progress(progresso, text=f'{progresso*100:.0f}%')
                placeholder.success(f"E-mail enviado para {email_destino}")

                sleep(intervalo_em_segundos)
            except Exception as e:
                df.at[index, 'enviado'] = 'erro'
            
                #BARRA DE PROGRESSO
                total_items = len(df['email'])
                enviados = len(df.query('enviado == "enviado" or enviado == "erro"')) 
                progresso = (enviados / total_items) if total_items > 0 else 0
                progress_container.progress(progresso, text=f'{progresso*100:.0f}%')

                placeholder.warning(f"Erro ao enviar e-mail para {email_destino}: {e}")
    return placeholder.empty(), progress_container.empty()

                 
email_origem = st.text_input(label='E-mail de origem ')
senha_email = st.text_input(label='Senha e-mail', type="password")

col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    servidor = st.text_input(label='servidor',value='smtp.gmail.com')
with col2:
    porta = st.number_input(label='porta', min_value=0, value=465, step=1, format="%d")
with col3:
    intervalo_em_segundos = st.number_input(label='intervalo entre envios', min_value=5, max_value=300, value=45, step=1, format="%d")

topico = st.text_input(label='Topico')
mensagem = st.text_area(label='Mensagem')


uploaded_file = st.file_uploader("Choose a file", type=["xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    if df.columns[0] == 'email' and df.columns[1] == 'enviado':
        st.success("e-mails carregados!")
        st.write(df)

        if st.button("Iniciar disparo"):
            disparar_emails(email_origem, senha_email, servidor, porta, topico, mensagem, df, intervalo_em_segundos)

            st.write('Baixe a sua lista atualizada!')
            excel_file = to_excel(df)
            # Cria o botão de download
            st.download_button(
                label="Baixe a sua lista",
                data=excel_file,
                file_name="dataframe.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",)
            
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.error("Use esse modelo para colocar os e-mail:")
        with col2:
            dados = [
            {'email': 'usuario1@example.com', 'enviado': 'pendente'},
            {'email': 'usuario2@example.com', 'enviado': 'pendente'},
            {'email': 'usuario3@example.com', 'enviado': 'enviado'}]

            modelo_excel = pd.DataFrame(dados)
            # Converte o DataFrame para Excel
            excel_file = to_excel(modelo_excel)
            # Cria o botão de download
            st.download_button(
                label="Baixe o modelo",
                data=excel_file,
                file_name="dataframe.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",)
            