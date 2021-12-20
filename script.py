import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def envio_de_email(msg, subject, destinatary):
    host = 'smtp.gmail.com'
    port = 587
    user = 'devmail.jluza@gmail.com'
    password = 'xrrst45-f12'

    try:
        server = smtplib.SMTP(host, port)
        server.ehlo()
        server.starttls()
        server.login(user, password)
    except Exception as error:
        print('Erro: ', error)
    message = msg
    email_msg = MIMEMultipart()
    email_msg['From'] = user
    email_msg['To'] = destinatary
    email_msg['Subject'] = subject
    print('Adicionando texto...')
    email_msg.attach(MIMEText(message, 'html'))
    # Enviando mensagem
    print('Enviando mensagem...')
    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
    print('Mensagem enviada!')
    server.quit()

data = pd.read_excel("Vendas.xlsx")
faturamento = data[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
quantidades = data[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
ticket = (faturamento['Valor Final'] / quantidades['Quantidade']).to_frame()
ticket = ticket.rename(columns={0 : 'Ticket Médio'})
msg = f'''
<html>
    <head>
        <meta charset="UTF-8">
    </head>
    <body>
        <h1>Relatórios por Loja</h1>
        <p>Prezados, segue os relatórios de faturamento, quantidade de vendas e ticket médio dos produtos por loja:</p>
        <h2>Faturamento:</h2>
        {faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}
        <h2>Quantidade Vendida</h2>
        {quantidades.to_html()}
        <h2>Ticket Médio</h2>
        {ticket.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}
        <p>Qualquer dúvida, seguimos à disposição</p>
        <p>Att.,</p>
        <p>João Luza</p>
    </body>
</html>
'''
envio_de_email(msg, "Relatório", 'joao.m.luza@gmail.com')