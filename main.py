# import database
import pandas as pd
##import send email
import smtplib
import email.message

# read database
database = pd.read_excel("Vendas.xlsx")

# visualizar a database
##pd.set_option("display.max_columns", None)
##print(database)

# faturamento po loja
faturamento=database[["ID Loja","Valor Final"]].groupby("ID Loja").sum()
##print(faturamento)

# produtos vendidos por loja
vendas=database[["ID Loja","Quantidade"]].groupby("ID Loja").sum()
##print(vendas)

# ticket médio por loja
ticket_medio=(faturamento["Valor Final"]/vendas["Quantidade"]).to_frame("Ticket Médio")
##print(ticket_medio)

# relatório por e-mail
def enviar_email():
    corpo_email = f"""
    <p>Segue o relatório de vendas:</p>
    <p>Faturamento:</p>
    {faturamento.to_html(formatters={"Valor Final":"R${:,.2f}".format})}
    <p>Itens vendidos:</p>
    {vendas.to_html()}
    <p>Ticket Médio:</p>
    {ticket_medio.to_html(formatters={"Ticket Médio":"R${:,.2f}".format})}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Atenciosamente, Felipe</p>
    """

    msg = email.message.Message()
    msg['Subject'] = "Assunto"
    msg['From'] = 'remetente'
    msg['To'] = 'destinatário'
    password = 'senha'#no caso do gmail, precisei configurar a autenticação em 2 fatores e criar uma senha para app, você deverá colocar a senha de 16 dígitos gerada automaticamente... qualquer dúvida acessar esse vídeo: https://www.youtube.com/watch?v=ndxUgivCszE
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()