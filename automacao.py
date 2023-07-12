import pandas as pd
import smtplib
import email.message

#visualizar base de dados
tab_vendas=pd.read_excel('Vendas.xlsx')
print(tab_vendas)
#mostrar todas as colunas da tabela
pd.set_option("display.max_columns", None)
# Faturamento por loja (filtar entre a coluna lojas e agrupa-las,e volar final)
faturamento=tab_vendas[['ID Loja','Valor Final']].groupby("ID Loja").sum()
print(faturamento)
# calcular a quantidade de produto vendido por loja(filtar e agrupar entre lojas e quantidades de produtos vendidos)
quantidade=tab_vendas[["ID Loja","Quantidade"]].groupby("ID Loja").sum()
print(quantidade)
#Calcular ticket médio por loja( faturamento/ para quantidade de produto vendido por loja),
# Nesta caso terei que filtrar informações entre tabelas.
ticket_medio=(faturamento["Valor Final"]/quantidade["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0:"ticket_medio"})
#Automatizar o envio de email personalizado
def enviar_email():
    corpo_email =f"""
    <meta charset="UTF-8">
    <h1>Prezados segue o relatório de vendas do mes</h1>
    
    <p><b> Faturamento</b></p>
   {faturamento.to_html(formatters={"Valor Final":"R${:,.2f}".format})}
    
    
    <p><b>Quantidade</b></p>
    {quantidade.to_html()}
    
    <p><b> ticket Médio</b></p>
    {ticket_medio.to_html(formatters={"Ticket Médio":"R${:,.2f}".format})}
    <p> Qualquer dúvida estou a disposição</p>
    <p> att  Ale &#x1F600 </p> 
    
    """
    msg = email.message.Message()
    msg['Subject'] = "Relatório de vendas  " #"Assunto"
    msg['From'] = "alessandra23333@gmail.com"#'remetente'
    msg['To'] = "alessandra23333@gmail.com"#'destinatario'
    password = 'qrarkzdlqupjqxzw'#senha de app do próprio gmail
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login('alessandra23333@gmail.com', password)
    s.sendmail('alessandra23333@gmail.com', ['alessandra23333@gmail.com'], msg.as_string().encode('utf-8'))
    print('Email elnviado')
    enviar_email
