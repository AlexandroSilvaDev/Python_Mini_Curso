import pandas as pd
import win32com.client as win32
# import smtplib
# import email.message

# Importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados

pd.set_option('display.max_columns', None)
# print(tabela_vendas)
# print(tabela_vendas[['ID Loja', 'Valor Final']]) --- Filtro para mostrar apenas as colunas ID Loja e Valor Final
# print(tabela_vendas.groupby('ID Loja').sum()) --- Agrupar as lojas com mesmo nome  e somar o faturamento por loja

# Faturamento por loja

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

# Ticket medio por produto em cada loja

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0 :'Ticket Médio'})
print(ticket_medio)


# Enviar um email com o relatorio

# Enviar email pela web
# def enviar_email():
#     corpo_email = f"""
#     <p>Prezados(as)</p>
#     <p>Segue o Relátorio de vendas por cada loja.</p>
#     <p>Faturamento de cada loja:
#     {faturamento}
#     </p>
#     <p>Quantidade vendida em cada loja:
#     {quantidade}
#     </p>
#     <p>Ticket médio dos produtos em cada loja:
#     {ticket_medio}
#     </p>
#     <p>Quanlquer dúvida estou a disposição.</p>
#     <p>At.te,</p>
#     <p>Alexandro Silva</p>
#     """
#
#     msg = email.message.Message()
#     msg['Subject'] = "Projeto Python mini curso!"
#     msg['From'] = 'xande.silva2503@outlook.com'
#     msg['To'] = 'xande.silva2503@gmail.com'
#     password = '25031997APs'
#     msg.add_header('Content-Type', 'text/html')
#     msg.set_payload(corpo_email)
#
#     s = smtplib.SMTP('smtp.gmail.com: 587')
#     s.starttls()
#
#     # Credenciais de Login para enviar o email
#     s.login(msg['From'], password)
#     s.sendmail(msg['From'], msg['To'], msg.as_string().encode('utf-8'))
#     print('Email enviado')
#
# enviar_email()

# Enviar email pelo aplicativo do outlook no windows

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'xande.silva2503@outlook.com'
mail.Subject = 'Projeto Python mini curso!'
mail.HTMLBody = f'''
    <p>Prezados(as)</p>
    <p>Segue o Relátorio de vendas por cada loja.</p>
    <p>Faturamento de cada loja:
    {faturamento.to_html()}
    </p>
    <p>Quantidade vendida em cada loja:
    {quantidade.to_html()}
    </p>
    <p>Ticket médio dos produtos em cada loja:
    {ticket_medio.to_html()}
    </p>
    <p>Quanlquer dúvida estou a disposição.</p>
    <p>At.te,</p>
    <p>Alexandro Silva</p>
 '''

mail.Send()

print('Email enviado!')
