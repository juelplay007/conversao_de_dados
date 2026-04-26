import pandas as pd
import smtplib
from email.message import EmailMessage


tabela_vendas = pd.read_excel('Vendas.xlsx')


pd.set_option('display.max_columns', None)
print(tabela_vendas)


faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)


quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)


EMAIL_ORIGEM = "juelplay2605@gmail.com"
EMAIL_SENHA = "bstn gsnh etla lvud"
EMAIL_DESTINO = "pythonimpressionador@gmail.com"

# Criar a mensagem
msg = EmailMessage()
msg["Subject"] = "Relatório de Vendas por Loja"
msg["From"] = EMAIL_ORIGEM
msg["To"] = EMAIL_DESTINO


corpo_html = f"""
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p><strong>Faturamento:</strong></p>
{faturamento.to_html(formatters={'Valor Final': lambda x: f'R${x:,.2f}'})}

<p><strong>Quantidade Vendida:</strong></p>
{quantidade.to_html()}

<p><strong>Ticket Médio dos Produtos em cada Loja:</strong></p>
{ticket_medio.to_html(formatters={'Ticket Médio': lambda x: f'R${x:,.2f}'})}

<p>Qualquer dúvida estou à disposição.</p>


"""

msg.set_content(corpo_html, subtype='html')


try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_ORIGEM, EMAIL_SENHA)
        smtp.send_message(msg)
    print("E-mail enviado com sucesso")
except Exception as erro:
    print(f"❌ Erro ao enviar e-mail: {erro}")