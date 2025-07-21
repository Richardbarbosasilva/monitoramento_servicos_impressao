import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
import re
from datetime import datetime, date, timedelta



# ===========================
# Parâmetros de data e log
# ===========================



hoje = date.today()
data_coleta = (datetime.today() - timedelta(days=3)).strftime('%Y-%m-%d')
caminho_log = f'\\\\arquivosdti.clickip.local\\automacao_dados\\IA_HUBSOFT\\IA_HUBSOFT_PROJECT\\IA\\logs\\logs_clickip\\logs_analise_{data_coleta}.txt'


# ===========================
# Leitura do log e captura dos alertas
# ===========================


alertas = []
nome_cliente = ''

with open(caminho_log, 'r', encoding='utf-8') as f:
    linhas = f.readlines()

i = 0
while i < len(linhas):
    linha = linhas[i].strip()
    if linha.startswith('Anexos únicos analisados de'):
        match = re.search(r"de (.+?) \(id", linha)
        if match:
            nome_cliente = match.group(1).strip()

    elif linha.startswith('ALERTA:'):
        alerta = linha.replace('ALERTA: ', '')
        vendedor = 'NÃO INFORMADO'
        if i + 2 < len(linhas):
            prox = linhas[i + 2].strip()
            if prox.lower().startswith('vendedor:'):
                vendedor = prox.split(':', 1)[1].strip()
        alertas.append((nome_cliente, vendedor, alerta))

    i += 1



# ================================
# Geração do HTML do e-mail
# ================================



# Gera um Content-ID para o logo

logo_cid = make_msgid(domain='clickip.com.br')

html = f"""
<html>
  <head>
    <style>
      body {{
        font-family: Arial, sans-serif;
        color: #333;
        line-height: 1.4;
      }}
      .header {{
        display: flex;
        align-items: center;
        border-bottom: 2px solid #0072C6;
        padding-bottom: 10px;
        margin-bottom: 20px;
      }}
      .header img {{
        height: 50px;
        margin-right: 15px;
      }}
      .title {{
        color: #0072C6;
        font-size: 1.5em;
        margin: 0;
      }}
      .subtitle {{
        font-size: 1em;
        color: #555;
      }}
      table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 15px;
      }}
      th, td {{
        padding: 8px 12px;
        border: 1px solid #ddd;
        text-align: left;
      }}
      th {{
        background-color: #0072C6;
        color: white;
      }}
      tr:nth-child(even) {{
        background-color: #f9f9f9;
      }}
      .footer {{
        margin-top: 25px;
        font-size: 0.85em;
        color: #777;
        border-top: 1px solid #eee;
        padding-top: 10px;
      }}
    </style>
  </head>
  <body>
    <div class="header">
      <img src="cid:{logo_cid[1:-1]}" alt="ClickIP Logo">
      <div>
        <h1 class="title">Clientes Identificados por IA com possíveis irregularidades nos anexos de contrato</h1>
        <p class="subtitle">Data da coleta: {data_coleta}</p>
      </div>
    </div>

    <p>Foram encontrados <strong>{len(alertas)}</strong> alertas durante a análise:</p>

    <table>
      <tr>
        <th>Cliente</th>
        <th>Vendedor</th>
        <th>Descrição do Alerta</th>
      </tr>
      {"".join(f"""
      <tr>
        <td>{cliente}</td>
        <td>{vendedor}</td>
        <td>{desc}</td>
      </tr>""" for cliente, vendedor, desc in alertas)}
    </table>

    <div style="margin-top: 20px; text-align: center;">
      <a href="http://analytics.clickip.local:8282"
         target="_blank" 
         style="background-color: #0072C6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold;">
         Visualizar na plataforma Analytics ClickIP
      </a>
    </div>

    <p class="footer">
      Este é um e-mail automático enviado pelo sistema de análise da ClickIP.  
      Em caso de dúvidas, contate dti@clickip.com.br.
    </p>
  </body>
</html>
"""



# ================================
# Envio do e-mail com SMTP e logo inline
# ================================



Clickip = 'Clickip'

msg = EmailMessage()
msg['Subject'] = f'📊 Relatório de Análise de Novos Clientes Hubsoft {Clickip} – Alertas'
msg['From'] = 'ia_no-reply@clickip.com.br'
msg['To'] = 'richard.silva@clickip.com.br'
msg['Cc'] = 'kethlen.santana@clickip.com.br'
#msg.set_content("")
msg.add_alternative(html, subtype='html')

# Anexa o logo como imagem inline
with open(r'\\arquivosdti.clickip.local\automacao_dados\IA_HUBSOFT\IA_HUBSOFT_PROJECT\Emails\clickip-logo.png', 'rb') as img:
    msg.get_payload()[0].add_related(
        img.read(),
        maintype='image', subtype='png',
        cid=logo_cid
    )

try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login('ia_no-reply@clickip.com.br', 'jkxr fgmq etdi akly')
        smtp.send_message(msg)
    print(f"E-mail enviado com sucesso para {msg['To']}")
except Exception as e:
    print(f"Erro ao enviar e-mail: {e}")
