import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
import csv
import glob
import os
from datetime import datetime, timedelta

# ===========================
# Parâmetros de data e caminhos
# ===========================

# Data de hoje e do início do período (últimos 7 dias)
hoje = datetime.today()
inicio_periodo = hoje - timedelta(days=7)

inicio_str = inicio_periodo.strftime('%Y-%m-%d')
fim_str = hoje.strftime('%Y-%m-%d')

# Pasta onde estão os arquivos CSV diários
pasta_csv = r"\\arquivosdti.clickip.local\automacao_dados\papercut_history_users"
padrao_arquivos = os.path.join(pasta_csv, '*.csv')

# ================================
# Coleta, parsing e filtragem dos dados
# ================================

dados = []
headers = None

for caminho in glob.glob(padrao_arquivos):
    with open(caminho, 'r', encoding='latin-1') as f:
        # Ignora a primeira linha personalizada
        f.readline()
        sample = f.read(2048)
        f.seek(0)
        f.readline()

        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=';,')
        except csv.Error:
            dialect = csv.get_dialect('excel')

        leitor = csv.DictReader(f, dialect=dialect)
        if headers is None:
            headers = leitor.fieldnames

        for linha in leitor:
            time_val = linha.get('Time', '')
            try:
                data_registro = datetime.strptime(time_val, "%Y-%m-%d %H:%M:%S")
                if inicio_periodo <= data_registro <= hoje:
                    dados.append(linha)
                    print(f"Registro encontrado: {linha}")
            except ValueError:
                print(f"Formato de data inválido: {time_val}")

if not dados:
    print(f"Nenhum registro encontrado entre {inicio_str} e {fim_str}. Saindo...")
    exit(0)

# ================================
# Geração do HTML do e-mail
# ================================

logo_cid = make_msgid(domain='clickip.com.br')
th_html = ''.join(f'<th>{col}</th>' for col in headers)
tr_html = ''
for reg in dados:
    cols = ''.join(f'<td>{reg[col]}</td>' for col in headers)
    tr_html += f'<tr>{cols}</tr>\n'

html = f"""
<html>
  <head>
    <style>
      body {{ font-family: Arial, sans-serif; color: #333; line-height: 1.4; }}
      .header {{ display: flex; align-items: center; border-bottom: 2px solid #0072C6; padding-bottom: 10px; margin-bottom: 20px; }}
      .header img {{ height: 50px; margin-right: 15px; }}
      .title {{ color: #0072C6; font-size: 1.5em; margin: 0; }}
      .subtitle {{ font-size: 1em; color: #555; }}
      table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
      th, td {{ padding: 8px 12px; border: 1px solid #ddd; text-align: left; }}
      th {{ background-color: #0072C6; color: white; }}
      tr:nth-child(even) {{ background-color: #f9f9f9; }}
      .footer {{ margin-top: 25px; font-size: 0.85em; color: #777; border-top: 1px solid #eee; padding-top: 10px; }}
    </style>
  </head>
  <body>
    <div class="header">
      <img src="cid:{logo_cid[1:-1]}" alt="ClickIP Logo">
      <div>
        <h1 class="title">Relatório Semanal de Impressões</h1>
        <p class="subtitle">Período: {inicio_str} a {fim_str}</p>
      </div>
    </div>

    <p>Foram lidos <strong>{len(dados)}</strong> registros de impressão entre {inicio_str} e {fim_str}.</p>

    <table>
      <tr>{th_html}</tr>
      {tr_html}
    </table>

    <div style="margin-top: 20px; text-align: center;">
      <a href="http://dti-grafana.clickip.local:3600/login"
         target="_blank"
         style="background-color: #0072C6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold;">
         Para uma visão geral no Grafana-dti, clique aqui!
      </a>
    </div>

    <p class="footer">
      Este é um e-mail automático enviado pelo sistema da ClickIP.  
      Em caso de dúvidas, contate dti@clickip.com.br.
    </p>
  </body>
</html>
"""

# ================================
# Envio do e-mail
# ================================

msg = EmailMessage()
msg['Subject'] = f'📊 Relatório Semanal de Impressões ({inicio_str} a {fim_str})'
msg['From'] = 'ia_no-reply@clickip.com.br'
msg['To'] = 'richard.silva@clickip.com.br'
#msg['Cc'] = 'kethlen.santana@clickip.com.br'
#msg['Bcc'] = 'paulo.farias@clickip.com.br'
msg.add_alternative(html, subtype='html')

with open(r"\\arquivosdti.clickip.local\automacao_dados\IA_HUBSOFT\IA_HUBSOFT_PROJECT\Emails\clickip-logo.png", 'rb') as img:
    msg.get_payload()[0].add_related(img.read(), maintype='image', subtype='png', cid=logo_cid)

try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login('ia_no-reply@clickip.com.br', 'jkxr fgmq etdi akly')
        smtp.send_message(msg)
    print(f"E-mail enviado com sucesso com registros de {inicio_str} a {fim_str}.")
except Exception as e:
    print(f"Erro ao enviar e-mail: {e}")
