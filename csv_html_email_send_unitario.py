import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
import csv
from datetime import date


pasta_csv = r"UNC_PATH"




# ===========================
# Parâmetros de data e caminhos
# ===========================

hoje = date.today()
# Ajuste o caminho para onde seu CSV está armazenado
caminho_csv = r"caminho_do_log_csv"

# ================================
# Leitura do CSV e captura dos dados
# ================================

dados = []
with open(caminho_csv, 'r', encoding='latin-1') as f:
    # Ignora a primeira linha de cabeçalho customizado
    f.readline()
    # Detecta delimitador (ex.: ';' ou ',') com sample a partir da segunda linha
    sample = f.read(2048)
    f.seek( f.tell() )  # já está no início da leitura de dados
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=';,')
    except csv.Error:
        dialect = csv.get_dialect('excel')
    f.seek( f.tell() - len(sample) )  # volta para o início dos dados reais
    leitor = csv.DictReader(f, dialect=dialect)
    headers = leitor.fieldnames
    for linha in leitor:
        dados.append(linha)

# ================================
# Geração do HTML do e-mail
# ================================

logo_cid = make_msgid(domain='domain')
th_html = ''.join(f'<th>{col}</th>' for col in headers)
tr_html = ''
for registro in dados:
    cols = ''.join(f'<td>{registro[col]}</td>' for col in headers)
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
      <img src="cid:{logo_cid[1:-1]}" alt="Logo">
      <div>
        <h1 class="title">Relatório Papercut - Dados CSV</h1>
        <p class="subtitle">Data de geração: {hoje}</p>
      </div>
    </div>

    <p>Foram lidas <strong>{len(dados)}</strong> linhas do arquivo CSV (exceto cabeçalho inicial).</p>

    <table>
      <tr>{th_html}</tr>
      {tr_html}
    </table>

    #<div style="margin-top: 20px; text-align: center;">
    #  <a href="http://site"
    #     target="_blank" 
    #     style="background-color: #0072C6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold;">
         
      </a>
    </div>

    <p class="footer">
      Este é um e-mail automático enviado pelo sistema de análise.  
      Em caso de dúvidas, contate email@email.com.
    </p>
  </body>
</html>
"""

# ================================
# Envio do e-mail com SMTP e logo inline
# ================================

msg = EmailMessage()
msg['Subject'] = '📊 Relatório Papercut CSV'
msg['From'] = 'remetente'
msg['To'] = 'destinatário'
#msg['Cc'] = 'destinatário'
msg.add_alternative(html, subtype='html')

with open(r'png_caminho', 'rb') as img:
    msg.get_payload()[0].add_related(
        img.read(),
        maintype='image', subtype='png',
        cid=logo_cid
    )

try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login('remetente@email.com', 'senha_segura')
        smtp.send_message(msg)
    print(f"E-mail enviado com sucesso para {msg['To']}")
except Exception as e:
    print(f"Erro ao enviar e-mail: {e}")
