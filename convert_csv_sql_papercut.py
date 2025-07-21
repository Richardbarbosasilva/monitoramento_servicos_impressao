import os
import mariadb
import pyautogui

pasta_csv = r"\\arquivosdti.clickip.local\\automacao_dados\\papercut_history_users"

########################## Conexão com o banco de dados MariaDB ##########################

try:

    conn = mariadb.connect(

        host="172.25.200.120",
        user="root",
        port=3333,
        password="mariadb_2NAz5k",
        database="papercut_impressoras_logs_db"

    )

    cursor = conn.cursor()

    print("Conexão com o banco de dados estabelecida.")

except mariadb.Error as e:

    print(f"Erro ao conectar no banco de dados: {e}")
    exit(1)

print("Conexão bem-sucedida com o banco mariadb.\n")
print("Iniciando a importação dos arquivos CSV...\n")

######################### Criação da tabela logs_impressoras caso não exista #########################


cursor.execute("""
CREATE TABLE IF NOT EXISTS logs_impressoras (
    id INT AUTO_INCREMENT PRIMARY KEY,
    data DATETIME,
    usuario_logado VARCHAR(255),
    paginas INT,
    copias INT,
    impressora VARCHAR(255),
    nome_documento TEXT,
    cliente VARCHAR(255),
    papel_tamanho VARCHAR(50),
    linguagem VARCHAR(50),
    altura VARCHAR(50),
    largura VARCHAR(50),
    duplex VARCHAR(20),
    escala_de_cinza VARCHAR(20),
    tamanho VARCHAR(50),
    host VARCHAR(255),
    usuario_arquivo VARCHAR(255),
    UNIQUE KEY uidx_log (
        data, usuario_logado, nome_documento, impressora, host
    )
)
""")

def clean(val):
    return None if val == "" else val

#################### iteração pelos arquivos CSV na pasta especificada ####################

for arquivo in os.listdir(pasta_csv):
    if not arquivo.lower().endswith(".csv"):
        continue

    print(f"Lendo arquivo: {arquivo}")
    caminho = os.path.join(pasta_csv, arquivo)

    # Extrai host e usuário do nome do arquivo

    base = arquivo[:-4]

    try:
        _, _, rest = base.split("_", 2)
        idx = rest.rfind("-")
        host = rest[:idx]
        usuario_arquivo = rest[idx+1:]

    except Exception:

        print(f"  → nome '{arquivo}' não segue o padrão, pulando.\n")
        continue

# Lê o arquivo CSV e insere os dados na tabela logs_impressoras

    print(f"  → host: {host}, usuário do arquivo: {usuario_arquivo}\n")

    inseridas = 0
    with open(caminho, "r", encoding="latin1", errors="ignore") as f:
        # pula a primeira linha (banner) e a segunda (header)
        next(f)
        next(f)
        for linha in f:
            linha = linha.strip()
            if not linha:
                continue
            # split nos primeiros 13 separadores => 14 campos
            partes = linha.split(",", 13)
            if len(partes) < 14:
                print(f"  → linha com poucos campos, pulando: {linha}")
                continue
               

####################### Execução da Query de inserção #######################


            time, user, pages, copies, printer, docname, client, papersize, language, height, width, duplex, grayscale, size = partes
            try:
                cursor.execute("""
                               
                INSERT INTO logs_impressoras (
                    data, usuario_logado, paginas, copias, impressora,
                    nome_documento, cliente, papel_tamanho, linguagem,
                    altura, largura, duplex, escala_de_cinza, tamanho,
                    host, usuario_arquivo
                               
                ) VALUES (
                    %s, %s, %s, %s, %s,
                    %s, %s, %s, %s,
                    %s, %s, %s, %s, %s,
                    %s, %s
                )
                """, 
                (
                    clean(time),
                    clean(user),
                    clean(pages),
                    clean(copies),
                    clean(printer),
                    clean(docname),
                    clean(client),
                    clean(papersize),
                    clean(language),
                    clean(height),
                    clean(width),
                    clean(duplex),
                    clean(grayscale),
                    clean(size),
                    host,
                    usuario_arquivo
                ))

                inseridas += 1

            except mariadb.Error as e:

                print(f"! falha ao inserir linha: {e}")

    print(f"  → {inseridas} linhas inseridas de {arquivo}.\n")
    print("Finalizando a leitura do arquivo CSV\n")

# Commit das alterações no banco de dados e fechamento da conexão

conn.commit()
cursor.close()
conn.close()

print("Importação finalizada com sucesso.")


#SELECT 
#  data,
#  usuario_logado,
#  nome_documento,
##  paginas,
# copias,
#  impressora,
#  host
#FROM logs_impressoras
#ORDER BY data DESC
#LIMIT 100
