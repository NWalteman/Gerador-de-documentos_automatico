# Importa a biblioteca pandas, utilizada para ler e manipular dados em tabelas (como arquivos Excel)
import pandas as pd

# Importa a classe DocxTemplate da biblioteca docxtpl, usada para preencher modelos (.docx) com dados
from docxtpl import DocxTemplate

# Importa o módulo os, que permite interagir com o sistema de arquivos (como criar pastas)
import os

# Caminho do arquivo modelo (template .docx) que será preenchido com os dados da planilha
CAMINHO_TEMPLATE = "colocar/aqui/o/caminho/do/template.docx"

# Caminho da planilha Excel (.xlsx) que contém os dados a serem inseridos nos documentos
CAMINHO_PLANILHA = "colocar/aqui/o/caminho/da/planilha.xlsx"

# Nome da pasta onde os documentos gerados serão salvos
PASTA_SAIDA = "documentos_gerados"

# Lê a planilha do caminho informado e armazena os dados em um DataFrame (estrutura de tabela do pandas)
df = pd.read_excel(CAMINHO_PLANILHA)

# Cria a pasta de saída (caso ainda não exista) para salvar os documentos gerados
os.makedirs(PASTA_SAIDA, exist_ok=True)

for _, linha in df.iterrows():
    # Carrega o modelo do documento .docx para cada pessoa
    doc = DocxTemplate(CAMINHO_TEMPLATE)

    # Cria um dicionário com os dados que vão substituir as variáveis no template
    contexto = {
        "Nome": linha["Nome"],   # Substitui a variável {{ Nome }} no template pelo valor da coluna "Nome"
        "Cargo": linha["Cargo"]  # Substitui a variável {{ Cargo }} no template pelo valor da coluna "Cargo"
    }

    # Gera um nome de arquivo com base no nome da pessoa, substituindo espaços por "_"
    nome_arquivo = f"{PASTA_SAIDA}/ASO_{linha['Nome'].replace(' ', '_')}.docx"

    # Preenche o template com os dados do dicionário 'contexto'
    doc.render(contexto)

    # Salva o documento gerado no caminho especificado
    doc.save(nome_arquivo)

    # Mostra no terminal que o documento foi criado com sucesso
    print(f"✔ Documento gerado: {nome_arquivo}")
