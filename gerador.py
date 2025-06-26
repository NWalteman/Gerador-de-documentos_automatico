# Importa a biblioteca pandas, utilizada para ler e manipular dados em tabelas (como arquivos Excel)
import pandas as pd
# Importa a classe DocxTemplate da biblioteca docxtpl, usada para preencher modelos (.docx) com dados
from docxtpl import DocxTemplate
# Importa o módulo os, que permite interagir com o sistema de arquivos (como criar pastas)
import os

CAMINHO_TEMPLATE = "colocar/aqui/o/caminho/do/template.docx"
CAMINHO_PLANILHA = "colocar/aqui/o/caminho/da/planilha.xlsx"
PASTA_SAIDA = "documentos_gerados"

df = pd.read_excel(CAMINHO_PLANILHA)
os.makedirs(PASTA_SAIDA, exist_ok=True)

for _, linha in df.iterrows():
    doc = DocxTemplate(CAMINHO_TEMPLATE)
    contexto = {
        "Nome": linha["Nome"],
        "Cargo": linha["Cargo"]
    }
    nome_arquivo = f"{PASTA_SAIDA}/ASO_{linha['Nome'].replace(' ', '_')}.docx"
    doc.render(contexto)
    doc.save(nome_arquivo)
    print(f"✔ Documento gerado: {nome_arquivo}")
