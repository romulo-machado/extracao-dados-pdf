from pdfminer.high_level import extract_text
from pdfminer.layout import LAParams
import pandas as pd
from IPython.display import display
import re
import tabula

lista_tabela = tabula.read_pdf("FICHA DE REGISTRO.pdf", pages="all",guess=False)
tabela = lista_tabela[1]
tabela.columns = tabela.iloc[0]
salario = tabela['Valor Salário'][1]

# Substitua 'seu_arquivo.pdf' pelo caminho do seu arquivo PDF
pdf_path = 'FICHA DE REGISTRO.pdf'

# Criar um objeto LAParams com vários parâmetros
params_layout = LAParams(
    line_margin=1.5,
    char_margin=3.5,
    word_margin=2.5,
    boxes_flow=-1.0,
)

# Extrair texto do PDF usando o objeto LAParams
texto_pdf = extract_text(pdf_path, laparams=params_layout)

# Dividir o texto em linhas
linhas = texto_pdf.split('\n')

# Criar um dicionário para armazenar os dados
dados = {}

# Iterar pelas linhas e adicionar ao dicionário
for linha in linhas:
    # Verificar se a linha contém o caractere ':' para evitar o erro "not enough values to unpack"
    if ':' in linha:
        chave, valor = linha.split(":", 1)
        chave = chave.strip()
        valor = valor.strip()

        # Caso especial para tratar "Emissão RG"
        if chave == 'Estado' and 'Emissão RG' in valor:
            chave_emissao_rg, valor_emissao_rg = valor.split('Emissão RG', 1)
            dados['Estado'] = chave_emissao_rg.strip()
            dados['Emissão RG'] = valor_emissao_rg.replace(':', '').strip()
        else:
            dados[chave] = valor

# Criar um DataFrame a partir do dicionário
df_dados = pd.DataFrame([dados])

# Remover os números da coluna '1.412,00  Modo Pgto' e renomear a coluna
df_dados['1.412,00  Modo Pgto'] = df_dados['1.412,00  Modo Pgto'].apply(lambda x: re.sub(r'\d', '', x))
df_dados.rename(columns={'1.412,00  Modo Pgto': 'Modo Pgto'}, inplace=True)

# # Selecionar colunas desejadas
# df_selecionado = df_dados[['Nome', 'Remuneração', '1.412,00  Modo Pgto']]

# Selecionar apenas as colunas desejadas
colunas_desejadas = [
    'Código','Cargo', 'Nome', 'Pai', 'Mãe', 'Nascimento', 'Sexo', 'Est. Civil', 
    'Raça/Cor', 'Naturalidade', 'Nacionalidade', 'Endereço', 'Bairro', 'CEP', 
    'Município', 'CPF', 'RG', 'Órgão', 'Estado', 'Emissão RG',
    'PIS', 'Admissão',  'Remuneração', 'Organograma', 'Escala', 'CNPJ/CEI'
]

# Criar um DataFrame a partir do dicionário e selecionar as colunas desejadas
df = pd.DataFrame([dados])[colunas_desejadas]

df.at[0, 'Remuneração'] = salario

nome = df['Nome'].to_string(index=False)

df.to_excel(f'{nome}.xlsx', index=False)

# Exibir o DataFrame
display(df)