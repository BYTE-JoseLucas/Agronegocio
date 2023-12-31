<h1>PROJETO PARA VISÃO DE MERCADO DO AGRONEGÓCIO BRASILEIRO</h1>
<li>Demonstração do potencial de mercado do <strong>Agronegócio Brasileiro</strong>, todos os dados utilizados são de dominio público - <a href="https://www.ibge.gov.br/"> Clique aqui; </a></li>
<li>Demonstração de todas as etapas para construção do produto final, painel gerencial em Looker Studio.</li><br>

```
import os
import pandas as pd

# Diretório onde estão os arquivos Excel
diretorio = 'C:/Users/jose.lucas/Desktop/Agronegocio/Todas lavouras/'

# Inicia um DataFrame vazio
dados = pd.DataFrame(columns=['Coluna1', 'Coluna2', 'Nome da Aba', 'Nome do Arquivo'])

# Percorra todos os arquivos no diretório
for arquivo in os.listdir(diretorio):
    if arquivo.endswith('.xlsx'):
        # Abra o arquivo Excel
        xls = pd.ExcelFile(os.path.join(diretorio, arquivo))
        
        # Percorra todas as abas no arquivo
        for aba in xls.sheet_names:
            # Leia o conteúdo da aba para um DataFrame
            df = pd.read_excel(xls, aba, header=None)
            
            # Extraia as informações e as adicione ao DataFrame final
            dados = pd.concat([dados, pd.DataFrame({
                'Coluna1': df.iloc[5:, 0].values,
                'Coluna2': df.iloc[5:, 1].values,
                'Nome da Aba': aba,
                'Nome do Arquivo': arquivo
            })], ignore_index=True)

# Salve o DataFrame em um arquivo Excel
dados.to_excel('dados_completos.xlsx', index=False)
```
