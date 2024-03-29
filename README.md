<h1>PROJETO PARA VISÃO DE MERCADO DO AGRONEGÓCIO BRASILEIRO</h1>
<li>Demonstração do potencial de mercado do <strong>Agronegócio Brasileiro</strong>, todos os dados utilizados são de dominio público - <a href="https://www.ibge.gov.br/"> Clique aqui; </a></li>
<li>Demonstração de todas as etapas para construção do produto final, painel gerencial em Looker Studio.</li><br>

Linguagem e pacotes necessários (Python 3.12.2 64-bit):

```
pip install pandas
```
```
pip install openpyxl
```
<br>Código para extrair todas informações dos 54 arquivos .xlsx
```
import os
import pandas as pd

# Diretório onde estão os arquivos Excel
diretorio = '.../Downloads/Todas lavouras/'

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

<br>Layout original de todos arquivos .xlsx
<br>Objetivo inicial é extrair as seguites informações:
<li>Extrair todos os municípios.</li>
<li>Extrair toda área plantada.</li>
<li>Extrair todas as culturas.</li>
<li>Extrair todos os estados.</li><br>

![image](https://github.com/BYTE-JoseLucas/Agronegocio/assets/99023240/3285e987-ec93-451e-a55f-e46f2a596ade)

<br>Layout após código rodar sobre todos os arquivos .xlsx gerou o arquivo 'dados_completos.xlsx' com um total de 462.179 linhas.
<br>As seguintes colunas contem as seguintes informações:
<li>Coluna1 - Contêm todos os municípios.</li>
<li>Coluna2 - Contêm toda área plantada.</li>
<li>Nome da Aba - Contêm todas as culturas.</li>
<li>Nome do Arquivo - Contêm todos os estados.</li><br>

![image](https://github.com/BYTE-JoseLucas/Agronegocio/assets/99023240/c0652e7c-fd72-4897-851c-5f7a20ebc8ac)

