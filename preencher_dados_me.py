import pandas as pd
from bs4 import BeautifulSoup
import pyautogui

caminho_da_pb = f"//saowsr0011/Schindler/IQL/BASELSS/Projeto_JRS/desenvolvedor/minifabrica_me.html"


lendo_caminho_pb = f"//saowsr0011/Schindler/IQL/BASELSS/Projeto_JRS/dados_producao_me/Macro ME.xlsm"

planilha_pb = pd.read_excel(lendo_caminho_pb, sheet_name="Qtd. Células de produção",
                            usecols=["Id da tag", "Quantidade produzida"])

# Filtra as linhas que não contêm NaN na coluna "Quantidade produzida"
planilha_filtrada = planilha_pb.dropna(subset=["Id da tag"])

# Conectando ao HTML ##################################################################
# Carregar o conteúdo do arquivo HTML, lendo e conectando á pagina
with open(caminho_da_pb, 'r') as file:
    html_content = file.read()
# Parsear o HTML
soup = BeautifulSoup(html_content, 'html.parser')   

# Primeira tag
# Pegar o primeiro valor do DataFrame
primeiro_valor = int(planilha_filtrada['Quantidade produzida'].iloc[0])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_1 = soup.find(id='ordens_produzidas_me13')
if tag_1:
    tag_1['value'] = primeiro_valor
#------------------------------------------------------------------

# Segunda tag
# Pegar o segunda valor do DataFrame
segundo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[1])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_2 = soup.find(id='ordens_produzidas_me19')
if tag_2:
    tag_2['value'] = segundo_valor
#------------------------------------------------------------------

# Terceira tag
# Pegar o Terceira valor do DataFrame
terceiro_valor = int(planilha_filtrada['Quantidade produzida'].iloc[2])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_3 = soup.find(id='ordens_produzidas_me18')
if tag_3:
    tag_3['value'] = terceiro_valor
#------------------------------------------------------------------

# Quarta tag
# Pegar o Quarto valor do DataFrame
quarto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[3])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_4 = soup.find(id='ordens_produzidas_me09')
if tag_4:
    tag_4['value'] = quarto_valor
#------------------------------------------------------------------

# Quinta tag
# Pegar o Quinto valor do DataFrame
quinto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[4])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_5 = soup.find(id='ordens_produzidas_me16')
if tag_5:
    tag_5['value'] = quinto_valor
#------------------------------------------------------------------

# Sexta tag
# Pegar o Sexto valor do DataFrame
sexto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[5])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_6 = soup.find(id='ordens_produzidas_me08')
if tag_6:
    tag_6['value'] = sexto_valor
#------------------------------------------------------------------

# Sétima tag
# Pegar o Sétimo valor do DataFrame
setimo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[6])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_7 = soup.find(id='ordens_produzidas_me11')
if tag_7:
    tag_7['value'] = setimo_valor
#------------------------------------------------------------------

# Oitava tag
# Pegar o Oitavo valor do DataFrame
oitavo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[7])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_8 = soup.find(id='ordens_produzidas_me17')
if tag_8:
    tag_8['value'] = oitavo_valor
#------------------------------------------------------------------



#--------------------------------------------  
# Salvar o HTML modificado de volta ao arquivo
with open(caminho_da_pb, 'w') as file:
    file.write(str(soup))

pyautogui.alert("Abrir página para preencher formulário")


