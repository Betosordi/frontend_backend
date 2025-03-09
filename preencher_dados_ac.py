import pandas as pd
from bs4 import BeautifulSoup
import pyautogui
import os
import urllib.request

caminho_da_pb = f"//saowsr0011/Schindler/IQL/BASELSS/Projeto_JRS/desenvolvedor/minifabrica_ac.html"


lendo_caminho_pb = f"//saowsr0011/Schindler/IQL/BASELSS/Projeto_JRS/dados_producao_ac/Macro AC.xlsm"

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
tag_1 = soup.find(id='ordens_produzidas_ac03')
if tag_1:
    tag_1['value'] = primeiro_valor
#------------------------------------------------------------------

# Segunda tag
# Pegar o segunda valor do DataFrame
segundo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[1])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_2 = soup.find(id='ordens_produzidas_ac18')
if tag_2:
    tag_2['value'] = segundo_valor
#------------------------------------------------------------------

# Terceira tag
# Pegar o Terceira valor do DataFrame
terceiro_valor = int(planilha_filtrada['Quantidade produzida'].iloc[2])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_3 = soup.find(id='ordens_produzidas_ac14')
if tag_3:
    tag_3['value'] = terceiro_valor
#------------------------------------------------------------------

# Quarta tag
# Pegar o Quarto valor do DataFrame
quarto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[3])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_4 = soup.find(id='ordens_produzidas_ac09')
if tag_4:
    tag_4['value'] = quarto_valor
#------------------------------------------------------------------

# Quinta tag
# Pegar o Quinto valor do DataFrame
quinto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[4])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_5 = soup.find(id='ordens_produzidas_ac24')
if tag_5:
    tag_5['value'] = quinto_valor
#------------------------------------------------------------------

# Sexta tag
# Pegar o Sexto valor do DataFrame
sexto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[5])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_6 = soup.find(id='ordens_produzidas_ac08')
if tag_6:
    tag_6['value'] = sexto_valor
#------------------------------------------------------------------

# Sétima tag
# Pegar o Sétimo valor do DataFrame
setimo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[6])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_7 = soup.find(id='ordens_produzidas_ac13')
if tag_7:
    tag_7['value'] = setimo_valor
#------------------------------------------------------------------

# Oitava tag
# Pegar o Oitavo valor do DataFrame
oitavo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[7])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_8 = soup.find(id='ordens_produzidas_ac11')
if tag_8:
    tag_8['value'] = oitavo_valor
#------------------------------------------------------------------

# Nona tag
# Pegar o Nona valor do DataFrame
nono_valor = int(planilha_filtrada['Quantidade produzida'].iloc[8])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_9 = soup.find(id='ordens_produzidas_ac17')
if tag_9:
    tag_9['value'] = nono_valor
#------------------------------------------------------------------

# Décima tag
# Pegar o Décimo valor do DataFrame
decimo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[9])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_10 = soup.find(id='ordens_produzidas_ac19')
if tag_10:
    tag_10['value'] = decimo_valor
#------------------------------------------------------------------

# 11° tag
# Pegar o 11° valor do DataFrame
decimo_primeiro_valor = int(planilha_filtrada['Quantidade produzida'].iloc[10])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_11 = soup.find(id='ordens_produzidas_ac21')
if tag_11:
    tag_11['value'] = decimo_primeiro_valor
#------------------------------------------------------------------

# 12° tag
# Pegar o 12° valor do DataFrame
decimo_segundo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[11])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_12 = soup.find(id='ordens_produzidas_ac12')
if tag_12:
    tag_12['value'] = decimo_segundo_valor
#------------------------------------------------------------------ 

# 13° tag
# Pegar o 13° valor do DataFrame
decimo_terceiro_valor = int(planilha_filtrada['Quantidade produzida'].iloc[12])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_13 = soup.find(id='ordens_produzidas_ac22')
if tag_13:
    tag_13['value'] = decimo_terceiro_valor
#------------------------------------------------------------------ 

# 14° tag
# Pegar o 14° valor do DataFrame
decimo_quarto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[13])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_14 = soup.find(id='ordens_produzidas_ac25')
if tag_14:
    tag_14['value'] = decimo_quarto_valor
#------------------------------------------------------------------ 

# 15° tag
# Pegar o 15° valor do DataFrame
decimo_quinto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[14])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_15 = soup.find(id='ordens_produzidas_ac26')
if tag_15:
    tag_15['value'] = decimo_quinto_valor
#------------------------------------------------------------------ 

# 16° tag
# Pegar o 16° valor do DataFrame
decimo_sexto_valor = int(planilha_filtrada['Quantidade produzida'].iloc[15])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_16 = soup.find(id='ordens_produzidas_ac10')
if tag_16:
    tag_16['value'] = decimo_sexto_valor
#------------------------------------------------------------------ 

# 17° tag
# Pegar o 17° valor do DataFrame
decimo_setimo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[16])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_17 = soup.find(id='ordens_produzidas_ac27')
if tag_17:
    tag_17['value'] = decimo_setimo_valor
#------------------------------------------------------------------ 

# 18° tag
# Pegar o 18° valor do DataFrame
decimo_oitavo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[17])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_18 = soup.find(id='ordens_produzidas_ac01')
if tag_18:
    tag_18['value'] = decimo_oitavo_valor
#------------------------------------------------------------------ 

# 19° tag
# Pegar o 19° valor do DataFrame
decimo_nono_valor = int(planilha_filtrada['Quantidade produzida'].iloc[18])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_19 = soup.find(id='ordens_produzidas_ac02')
if tag_19:
    tag_19['value'] = decimo_nono_valor
#------------------------------------------------------------------ 

# 20° tag
# Pegar o 20° valor do DataFrame
vigesimo_valor = int(planilha_filtrada['Quantidade produzida'].iloc[19])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_20 = soup.find(id='ordens_produzidas_ac04')
if tag_20:
    tag_20['value'] = vigesimo_valor
#------------------------------------------------------------------ 

# 21° tag
# Pegar o 21° valor do DataFrame
vigesimo1_valor = int(planilha_filtrada['Quantidade produzida'].iloc[20])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_21 = soup.find(id='ordens_produzidas_ac05')
if tag_21:
    tag_21['value'] = vigesimo1_valor
#------------------------------------------------------------------ 

# 22° tag
# Pegar o 22° valor do DataFrame
vigesimo2_valor = int(planilha_filtrada['Quantidade produzida'].iloc[21])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_22 = soup.find(id='ordens_produzidas_ac06')
if tag_22:
    tag_22['value'] = vigesimo2_valor
#------------------------------------------------------------------ 

# 23° tag
# Pegar o 23° valor do DataFrame
vigesimo3_valor = int(planilha_filtrada['Quantidade produzida'].iloc[22])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_23 = soup.find(id='ordens_produzidas_ac15')
if tag_23:
    tag_23['value'] = vigesimo3_valor
#------------------------------------------------------------------ 

# 24° tag
# Pegar o 24° valor do DataFrame
vigesimo4_valor = int(planilha_filtrada['Quantidade produzida'].iloc[23])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_24 = soup.find(id='ordens_produzidas_ac16')
if tag_24:
    tag_24['value'] = vigesimo4_valor
#------------------------------------------------------------------ 

# 25° tag
# Pegar o 25° valor do DataFrame
vigesimo5_valor = int(planilha_filtrada['Quantidade produzida'].iloc[24])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_25 = soup.find(id='ordens_produzidas_acphase')
if tag_25:
    tag_25['value'] = vigesimo5_valor
#------------------------------------------------------------------ 

# 26° tag
# Pegar o 26° valor do DataFrame
vigesimo6_valor = int(planilha_filtrada['Quantidade produzida'].iloc[25])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_26 = soup.find(id='ordens_produzidas_diversos')
if tag_26:
    tag_26['value'] = vigesimo6_valor
#------------------------------------------------------------------ 

# 27° tag
# Pegar o 27° valor do DataFrame
vigesimo7_valor = int(planilha_filtrada['Quantidade produzida'].iloc[26])
# Encontrar a tag pelo ID e adicionar o valor desejado no input
tag_27 = soup.find(id='ordens_produzidas_ac23')
if tag_27:
    tag_27['value'] = vigesimo7_valor
#------------------------------------------------------------------ 

#--------------------------------------------  
# Salvar o HTML modificado de volta ao arquivo
with open(caminho_da_pb, 'w') as file:
    file.write(str(soup))

pyautogui.alert("Abrir página para preencher formulário")


