import pyautogui
import time
from datetime import datetime, timedelta
import os
import pandas as pd
import re
import numpy as np
from openpyxl import load_workbook


hoje = datetime.now()
first_day_of_this_month = hoje.replace(day=1)
last_day_previous_month = first_day_of_this_month - timedelta(days=1)
#print(last_day_previous_month)
#print(last_day_previous_month.strftime("%d/%m/%Y"))
day=last_day_previous_month.strftime("%d")
month=last_day_previous_month.strftime("%m")
this_month = str(int(month)+1)
year=last_day_previous_month.strftime("%Y")

def multipletabs(num):
    for i in range(num):    
        pyautogui.press('tab')
        time.sleep(0.4)


#melhorias: Ao clicar em relatorios, procurar pela imagem. Dessa forma Não importa se algum menu estiver recolhido ou mudar de posição


#É preciso que o maxbot esteja na tela inicial com todos os menus recolhidos
print("Downloading files...")
nomes = ['Guilherme Soares', 'Luciano Costa', 'Valeria Mendes']
#print(nomes[0])
time.sleep(1)
#go to menu "Relatorios"
pyautogui.moveTo(x=1400, y=460)
time.sleep(0.5)
pyautogui.leftClick()
time.sleep(0.5)
pyautogui.leftClick()
time.sleep(0.5)
pyautogui.moveTo(x=1400, y=540)
time.sleep(0.5)
#Click Comissões de Parceria
pyautogui.leftClick()
time.sleep(0.2)
#Inside menu relatorios go to comissão de parceria
#Move to parceiro
pyautogui.moveTo(x=1765, y=365)
time.sleep(0.4)
#Click Parceiro field
pyautogui.leftClick()
time.sleep(1)
#Write salesperson names
pyautogui.write(nomes[0])
time.sleep(1)
pyautogui.press('Enter')
pyautogui.press('tab')
pyautogui.write('01/01/2018')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.write(f'{day}/{month}/{year}')
time.sleep(0.5)
pyautogui.press('tab')
pyautogui.press('Enter')
time.sleep(0.5)
pyautogui.press('Enter')
time.sleep(3)
multipletabs(2)
pyautogui.press('Enter')
time.sleep(3)
pyautogui.write(f'{month}-{year}-Relatorio-de-comissao_{nomes[0]}_completo')
time.sleep(1)
pyautogui.press('Enter')
#Aqui já clica para baixar o relatório e abre o diretório para escolher onde salvar. 
time.sleep(2)
print("Guilherme`s report saved")
#Close popup with download options
pyautogui.press('escape')
time.sleep(0.3)
#Click Parceiro Field
pyautogui.moveTo(x=1765, y=365)
pyautogui.leftClick()
time.sleep(0.3)
pyautogui.write(nomes[1])
time.sleep(0.5)
pyautogui.press('Enter')
multipletabs(3)
time.sleep(0.5)
pyautogui.press('Enter')
time.sleep(3)
multipletabs(2)
pyautogui.press('Enter')
time.sleep(3)
pyautogui.write(f'{month}-{year}-Relatorio-de-comissao_{nomes[1]}_completo')
time.sleep(1)
pyautogui.press('Enter')
print("Luciano`s report saved")


time.sleep(2)
#Close popup with download options
pyautogui.press('escape')
time.sleep(0.3)
pyautogui.moveTo(x=1765, y=365)
pyautogui.leftClick()
time.sleep(0.3)
pyautogui.write(nomes[2])
time.sleep(0.5)
pyautogui.press('Enter')
multipletabs(3)
time.sleep(0.5)
pyautogui.press('Enter')
time.sleep(3)
multipletabs(2)
pyautogui.press('Enter')
time.sleep(3)
pyautogui.write(f'{month}-{year}-Relatorio-de-comissao_{nomes[2]}_completo')
time.sleep(1)
pyautogui.press('Enter')
time.sleep(1)
pyautogui.press('escape')
time.sleep(0.3)
print("Valeria`s report saved")
print("Data extraction completed successfully")


#
#
#
#
# Changing moving files to the correct directory

hoje = datetime.now()
first_day_of_this_month = hoje.replace(day=1)
last_day_previous_month = first_day_of_this_month - timedelta(days=1)
#print(last_day_previous_month)
#print(last_day_previous_month.strftime("%d/%m/%Y"))
day=last_day_previous_month.strftime("%d")
month=last_day_previous_month.strftime("%m")
this_month = hoje.strftime('%m')
year=last_day_previous_month.strftime("%Y")
this_year=last_day_previous_month.strftime("%Y")

print("Moving files to the correct folder...")
#Moving to the correct folder
origem = '/Users/mateushorta/Desktop'
def destino():
    destino = '/Users/mateushorta/Desktop/03-MAXbot/2 - Boletos/Boletos ' + this_year +'/'+this_month+'-'+this_year+'/Apoio Comissões de Parceria'
    return(destino)

#Verificando a existência do caminho final, se não existir, cria a pasta.

if not os.path.exists(destino()):
    os.makedirs(destino())
    print('Comission folder created')


#verificando quais arquivos tem extensão xlsx e tem Relatorio de comissão no nome.
comissions_report = [
    file for file in os.listdir(origem)
    if file.endswith('.xlsx') and 'Relatorio-de-comissao' in file
    ]
#movendo os arquivos
for file in comissions_report:
    final_destiny = os.path.join(destino(), file)
    os.rename(origem+'/'+file, final_destiny)
    print(f'File {file} moved successfully')


os.chdir(destino())
lista=[]
for file in os.listdir():
    if file.endswith('.xlsx'):
        lista.append(file)

def last_months(x):
    now = pd.Timestamp(datetime.now())
    dates = []
    for i in range (x):
        date = now - pd.DateOffset(months=i+1)
        year_month = date.strftime('%Y/%m')
        dates.append(year_month)
    return dates
last_month = (last_months(1))
last_3_months = last_months(3)

print("Data Analysis in process...")
for file in lista:
    df1 = pd.read_excel(file)
    # Accessing the wanted information
    seller = df1.at[0, 'Unnamed: 1']

    # Substitute the middle part of the CPF with asterisks
    seller = re.sub(r'(\d{3})\.\d{3}\.\d{3}-(\d{2})', r'\1.***.***-\2', seller)
    df1.at[0, 'Unnamed: 1'] = seller
    # Assign the modified seller back to the dataframe
    df1['SELLER'] = seller
    #print(df1.head(15))
    # Creating a new column with the seller's name information
    seller = df1.at[0, 'Unnamed: 1'] #accessing the wanted information
    df1['SELLER'] = seller

    #Creating a new column with the evaluated period
    period = df1.at[3, 'Unnamed: 1']
    df1['PERIOD'] = period

    #utilizing the correct line as column names
    df1.columns = df1.iloc[5]
    df1.drop(df1.index[:7], inplace = True)

    #Deleting a useless column 'VALOR COMISSÃO'
    df1.drop('VALOR COMISSÃO', axis=1, inplace = True)

    #renaming the second column as USERNAME instead of NaN
    df1 = df1.rename(columns={np.nan: 'USERNAME'})
    #renaming the index column


    #reseting the index column
    df1.reset_index(drop = True, inplace = True)
    #Removing index name
    df1 = df1.rename_axis(None, axis=1)
    #Converting the data type of the 'DATA PAGTO' column to a date.

    #Joining the two following steps into one.
    #df['DATA PAGTO'] = pd.to_datetime(df['DATA PAGTO'], format='%d/%m/%Y')
    #df['DATA PAGTO'] = df['DATA PAGTO'].dt.strftime('%Y/%m')

    df1['DATA PAGTO'] = pd.to_datetime(df1['DATA PAGTO'], format='%d/%m/%Y').dt.strftime('%Y/%m')
    table1 = pd.pivot_table(df1, values = 'VALOR PAGO', index = 'USERNAME', columns = 'DATA PAGTO', aggfunc= 'sum', fill_value = 0)
    table1['NUM_PAYMENTS'] = (table1>0).sum(axis=1)
    table1['NUM_PAYMENTS_LAST3_MON'] = table1.loc[:,last_3_months].gt(0).sum(axis=1)
    table1['NUM_PAYMENTS_LAST3_MON'] = table1.loc[:,last_3_months].gt(0).sum(axis=1)
    dict_map = {True: 'Elegible', False:'Ineligible'}
    table1['COMISSION'] = ((table1['NUM_PAYMENTS'] <=3 ) & (table1['NUM_PAYMENTS'] == table1['NUM_PAYMENTS_LAST3_MON']) & (table1[last_month[0]] > 0)).map(dict_map)
    table1 = table1.sort_values(by='COMISSION')
    with pd.ExcelWriter(file, engine = 'openpyxl', mode = 'a', if_sheet_exists = 'replace') as writer:
        table1.to_excel(writer, sheet_name='DadosComissoes')
        print(f'ETL process complete for {file}')