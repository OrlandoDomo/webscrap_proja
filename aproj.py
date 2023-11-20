import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

def inei_data(year, month):

    if year!=2022 or year!=2023:
        return "Ingresar un anio 2022 o 2023"
    if month not in range(12):
        return "Ingresar un indice de mes valido de 0 a 11"

    url='https://www.inei.gob.pe/principales_indicadores/'

    browser= webdriver.Firefox()
    browser.get(url)

    browser.find_element(By.XPATH, "//ul[@class='thumbs barraindicadores']/li[2]").click() # Seleccionamos la categoria de datos correspondientes

    data= WebDriverWait(browser, 10).until(lambda x: x.find_element(By.XPATH, "//li[@class='li_1']/div[2]/div[2]/div[4]/div[2]/a")) # Seleccionamos el boton para descargar

    data.click()
    browser.close() # Cerramos el browser luego de descargar

    download_path='/home/a8neru/Downloads/'
    filename='Serie_del_IPCLM_-_Base_Dic2021_10-2023.xlsx' # Seteamos el path para leer el archivo con pandas

    df=pd.read_excel(download_path+filename, sheet_name="B2009")

    ## Este primer 'for' es para eliminar todos los NaN y reemplazarlos con el respectivo anio que les corresponde

    for i in range(len(df.iloc[2])):
        value=df.iloc[2].iloc[i]
        if pd.isna(value):
            value=temp
        else:
            temp=value
        df.iloc[2].iloc[i]= temp

    ## Este segundo 'for' es para eliminar todos los NaN y reemplazarlos con el respectivo anio que les corresponde

    df2=pd.read_excel(download_path+filename, sheet_name="BDic2021")
    for i in range(len(df2.iloc[2])):
        value=df2.iloc[2].iloc[i]
        if pd.isna(value):
            value=temp
        else:
            temp=value
        df2.iloc[2].iloc[i]= temp

    ## Ahora iteramos para bucsar el valor correspondiente al anio y mes que se ingresa como input en la funcion

    for i in range(len(df2.iloc[2])):
        value=df2.iloc[2].iloc[i]
        if value==year:
            latest_value2= float(df2.iloc[4].iloc[i+month])
            # print('Valor en el year:{} y month {} es {}'.format(year, df2.iloc[3].iloc[i+month], float(df2.iloc[4].iloc[i+month])))
            break;

    latest_month2= df2.iloc[3].iloc[i+month] # Guardamos el valor obtenido

    ## Creamos un dataframe con los valores obtenidos

    excel_df=pd.DataFrame({
        'Nivel_de_desagregacion': 'Indice General',
        str(latest_year)+'_'+latest_month: [latest_value],
        str(year)+'_'+latest_month2: [latest_value2]
    })

    print(excel_df)
