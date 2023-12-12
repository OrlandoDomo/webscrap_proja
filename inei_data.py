import zipfile
import os

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager


def download_inei_data_pbi():
    '''
    Downloads the excel file for "Produccion nacional", which contains updated data for past months.
    '''

    url='https://www.inei.gob.pe/principales_indicadores/' 
    local_dir= os.path.dirname(os.path.realpath(__file__))
    file_name='CalculoPBI_124.zip' # The file name should remain the same always so its just cosntant.

    prefs= {"download.default_directory": local_dir}

    # We setup the edge driver to use with selenium.

    edge_options = Options()
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument('headless') ## Need headless option cuz linux doesnt support gui
    edge_options.add_experimental_option('prefs', prefs)

    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options= edge_options)
    driver.get(url)
    
    # We need to wait for the element to load

    data= WebDriverWait(driver, 10).until(lambda x: x.find_element(By.XPATH, "//li[@class='li_5']/div[2]/div[2]/div[4]/div[2]/a"))
    data.click()

    # Waiting for the file to download

    tiempo_de_espera_descarga = 5
    time.sleep(tiempo_de_espera_descarga)
    driver.quit()

    file_path= local_dir+file_name
    member_to_extract= '1 Cálculo Agropecuario B 2007 r.xlsx' # This file name seems to not change over time so remains constant

    # Unzip and delete

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extract(member=member_to_extract, path=local_dir)

    bash_command = "rm {}".format(file_path)
    os.system(bash_command)

    return file_path