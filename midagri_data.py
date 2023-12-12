import time
import zipfile
import os

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager


def download_midagri_data():

    '''
    Downloads the most recent excel file at the top of the page. Needs to change url according to the year.
    We need Chrome or edge drivers for this, since the page doesnt support Firefox (geckodriver)
    '''

    url='https://siea.midagri.gob.pe/portal/publicacion/boletines-mensuales/19-vbp-agropecuaria/188-vbp-23'
    
    local_dir= os.path.dirname(os.path.realpath(__file__))
    prefs= {"download.default_directory": local_dir}

    # We setup the edge driver to use with selenium.

    edge_options = Options()
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument('headless') # Need headless option cuz linux doesnt support gui
    edge_options.add_experimental_option('prefs', prefs)

    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options= edge_options)
    driver.get(url)

    data= driver.find_element(By.XPATH,"//a[@class='btn btn-success']") # Downloads the element at the top of the page since its the latest.
    data.click()

    tiempo_de_espera_descarga = 5   # Set a time to wait for it to download
    time.sleep(tiempo_de_espera_descarga)
    link= data.get_attribute('href')
    driver.quit()

    file_name, member_to_extract= get_file_name(link) 
    file_path= local_dir + file_name

    # Now we extract the excel and delete the zip file

    with zipfile.ZipFile(file_path, 'r') as zip_ref: 
        zip_ref.extract(member=member_to_extract, path=local_dir)

    bash_command = "rm {}".format(file_path)
    os.system(bash_command)

    return file_path

def get_file_name(link: str):
    '''
    Gets the file names for the excel and the downloaded file.
    '''
    irre, rel= link.split('vbp-')
    year, month= rel.split('-')
    file_name= 'datos_vbp_{}_{}.zip'.format(month, year)
    pdf_name= 'VBP_{}_{}.pdf'.format(month, year)
    excel= 'VBP_{}{}.xlsx'.format(month, year[2:4])
    #print('{}\n{}\n{}\n'.format(file_name, pdf_name, excel))
    return file_name, excel