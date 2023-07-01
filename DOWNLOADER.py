from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from time import sleep
import glob
import os.path
import shutil
import os
from bs4 import BeautifulSoup
import re


class Downloader_files:
    def __init__(self):
        self.URL = 'https://www.coes.org.pe/Portal/mercadomayorista/liquidaciones' 
        self.option = webdriver.ChromeOptions()
        self.option.add_argument("--incognito")
        self.driver = webdriver.Chrome(options=self.option)
        self.meses = {
            '01':'Enero',
            '02':'Febrero',
            '03':'Marzo',
            '04':'Abril',
            '05':'Mayo',
            '06':'Junio',
            '07':'Julio',
            '08':'Agosto',
            '09':'Setiembre',
            '10':'Octubre',
            '11':'Noviembre',
            '12':'Diciembre'
            }
    
        
    def go_potencias_contratadas(self):
        self.driver.maximize_window()
        self.driver.get(self.URL)
        xpath_mcp = '//*[@id="Mercado Mayorista/Liquidaciones del MME/01 Mercado de Corto Plazo/"]/td[3]'
        mcp_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
            By.XPATH, xpath_mcp)))
        mcp_button.click()
        sleep(4)
        
        xpath_potc = '//*[@id="Mercado Mayorista/Liquidaciones del MME/01 Mercado de Corto Plazo/Potencias Contratadas/"]/td[3]'
        potc_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
            By.XPATH, xpath_potc)))
        potc_button.click()
        sleep(4)
     
        
    def _recognize_last_file_version(self):
        body = self.driver.execute_script("return document.body")
        source = body.get_attribute('innerHTML')
        soup = BeautifulSoup(source,'html.parser')
        
        ##Identificar la etiqueta que contiene los archivos a descargar
        main_table = soup.find('table',id='tbDocumentLibrary').find('tbody')
        rows = main_table.find_all('tr')
        names_files = []
        num_version = []
        for tr in rows:
            names_files.append(tr.find_all()[4].get_text())
        
        for i,name in enumerate(names_files):
            try:
                num_version.append(int(re.findall(r'[0-9]+',names_files[i])[1])) 
            except:
                num_version.append(0)
        def findposMay(array):
            return (array.index(max(array)))
        
        file_term = names_files[findposMay(num_version)]
        return file_term
        
    
    def downloading_file(self,month,year):
        mes = self.meses[month]
        self.go_potencias_contratadas()
        #Identificando el año a entrar
        xpath_year = (r'//*[@id='
                     r'"Mercado Mayorista/'
                     r'Liquidaciones del MME/'
                     r'01 Mercado de Corto Plazo/'
                     r'Potencias Contratadas/'
                     fr'{year}/"]')
        xpath_month = (r'//*[@id='
                     r'"Mercado Mayorista/'
                     r'Liquidaciones del MME/'
                     r'01 Mercado de Corto Plazo/'
                     r'Potencias Contratadas/'
                     fr'{year}/'
                     fr'{month}_{mes}/'
                     r'"]')
        #ACCEDIENDO AL AÑO
        year_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
            By.XPATH, xpath_year)))
        year_button.click()
        sleep(4)
        #ACCEDIENDO AL MES
        month_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
            By.XPATH, xpath_month)))
        month_button.click()
        sleep(5)
        #IDENTIFICANDO LA ULTIMA VERSION DEL ARCHIVO
        file_term = self._recognize_last_file_version()
        
        xpath_download_file = (r'//*[@id='
                     r'"Mercado Mayorista/'
                     r'Liquidaciones del MME/'
                     r'01 Mercado de Corto Plazo/'
                     r'Potencias Contratadas/'
                     fr'{year}/'
                     fr'{month}_{mes}/'
                     fr'{file_term}'
                     r'"]')
       #DESCARGANDO EL ARCHVO EXCEL 
        download_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
            By.XPATH,xpath_download_file)))
        download_button.click()
        sleep(25)
        
        # IDENTIFICAR EL ULTIMO ARCHIVO DESCARGADO
        folder_path = r'C:\Users\rchavez\Downloads'#r'C:\Users\Toshiba\Downloads'
        file_type = r'\*xlsx'
        files = glob.glob(folder_path + file_type)
        max_file = max(files, key=os.path.getctime)
        # Movemos el archivo descargado a una carpeta dentro del proyecto
        source = f'{max_file}'
        ## Nuevo nombre de archivo
        destination = f"EXCEL_FILES/{month}_{self.meses[month]}_{year}.xlsx"
        shutil.move(source,destination)
       
        ## RETORNA A POTENCIAS CONTRATADAS
        xpath_pot_cont = '//*[@id="browserDocument"]/div[1]/ul/li[3]/a'
        pot_cont_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
        By.XPATH, xpath_pot_cont)))
        pot_cont_button.click()
        sleep(5)
        
        
    #DESCARGANDO EL ULTIMO ARCHIVO     
    def download_last_file_uploaded_on_web(self,cod_last_downloaded):
        self.go_potencias_contratadas()
        
        #RECUPERANDO HTML DE LA PAGINA
        body = self.driver.execute_script("return document.body")
        source = body.get_attribute('innerHTML')
        soup = BeautifulSoup(source,'html.parser')
        table = soup.find('table',id='tbDocumentLibrary')
        tr = table.find_all('tr')[1]
        td = tr.find_all('td')[2] 
        year = td.get_text()
        
        ### ACCEDEMOS AL AÑO MAYOR
        xpath_year = (r'//*[@id='
                     r'"Mercado Mayorista/'
                     r'Liquidaciones del MME/'
                     r'01 Mercado de Corto Plazo/'
                     r'Potencias Contratadas/'
                     fr'{year}/"]')
        
        year_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
            By.XPATH, xpath_year)))
        year_button.click()
        sleep(2)
        
        ### identificando el ultimo mes
        #RECUPERANDO HTML DE LA PAGINA
        body = self.driver.execute_script("return document.body")
        source = body.get_attribute('innerHTML')
        soup = BeautifulSoup(source,'html.parser')
        table = soup.find('table',id='tbDocumentLibrary')
        tr = table.find_all('tr')[1]
        td = tr.find_all('td')[2] 
        month = td.get_text()
        
        generated_code = month+'_'+year
        print("ULTIMO ARCHIVO SUBIDO A LA PAGINA:",generated_code)
        print("ULTIMO ARCHIVO DESCARGADO Y REGISTRADO",cod_last_downloaded)
        if generated_code != cod_last_downloaded:
            print('SE HA SUBIDO NUEVO ARCHIVO')
            print('INICIANDO LA DESCARGA DE NUEVO ARCHIVO')
            xpath_month = (r'//*[@id='
                         r'"Mercado Mayorista/'
                         r'Liquidaciones del MME/'
                         r'01 Mercado de Corto Plazo/'
                         r'Potencias Contratadas/'
                         fr'{year}/'
                         fr'{month}/'
                         r'"]')
            #ACCEDIENDO AL MES
            month_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
                By.XPATH, xpath_month)))
            month_button.click()
            sleep(5)
            ##DESCARGANDO EL ARCHIVO
            file_term = self._recognize_last_file_version()
            xpath_download_file = (r'//*[@id='
                         r'"Mercado Mayorista/'
                         r'Liquidaciones del MME/'
                         r'01 Mercado de Corto Plazo/'
                         r'Potencias Contratadas/'
                         fr'{year}/'
                         fr'{month}/'
                         fr'{file_term}'
                         r'"]')
            download_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((
                By.XPATH,xpath_download_file)))
            download_button.click()
            sleep(20)
            return generated_code,True
        else:
            return None,False
    '''
    def _get_last_code_file(self):
        f = open('LAST_DOWNLOADER_FILE.txt','r')
        code = f.read()
        f.close()
        return code
    '''   
    def FINISH_DOWNLOAD_FILE(self):
        f = open('LAST_DOWNLOADER_FILE.txt','r')
        last_code_file = f.read()
        f.close()
        uploaded = self.download_last_file_uploaded_on_web(last_code_file) 
        if uploaded[1]:
            #GUARDAMOS EL ARCHIVO EN UNA CARPETA DENTRO DEL PROYECTO
            #iDENTIFICANDO NOMBRE DE ARCHIVO
            folder_path = r'C:\Users\rchavez\Downloads'#r'C:\Users\Toshiba\Downloads'
            file_type = r'\*xlsx'
            files = glob.glob(folder_path + file_type)
            max_file = max(files, key=os.path.getctime)
            # Movemos el archivo descargado a una carpeta dentro del proyecto
            source = f'{max_file}'
            ## Nuevo nombre de archivo
            destination = f"EXCEL_FILES/{uploaded[0]}.xlsx"
            shutil.move(source,destination)
            with open('LAST_DOWNLOADER_FILE.txt','w') as f:
                f.write(uploaded[0])
        else:
            print("NO SE HA SUBIDO NUEVOS ARCHIVOS")
        return uploaded
        
        
    def CLOSE_DRIVER(self):
        self.driver.close()
        

        
        
        
    