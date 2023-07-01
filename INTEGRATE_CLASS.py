from DOWNLOADER import Downloader_files
from ESTRACTER import Estracter_data
class INTEGRATE_CLASS:
    def __init__(self):
        self.obj_downloader = Downloader_files()
    

    
    def START_ETL_PROCCESS_JOB(self):
        uploaded = self.obj_downloader.FINISH_DOWNLOAD_FILE()
        self.obj_downloader.CLOSE_DRIVER()
        if uploaded[1]:
            obj_estracter = Estracter_data(uploaded[0])
            print("ARCHIVO DESCARGADO Y LISTO PARA LA EXTRACCION")
            print("EXTRAYENDO DATOS DEL ARCHIVO EXCEL EN CUESTION...........")
            return obj_estracter.EXTRACT_ALL_DATA_FROM_EXCEL()
        else:
            return None

            
