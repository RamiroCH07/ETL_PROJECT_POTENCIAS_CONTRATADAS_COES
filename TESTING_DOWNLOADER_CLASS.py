from DOWNLOADER import Downloader_files

#%%

## TESTEANDO FUNCIONALIDADES BASCIAS DE LA CLASE DOWNLOADER


# OBJETIVO: VERIFICAR LA CORRECTA FUNCIONALIDAD DEL MÉTODO"_go_potencias_contratadas"
obj_downloader = Downloader_files()
obj_downloader.go_potencias_contratadas()

### TESTEANDO LA FUNCIONALIDAD DE DESCARGA DE ARCHIVOS EXCEL DE LA CLASE DOWNLOADER
#%%
# OBJETIVO: DESCARGAR EL ARCHIVO DEL MES ENERO DEL 2023
obj_downloader = Downloader_files()
obj_downloader.downloading_file('01', '2023')
obj_downloader.CLOSE_DRIVER()

#%%
# OBJETIVO: DESCARGAR TODOS LOS ARCHIVOS DEL AÑO 2022
obj_downloader = Downloader_files()
months = ['01','02','03','04','05','06','07','08','09','10','11','12']
year = '2022'

for month in months:
    obj_downloader.downloading_file(month, year)


#%%%
# OBJETIVO: VERIFICAR Y DESCARGAR EL ULTIMO ARCHIVO SUBIDO A LA PÁGINA DEL COES
obj_downloader = Downloader_files()
obj_downloader.FINISH_DOWNLOAD_FILE()
obj_downloader.CLOSE_DRIVER()