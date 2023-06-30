from DOWNLOADER import Downloader_files

downloader = Downloader_files()
#DOWNLOADER
#DESCARGA DEL AARCHVO SUBIDO AL COES
downloader.FINISH_DOWNLOAD_FILE()
#EXTRACTER



#%%

#with open("TEST.html","w",encoding = "utf-8") as f:
#    f.write(html)
#%%

print(r'HOLA MUNDO ' 
      r'hola mundo ' 
      r'HI WOrdl')    

year = 789
xpath_year = (r'//*[@id='
             r'Mercado Mayorista/'
             r'Liquidaciones del MME/'
             r'01 Mercado de Corto Plazo/'
             r'Potencias Contratadas/'
             fr'{year}')
             
print(xpath_year)      

#%%

test = ['2023', '2022', '2021', '2020', '2019', '2018']

#%%

with open('LAST_DOWNLOADER_FILE.txt','w',encoding = 'utf-8-sig') as f:
    f.write('06_junio_2023')
    
    #%%
    
f = open('LAST_DOWNLOADER_FILE.txt','r')
code = f.read()
f.close()
print(code)
#%%
def __get_last_code_file():
    f = open('LAST_DOWNLOADER_FILE.txt','r')
    code = f.read()
    f.close()
    return code

print(__get_last_code_file())