from datetime import datetime
from os import getcwd
time = datetime.now().strftime("%d-%m-%y | %H:%M:%S | ")
BASE_DIR = getcwd()
UPLOADCSV = r'UPLOADS\DOWNLOAD_RELATORIOS\csv'
UPLOADXLSX = r'UPLOADS\DOWNLOAD_RELATORIOS\xlsx'
RELATORIO = r'UPLOADS\UPLOAD_RELATORIOS'
LOG = r'codigo\LOG'
ERROR_LOG= r"codigo\LOG\error.log"