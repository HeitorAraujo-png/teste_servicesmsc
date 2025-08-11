from datetime import datetime
import os

time = datetime.now().strftime("%d-%m-%y | %H:%M:%S | ")
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOADCSV = r"\..\UPLOADS\DOWNLOAD_RELATORIOS\csv"
UPLOADXLSX = r"\..\UPLOADS\DOWNLOAD_RELATORIOS\xlsx"
UPLOAD_PROCESSADO = r'\..\UPLOADS\UPLOAD_RELATORIOS\UPLOAD_PROCESSADOS' 
RELATORIO = r"\..\UPLOADS\UPLOAD_RELATORIOS"
HISTORICO = r'\..\UPLOADS\DOWNLOAD_RELATORIOS\xlsx\HISTORICO'