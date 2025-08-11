from datetime import datetime
import os

time = datetime.now().strftime("%d-%m-%y | %H:%M:%S | ")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEV_LOG = r"\LOG\dev.log"
UPLOADCSV = r"\..\ARQUIVOS\DOWNLOAD_RELATORIOS\csv"
UPLOADXLSX = r"\..\ARQUIVOS\DOWNLOAD_RELATORIOS\xlsx"
UPLOAD_PROCESSADO = r'\..\ARQUIVOS\UPLOAD_RELATORIOS\UPLOAD_PROCESSADOS' 
RELATORIO = r"\..\ARQUIVOS\UPLOAD_RELATORIOS"
HISTORICO = r'\..\ARQUIVOS\DOWNLOAD_RELATORIOS\xlsx\HISTORICO'
LINK = rf"{BASE_DIR}{UPLOADCSV}\DPTDIA.csv"