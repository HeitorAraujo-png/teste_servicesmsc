from os import *
from settings import BASE_DIR, RELATORIO

if len(listdir(path.join(BASE_DIR, RELATORIO))) >= 2:
    print("at")
