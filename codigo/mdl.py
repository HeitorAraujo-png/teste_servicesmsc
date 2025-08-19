from watchdog.events import FileSystemEventHandler
from Utilidades import Querys, Utils
from collections import defaultdict
from datetime import datetime
from shutil import move
import pandas as pd
from scv import Xlsx
from path import *
import time

class Handler(FileSystemEventHandler):

    def __init__(self):
        self.last_events = defaultdict(float)
        self.debounce_interval = 3  # segundos
    
    def on_modified(self, event):
        if not event.is_directory:
            now = time.time()
            if now - self.last_events[event.src_path] < self.debounce_interval:
                return                
            self.last_events[event.src_path] = now
            try:
                query = Querys()
                Querys.Concatena(event.src_path)
                name = fr'{BASE_DIR}{UPLOADCSV}\DptTemp{datetime.now().today().strftime('%d-%m-%y')}.csv'
                xlsx = pd.read_excel(event.src_path, parse_dates=['DATA'])
                xlsx["DATA"] = xlsx["DATA"].dt.strftime("%d/%m/%y")
                xlsx.to_csv(name, index=False, encoding='latin1')
                csv = pd.read_csv(name, encoding='latin1')
                for codigo, cc, dia, hora in zip(csv.COD, csv['C.C'], csv.DATA, csv.HORA):
                    hora = hora[7:12]
                    query.Create(CenCus=cc, day=dia, hora=hora, cd=codigo)
                move(src=event.src_path, dst=f'{BASE_DIR}{UPLOAD_PROCESSADO}')
                Xlsx()
            except Exception as e:
                Utils.Log(f"{time} {e} {event.src_path}")
        else:
            Utils.Log(f"{time} Novo diretorio detectado! {event.src_path}")