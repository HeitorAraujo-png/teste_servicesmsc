from services import Relatorio
from watchdog.events import FileSystemEventHandler
import settings
from shutil import move
from collections import defaultdict
import os
import time

class MyHandler(FileSystemEventHandler):

    def __init__(self):
        self.last_events = defaultdict(float)
        self.debounce_interval = 5  # segundos

    
    def on_modified(self, event):
        if not event.is_directory:
            now = time.time()
            if now - self.last_events[event.src_path] < self.debounce_interval:
                return                
            self.last_events[event.src_path] = now

            ListaPath = []
            if len(os.listdir(settings.RELATORIO)) >= 2:
                ListaPath = os.listdir(settings.RELATORIO); ListaPath.remove('UPLOAD_PROCESSADOS')
                try:
                    relatorio = Relatorio(ListaPath)
                    relatorio.Converte()
                    relatorio.Espaco()
                    for i in ListaPath:
                        reme = fr'{settings.RELATORIO}\{i}'
                        dest = fr"{settings.UPLOAD_PROCESSADO}\{i}"
                        move(src=reme, dst=dest)
                except Exception as e:
                    print(e)
                    if not os.path.exists(r"codigo\LOG\error.log"):
                        with open(r"codigo\LOG\error.log", "x"):
                            with open(r"codigo\LOG\error.log", "a") as log:
                                log.write(f"{settings.time} {e} {event.src_path} \n")
                    else:
                        with open(r"codigo\LOG\error.log", "a") as log:
                            log.write(f"{settings.time} {e} {event.src_path} \n")
        else:
            if not os.path.exists(r"codigo\LOG\error.log"):
                with open(r"codigo\LOG\error.log", "x"):
                    with open(r"codigo\LOG\error.log", "a") as log:
                        log.write(
                            f"{settings.time}  Novo diretorio detectado! {event.src_path} \n"
                        )
            else:
                with open(r"codigo\LOG\error.log", "a") as log:
                    log.write(
                        f"{settings.time} Novo diretorio detectado! {event.src_path} \n"
                    )
