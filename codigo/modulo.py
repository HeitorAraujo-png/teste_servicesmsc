from services import Relatorio
from watchdog.events import FileSystemEventHandler
import settings
import os

class MyHandler(FileSystemEventHandler):

    def on_modified(self, event):
        if not event.is_directory:
            if len(os.listdir(settings.RELATORIO)) >= 2:
                path1, path2 = [i for i in os.listdir(settings.RELATORIO)]
                path1 = fr'{settings.RELATORIO}\{str(path1)}'
                path2 = fr'{settings.RELATORIO}\{str(path2)}'
                try:
                    print('tentou')
                    if path1.endswith(('xltx', 'xls', 'xlsm', '.xlsx')) and path2.endswith(('xltx', 'xls', 'xlsm', '.xlsx')):
                        print('entrou')
                        relatorio = Relatorio(path1, path2)
                        relatorio.Concatena()
                        relatorio.Converte()
                        relatorio.Espaco()
                except Exception as e:
                    if not os.path.exists(r"codigo\LOG\error.log"):
                        with open(r"codigo\LOG\error.log", 'x'):
                            with open(r"codigo\LOG\error.log", 'a') as log:
                                log.write(f'{settings.time} {e} {event.src_path} \n')
                    else:
                        with open(r"codigo\LOG\error.log", 'a') as log:
                                log.write(f'{settings.time} {e} {event.src_path} \n')
                for i in os.listdir(settings.RELATORIO):
                    os.remove(fr'{settings.RELATORIO}\{i}')
        else:
            if not os.path.exists(r"codigo\LOG\error.log"):
                    with open(r"codigo\LOG\error.log", 'x'):
                        with open(r"codigo\LOG\error.log", 'a') as log:
                            log.write(f'{settings.time}  Novo diretorio detectado! {event.src_path} \n')
            else:
                with open(r"codigo\LOG\error.log", "a") as log:
                    log.write(f"{settings.time} Novo diretorio detectado! {event.src_path} \n")