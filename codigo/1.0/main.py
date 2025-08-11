from settings import *
from watchdog.observers import Observer
from watchdog.events import LoggingEventHandler
from modulo import MyHandler
from time import sleep
import logging
import win32serviceutil
import win32service
import win32event
import servicemanager

class ServiceAuto(win32serviceutil.ServiceFramework):
    _svc_name_ = 'ServiceAuto'
    _svc_display_name_ = 'Serviço de automação'
    _svc_description_  = 'Teste de serviço de automação'
    
    
    def __init__(self, args): 
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True
        self.observer = None
        logging.basicConfig(
        filename=fr"{BASE_DIR}\LOG\dev.log",
        filemode="a",
        level=logging.INFO,
        format="%(asctime)s | %(process)d | %(message)s",
        datefmt="%d-%m-%y %H:%M:%S",
)
        
        
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.hWaitStop)
        if self.observer:
            self.observer.stop()
            self.observer.join()
        
        
    def SvcDoRun(self):
        servicemanager.LogInfoMsg(fr'Monitorando o diretorio: {BASE_DIR}{RELATORIO}')
        self.main()
            
    def main(self):
        try:
            path = rf"{BASE_DIR}\{RELATORIO}"
            log = LoggingEventHandler()
            event_handler = MyHandler()
            self.observer = Observer()
            self.observer.schedule(log, path, recursive=False)
            self.observer.schedule(event_handler, path, recursive=False)
            self.observer.start()
            while self.running:
                rc = win32event.WaitForSingleObject(self.hWaitStop, 100)
                if rc == win32event.WAIT_OBJECT_0:
                    break
        except Exception as erro:
            logging.exception("Erro no serviço: %s", erro)
            servicemanager.LogErrorMsg(str(erro))
        finally:
            self.observer.stop()
            self.observer.join()
            logging.info("Serviço encerrado com sucesso.")

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ServiceAuto)