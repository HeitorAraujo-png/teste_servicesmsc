from watchdog.events import LoggingEventHandler
from watchdog.observers import Observer
from Utilidades import Utils
import win32serviceutil
from mdl import Handler
import servicemanager
import win32service
from path import *
import win32event
    

class ServiceAuto(win32serviceutil.ServiceFramework):
    _svc_name_ = 'Automação'
    _svc_display_name_ = 'Serviço de automação'
    _svc_description_  = 'Serviço de automação com DB sqlite'
    
    
    def __init__(self, args): 
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True
        self.observer = None
        
        
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.hWaitStop)
        if self.observer:
            self.observer.stop()
            self.observer.join()
        
        
    def SvcDoRun(self):
        Utils.Log(fr'Monitorando o diretorio: {BASE_DIR}{RELATORIO}')
        self.main()
            
    def main(self):
        try:
            path = rf"{BASE_DIR}\{RELATORIO}"
            log = LoggingEventHandler()
            event_handler = Handler()
            self.observer = Observer()
            self.observer.schedule(log, path, recursive=False)
            self.observer.schedule(event_handler, path, recursive=False)
            self.observer.start()
            while self.running:
                rc = win32event.WaitForSingleObject(self.hWaitStop, 100)
                if rc == win32event.WAIT_OBJECT_0:
                    break
        except Exception as erro:
            Utils.Log("Erro no serviço: %s", erro)
            servicemanager.LogErrorMsg(str(erro))
        finally:
            self.observer.stop()
            self.observer.join()
            Utils.Log("Serviço encerrado com sucesso.")

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ServiceAuto)