from settings import *
from watchdog.observers import Observer
from watchdog.events import LoggingEventHandler
from modulo import MyHandler
from time import sleep
import logging
if __name__ == '__main__':
    logging.basicConfig(
    filename=r"codigo\LOG\dev.log",
    filemode="a",
    level=logging.INFO,
    format="%(asctime)s | %(process)d | %(message)s",
    datefmt="%d-%m-%y %H:%M:%S",
    )
    path = fr'{BASE_DIR}\{RELATORIO}'
    log = LoggingEventHandler()
    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(log, path, recursive=False)
    observer.schedule(event_handler, path, recursive=False)
    observer.start()
    try:
        while True:
            sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        observer.join() 