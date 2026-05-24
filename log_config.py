#log_config.py
import logging
from logging.handlers import RotatingFileHandler
import os

def set_log(logfile: str, level = logging.INFO):
    '''
    Initialiszies the central logging for the project.
    It creates a log file with the given name and sets the logging level to INFO by default.
    '''
    logger = logging.getLogger()
    logger.setLevel(level)

    # if a logger is already configured, return the logger without reconfiguring it
    if logger.handlers:
        return logger

    # Make sure the directory and logfile exists if not create gthe loggfile in the current directory
    logdir = os.path.dirname(logfile)
    if logdir and not os.path.exists(logdir):
        os.makedirs(logdir)

    formatter = logging.Formatter("%(asctime)s - %(name)s - %(message)s",  "%Y-%m-%d %H:%M:%S")

    # File Handler
    fh = RotatingFileHandler(logfile, mode="w" , maxBytes=5*1024*1024, backupCount=2, encoding="utf-8")
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    return logger

if __name__ == "__main__":

    set_log("logs/Log_config_test.log")
    logger = logging.getLogger(__name__)
    logger.info("Dieser Log wurde beim start der log_config initialisiert")

