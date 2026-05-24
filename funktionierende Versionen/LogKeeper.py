# LogKeeper.py
import logging


def setup_logging(LogFile):
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(LogFile, mode="a", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(message)s"))
    logger.addHandler(handler)

if __name__ == "__main__":
    setup_logging("LogKeeperTest.log")
    logger = logging.getLogger(__name__)
    logger.info("LogKeeperTest wird beim start der LogKeeper Main initialisiert")
