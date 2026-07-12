import sys
import os
import datetime as dtm
import logging
import log_config
import PPTXreaderClass as prc
import EStoPPTX as esi


if __name__ == "__main__":

    log_config.set_log("logs/pptxaktualisieren.log")
    logger = logging.getLogger(__name__)
    logger.info('Logging von pptxaktualisieren main  gestartet')

    if len(sys.argv) < 2:
        print("No presentation given, using default:", "data/Default INPR.pptx")
        logger.info("No presentation given, using default:  : %s", sys.argv[0])
        filepath = "data/Default INPR.pptx"
    else:
        filepath = sys.argv[1]
        logger.info ("original file: %s", filepath)

    if os.path.exists(filepath):
        prs = prc.PPTXdataPresentation(filepath)
        # 'r'D:\01 Nextcloud\Documents\Programmieren\Powerpoint Generator\Powerpoint\Output 2025-02.pptx''
        logger.info ("Reader %s gestartet", filepath)
        logger.info('slides: %d', len(prs.slides))

        prs.update_charts()
        prs.save(os.path.join(os.path.dirname(filepath), dtm.datetime.now().strftime("_%Y-%m-%d_%H-%M-%S") + os.path.basename(filepath)))
    else:
        logger.info("Die Datei existiert nicht.")
