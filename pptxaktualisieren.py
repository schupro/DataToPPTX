import sys
import os
import datetime as dtm
import logging
import log_config
import PPTXreaderClass as prc


if __name__ == "__main__":

    log_config.set_log("logs/pptxaktualisieren.log")
    logger = logging.getLogger(__name__)
    logger.info("Started by drag an drop exe file:")

    if len(sys.argv) < 2:
        print("No presentation given, using default:")
        logger.info("No presentation given, using default:  data/Default INPR.pptx", sys.argv[0])
        filepath = "data/Default INPR.pptx"
    else:
        filepath = sys.argv[1]
        logger.info ("Original file: %s", filepath)

    if os.path.exists(filepath):
        prs = prc.PPTXdataPresentation(filepath)
        # 'r'D:\01 Nextcloud\Documents\Programmieren\Powerpoint Generator\Powerpoint\Output 2025-02.pptx''
        logger.info ("Reader %s gestartet", filepath)
        logger.info('slides: %d', len(prs.slides))

        prs.update_charts()
        prs.save(os.path.join(os.path.dirname(filepath), dtm.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + os.path.basename(filepath)))
    else:
        logger.info("Die Datei existiert nicht.")