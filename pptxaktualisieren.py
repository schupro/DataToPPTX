import sys
import os
import datetime as dtm
import logging
import log_config
import PPTXreaderClass as prc

log_config.set_log("logs/pptxaktualisieren.log")
logger = logging.getLogger(__name__)

if __name__ == "__main__":

    logger.info("Started by drag an drop exe file:")

    if len(sys.argv) < 2:
        filepath = "data/Default INPR.pptx"
        print("No presentation given, using default:")
        logger.info("No presentation given, %s is using default %s" % (sys.argv[0], filepath))
    else:
        filepath = sys.argv[1]
        logger.info ("Original file: %s", filepath)

    if os.path.exists(filepath):
        logger.info ("Reader %s staren", filepath)
        prs = prc.PPTXdataPresentation(filepath)
        prs.update_charts()
        logger.info ("Charts updated:")
        prs.save(os.path.join(os.path.dirname(filepath), dtm.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + os.path.basename(filepath)))

    else:
        logger.info("Die Datei existiert nicht.")