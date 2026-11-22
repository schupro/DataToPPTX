import sys
import os
import datetime
import re
import logging
import log_config
import PPTXreaderClass as prc

log_config.set_log("logs/pptxaktualisieren.log")
logger = logging.getLogger(__name__)

if __name__ == "__main__":

    logger.info("Started by drag an drop exe file:")

    if len(sys.argv) < 2:
        filepath = "Default Graphs/Default.pptx"
        # print("No presentation given, using default:")
        logger.info("No presentation given, %s is using default %s" % (sys.argv[0], filepath))
    else:
        filepath = sys.argv[1]
        logger.info ("Original file: %s", filepath)

    if os.path.exists(filepath):
        logger.info ("start Reader on %s", filepath)
        prs = prc.PPTXdataPresentation(filepath)
        prs.update_charts()
        logger.info ("Charts updated:")
        new_filename = (datetime.datetime.now().strftime("%Y-%m-%d")+" "+
                        re.sub(r'[0-9]{4}-[0-2]{2}-[0-3][0-9]\s',
                               '',os.path.basename(filepath)))
        prs.save(os.path.join(os.path.dirname(filepath), new_filename))
        logger.info ("File saved to %s", os.path.join(os.path.dirname(filepath), new_filename))
    else:
        logger.info("Die Datei existiert nicht.")
