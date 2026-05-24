import sys
import os
import logging
import log_config
import PPTXreaderClass as prc
import EStoPPTX as esi


if __name__ == "__main__":

    log_config.set_log("logs/pptxaktualisieren.log")
    logger = logging.getLogger(__name__)
    logger.info('Logging von pptxaktualisieren main  gestartet')

    if len(sys.argv) < 2:
        print("sys.argv: ", sys.argv)
        logger.info("Bitte ziehe eine Datei auf dieses Script.  : {sys.argv[1]}")
        sys.exit(1)

    filepath = sys.argv[1]
    logger.info ("original file: %s", filepath)

    if os.path.exists(filepath):
        prs = prc.PPTXdataPresentation(filepath)
        # 'r'D:\01 Nextcloud\Documents\Programmieren\Powerpoint Generator\Powerpoint\Output 2025-02.pptx''
        logger.info ("Reader %{filepath} gestartet")
        logger.info('slides: %d', len(prs.slides))

        prs.update_charts()

        prs.save(filepath)
    else:
        logger.info("Die Datei existiert nicht.")
