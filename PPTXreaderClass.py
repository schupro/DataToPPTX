import os
import logging
import log_config
import datetime as dtm
import EStoPPTX as esi
from pptx import Presentation
#from pptx.enum.shapes import MSO_SHAPE
#from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
#from pptx.dml.color import RGBColor
from pptx.util import Inches, Cm
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

logger = logging.getLogger(__name__)


class PPTXdataPresentation:
    """
    Scans a Powerpoint file for slides containing charts and their associated dataset descriptions
    and calls the adequate ... toPPTX class to retrieve the data and update the chart.
    (currentls only Eurostat is implemented)

    searches all slides in  a power Point Presentation to find
    slides holding Placeholders for charts.
    If there is a chart, it checks the notes placeholder for a dataset code and a filter expression.
    to allow the Download of an Eurostat dataset.
        my_code holds the dataset-code and my_filter holds the filter expression to be used to
        retrieve the updated data.
        Source: gives the whole description of the dataset, as it is produced by the --
        For development purposes the function description prints some descriptions of the Chart.
    """

    def __init__(self, Filepath:str):
        self.prs = Presentation(pptx=Filepath)
        logger.info("PPTXdataPresentation Objekt erstellt")
        logger.info("mit Filepath: %s", Filepath)

    @property
    def slides(self):
        '''PPTX slides iterator via property, so calling code can use prs.slides.'''
        return self.prs.slides

    def update_charts(self):
        for sld in self.prs.slides:
            for pho in sld.placeholders:
                if pho.has_chart:
                    Chart = pho.chart
                    if sld.has_notes_slide:
                        logger.info('notes found on sld: %d', sld.slide_id)
                        logger.info('text found for query: %s', sld.notes_slide.notes_text_frame.text)
                        sld.pars = dict()
                        for par in sld.notes_slide.notes_text_frame.paragraphs:
                            match par.text.split('=')[0].strip():
                                case 'my_lang':
                                    my_lang =  par.text.split('=')[1]
                                    sld.pars['lang']=my_lang.strip()
                                    logger.info('LANG:  %s', my_lang)
                                case 'my_code':
                                    my_code = par.text.split('=')[1]
                                    sld.pars['code']=my_code.strip()
                                    logger.info('CODE:  %s', my_code)
                                case 'my_pars':
                                    my_pars = eval((par.text.split('=')[1]))
                                    sld.pars.update(my_pars)
                                    logger.info('FILTER:  %s', my_pars)
                                case 'my_time':
                                    my_time = eval((par.text.split('=')[1]))
                                    sld.pars.update(my_time)
                                    logger.info('TIME:  %s', my_time)

                        '''
                            Use the found parameters, to download the new data. Prepare the data to be inserted into the chart
                            First select the metadata for the chosen dataset via get toc and put them into a dictonary
                            use the fist list toc[0] containing the codes as key and the second list toc[1] containing
                            the Meta data for my_code as description for the dataseries.
                        '''


                        logger.info("received parameters: %s", sld.pars)

                        # Create an object of the EStoPptxData class with the found parameters
                        T = esi.EStoPptxData(sld.pars)
                        #T = esi.EStoPptxData(goodpars)
                        logger.info("chart data: %s", T.ChartData)

                        #self.description(Chart, show=True)
                    Chart.replace_data(T.ChartData)


        def chart_des(cls, my_chart:'object of the chart type of the pptx library', show = False) -> None:
            '''logs the chart type of the object, its Titel and whether it has a legend'''
            logger.info('type: %s', my_chart.chart_type)
            if my_chart.has_title:
                logger.info('the chart has the titel: %s', my_chart.tite)
            if my_chart.has_legend:
                logger.info('the chart has a Legend')
            if (my_chart.chart_type == XL_CHART_TYPE.LINE):
                logger.info('Holding the following %d Lines', my_chart.series.count)
            for L in my_chart.series:
                logger.info('Series name: %s', L.name)

    def save(self, Filepath:str):
        self.prs.save(Filepath)


if __name__ == "__main__":

    log_config.set_log("logs/PPTXreaderClass.log")
    logger = logging.getLogger(__name__)
    logger.info("Projektlog von PPTXreaderClass main gestartet")


    '''This is a test block to show the usage of the class
        goodpars = {"code": "STS_INPR_M",
                "flags": False,
                "indic_bt": ["PRD"],
                "nace_r2": ["C","D"],
                "s_adj": ["NSA"],
                "freq": ["M"],
                "unit": ["I21"],
                "geo": ["IT", "AT","DE"],
                "startPeriod": "2025-06",
                "endPeriod": "2025-07",
                "lang": "de"}
    '''
    Update=False
    filename = r'D:\01 Nextcloud\Documents\Programmieren\Powerpoint Generator\Powerpoint\Output Test1.pptx'
    prs = PPTXdataPresentation(filename)
    logger.info ("reader gestartet%s", filename)
    logger.info('slides: %d', len(prs.slides))
    prs.update_charts()
    prs.save(os.path.join(os.path.dirname(filename), dtm.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + os.path.basename(filename)))
    prs.save(filename)
