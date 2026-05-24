import eurostat  # https://pypi.org/project/eurostat/
import pandas as pd
import os
from pathlib import Path
import logging
from pptx.chart.data import CategoryChartData


class EStoPptxData:

    """
    Loads Data from Eurostat and returns a Powerpoint Chart data objects

    Attributes
    ----------
    query_params: dict a dictionary
            Containing the
             code: the Eurstat code of the dataset
             par_code1: filter_value, name of the
                        1. parameter and the chosen ["values"]
                         .
                         .
             par_code_n: filter_value name of the n_th.
                parameter and the chosen ["values"] for missing
                Parameter all versions are retourned
             startPeriod: If no value, defaults to first available data point.
             endPeriod:	If no value, defaults to last available data point
             data_frame:returns the data frame as it came from Eurostat

             params: returns the query params.
             chart_data(self): Returns a chart data Object with Data from the query
             descr: A list of strings describing the dataset and the selected parameters


    Methods:
    -----------

    """

    def __init__(self, query_par={"lang": "de","code": "STS_INPR_M","flags": False, }):
        self.par = query_par
        buffer = query_par.copy()
        self.code = buffer.pop("code")
        self.lang = buffer.pop("lang")
        self.filter = buffer
        toc = eurostat.get_toc(dataset=self.code, lang=self.lang)
        #print(toc[0], toc[1])
        self.meta_dict = dict(zip(toc[0], toc[1]))
        #print("der inhalt",self.meta_dict)

        self.descr =[
            self.meta_dict["code"],
            self.meta_dict["title"],
            self.meta_dict["last update of data"],
            self.meta_dict["data end"]]

        self.data = pd.DataFrame()
        self.data = eurostat.get_data_df(code=self.code,filter_pars=self.filter)
        self.raw_data=self.data.copy()
        #this copy holds the raw data as downloaded, it ist not effected by further processing

        self.cols = self.data.columns.values.tolist()
        self.split_idx = next((i for i, col in enumerate(iterable=self.cols) if "\\" in col), len(self.cols))
        '''
        split index (split_idx) counts the collumns until the heading contains the
         \\ character whicht marks the last collumn of the series parameters
        '''

        self.data.rename(columns=lambda x: x.partition("\\")[0], inplace=True)

        self.params = self.cols[:self.split_idx+1]
        self.params[self.split_idx] = self.params[self.split_idx].partition("\\")[0]
        self.values = self.cols[self.split_idx+1:]

        # self.medium_data = self.data[self.values]
            #drop(labels=self.params, axis=1)
        # self.medium_names =self.data[self.params]

        ''' get general descriptions and specific ones for the Parameters '''
        self.data.insert(loc=0, column="name", value="")
        '''create a column for the series names'''

        for par in self.params:
            # Get the descriptions for the Parameter items from Eurostat
            item_dict = eurostat.get_dic(code=self.code, par=par,  full=False, frmt="dict", lang=self.lang)
           # check if a downloaded dimension features only a single characteristic
            if (len(self.data [par].unique()) == 1):
                #if so put the description from the dictionary into the general description of the chart
                # and delete the collumn holding the unique value from the data frame, and the par Group
                self.descr.append(item_dict[self.data[par][0]])
                self.data.pop(par)
            else:
                # if there is more than one value ad a new collumn at the start of the dataset
                # create a description for the series containing the TXT and the Parameter and
                # the description for each item from the dictionary
                self.data['name'] = self.data['name']+self.data[par].map(item_dict) + " (" + self.data[par] + ")"
                self.data.pop(par)
                # finally remove the par collumn from the data frame

        ''' Creates a PowerPoint Chart Data Object from the prepared data frame'''
        #self.data=self.data.drop(labels=self.params, axis=1,inplace=False)
        self.ChartData = CategoryChartData()  #create the data object for the PowerPoint Chart
        self.ChartData.categories = list(self.data.columns)[1:]

        for index, row in self.data.iterrows():
            # print((row.count))
                self.ChartData.add_series(row["name"],
                                          list(row[self.values].fillna('')))
            # adds series line by line to the Chart using the series names deduced from the data
            # as Series Names and the Value colums as Data



if __name__ == "__main__":
    '''This is a test block it creates an object of the EStoPptxData class with a predefined parameter set
        and prints the results of the Eurostat data retrieval to a log file.'''

    my_pars = {"lang": "de",
                "code": "STS_INPR_M",
                "flags": False,
                "indic_bt": ["PRD"],
                "nace_r2": ["C","D"],
                "s_adj": ["NSA"],
                "freq": ["M"],
                "unit": ["I21"],
                "geo": ["HR", "AT"],
                "startPeriod": "2025-08",
                "endPeriod": "2025-09"}

    cwd = Path.cwd()
    print(cwd)
    LogFile = cwd / "ESktoPPts.log"
    logging.basicConfig(filename= LogFile,
                    format='%(asctime)s %(levelname)s: %(message)s',
                    filemode='w')

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.info("gestartet")

    # Hier wird ein Objekt der Klasse EStoPptxData mit den in my_pars definierten
    # Parametern erstellt,
    # die Daten von Eurostat abgerufen und in ein PowerPoint Chart Data Objekt umwandelt.

    T = EStoPptxData(my_pars)

    # logger.info(T.__dict__)
    logger.info(f"Serien code {T.code}")

    logger.info(f"Filter Parameter:  {T.par}")
    logger.info(f"Beschreibung: {T.descr}")
    logger.info(f"Rohdaten: {T.raw_data}")
    # logger.info("Chart data: ", T.medium_data)
    logger.info(f"Params: {T.params}")
    logger.info(f"data: {T.data}")
    logger.info(f"Values: {T.values}")
    logger.info(f"Names: {T.ChartData}")
