# SDMX Reader
# reading sdmx datat to create chart data

'''
potential packages are:
pip install SDMXthon 
documented at: https://docs.sdmxthon.meaningfuldata.eu/index.html
pip install pandasdmx 
documented at: https://pandasdmx.readthedocs.io/en/v1.0/
pip install eurostat https://pypi.org/project/eurostat/
I gave pandasdmx a first try as it seems well sutied havin 20 predefined
data providers among them Eurostat and the ECB
'''

import pandasdmx as sdmx

# Requesting a the datadirectionary offerd by a provider
ecb = sdmx.Request('ECB')

# Open the Page describing the data (more precise the url given in the data message)
# ecb.view_doc() 

flow_msg = ecb.dataflow()
#dataflows = sdmx.to_pandas(flow_msg.dataflow)

ame_msg = ecb.dataflow('AME')
dataflows = sdmx.to_pandas(ame_msg.dataflow)

'''
#exr_msg.response.url
#exr_flow = exr_msg.dataflow.EXR
#dsd = exr_flow.structure
#print(dsd.dimensions.components)
#dsd.attributes.components
#dsd.measures.components
cl = dsd.dimensions.get('FREQ').local_representation.enumerated
'''
print("done")