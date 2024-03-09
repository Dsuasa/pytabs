# PyTABS - ETABS .NET API python wrapper
# ETABS Configuration - ETABS API DLL handler

# general library imports
import pytabs.pytabs_config as config
import os
import clr

# read pytabs config file
pytabs_config = config.read_config()

# check to see if ETABS API DLL path from config file exists
etabs_api_path = pytabs_config['ETABS']['API_DLL_PATH']
if not os.path.isfile(etabs_api_path):
    raise FileNotFoundError(etabs_api_path)

# pythonnet clr-loader import of Marshal - ETABS API requirement
clr.AddReference("System.Runtime.InteropServices")
# import of ETABS API DLL path from config file
clr.AddReference(etabs_api_path)
import ETABSv1 as etabs
