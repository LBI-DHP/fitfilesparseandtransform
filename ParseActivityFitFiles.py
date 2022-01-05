import fitparse # parse ANT/Garmin .FIT files
import pandas # data analysis and manipulation 
import sys # System-specific parameters and functions
import os # using operating system dependent functionality
from os import walk # generating directory tree
import datetime as dt

def getSwVersion(fitfile): # defining a function
    for element in fitfile.get_messages(): # get list of DataMessage record objects that are contained in this FitFile
        for d in element:
            if d.name == "software_version":
                return d.value

def getFileType(fitfile): 
    for element in fitfile.get_messages(): 
        for d in element:
            if d.name == "type":
                return d.value

def getTimeOffset(fitfile):
    for message in fitfile.get_messages("device_settings"):
        for data in message:
            if(data.name == "time_offset"):
                return data.value
    return 0

def appendToDataframe(data_frame_name, fitfile, time_offset, dataframe):
    for message in fitfile.get_messages(data_frame_name):
        newRow = {}
        for data in message:
            if(data.name == "timestamp"):
                utc_datetime = data.value
                local_datetime = utc_datetime + dt.timedelta(seconds=time_offset)
                newRow["utc_time"] = utc_datetime
                newRow["local_time"] = local_datetime
            elif not data.name.startswith("unknown"):
                newRow[data.name] = data.value
        dataframe = dataframe.append(newRow, ignore_index=True)
    return dataframe

IMPORT_FOLDER="import/"
OUTPUT_FOLDER="export/"
(_, _, filenames) = next(walk(IMPORT_FOLDER)) #retrieving file names from the files in the input folder

for filename in filenames:
    print ("=============================================\n")
    print("Parsing ",filename)
    fitfile = fitparse.FitFile(IMPORT_FOLDER+filename) 
    sw_version = getSwVersion(fitfile)
    file_type = getFileType(fitfile)

    resample = False

    if(len(sys.argv) == 2):
        if(sys.argv[1] == "resample"):
            resample = True

    if(file_type == "activity"): 

        time_offset = getTimeOffset(fitfile)
        dataframe = pandas.DataFrame()
        dataframe = appendToDataframe("record", fitfile, time_offset, dataframe)
        
        if(resample):
            dataframe.index = pandas.to_datetime(dataframe.local_time, format='%Y-%m-%D %H:%M:%S')
            dataframe = dataframe.resample("1S").asfreq()

        if not os.path.exists(OUTPUT_FOLDER):
            os.makedirs(OUTPUT_FOLDER)

        try:
            dataframe.to_excel(OUTPUT_FOLDER+ os.path.splitext(filename)[0]+"_fitfilev"+str(sw_version)+".xlsx")
            print("SUCCESS: The converted file is located in the \"export\" folder.")
        except:
            print("ERROR: File could not be saved.")
            print("If the file to be exported already exists, please make sure that the file is not open (e.g. in Excel) when running this script.") 

    else: 
        print("ERROR: This script converts files of type \"activity\" only.")
        print("Files of type \"", file_type, "\" can not be converted.")

    print("")