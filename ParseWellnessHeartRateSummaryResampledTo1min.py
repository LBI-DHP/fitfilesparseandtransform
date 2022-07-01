import fitparse  # parse ANT/Garmin .FIT files
import pandas  # data analysis and manipulation
import sys  # System-specific parameters and functions
import os  # using operating system dependent functionality
from os import walk  # generating directory tree
import datetime as dt


def getSwVersion(fitfile):  # defining a function
    # get list of DataMessage record objects that are contained in this FitFile
    for element in fitfile.get_messages():
        for d in element:
            if d.name == "software_version":
                return d.value


def getFileType(fitfile):
    for element in fitfile.get_messages():
        for d in element:
            if d.name == "type":
                return d.value


def getTimeOffset(fitfile):
    local_timestamp = 0
    unix_timestamp = 0
    for message in fitfile.get_messages("monitoring_info"):
        for data in message:
            if(data.name == "local_timestamp"):
                local_timestamp = data.value
            if(data.name == "timestamp"):
                unix_timestamp = data.value

    if local_timestamp != 0 and unix_timestamp != 0:
        time_offset = local_timestamp - unix_timestamp
        return time_offset.total_seconds()
    return 0


def appendToDataframe(data_frame_name, fitfile, time_offset, dataframe):
    unix_datetime = 0
    for message in fitfile.get_messages(data_frame_name):
        newRow = {}
        hasHeartRate = False
        for data in message:
            if(data.name == "timestamp"):
                unix_datetime = data.value
                local_datetime = unix_datetime + \
                    dt.timedelta(seconds=time_offset)
                newRow["utc_time"] = unix_datetime
                newRow["local__time"] = local_datetime
            elif(data.name == "timestamp_16" and unix_datetime != 0):
                # https://stackoverflow.com/questions/57774180/how-to-handle-timestamp-16-in-garmin-devices
                unix_timestamp_garmin_epoch = int(
                    dt.datetime.timestamp(unix_datetime)) - 631065600
                timestamp_16 = data.value
                calculated_timestamp_garmin_epoch = unix_timestamp_garmin_epoch + \
                    ((timestamp_16 - unix_timestamp_garmin_epoch) & 0xffff)
                local_calculated_datetime = dt.datetime.utcfromtimestamp(
                    calculated_timestamp_garmin_epoch + 631065600 + time_offset)
                newRow["timestamp_16"] = timestamp_16
                newRow["local_time_calculated_with_timestamp_16"] = local_calculated_datetime
            elif data.name == "heart_rate":
                row_name = data.name
                if(data.units):
                    row_name = data.name + " [" + data.units + "]"
                newRow[row_name] = data.value
                hasHeartRate = True
        if hasHeartRate:
            dataframe = dataframe.append(newRow, ignore_index=True)
    return dataframe


IMPORT_FOLDER = "import/"
OUTPUT_FOLDER = "export/"
# retrieving file names from the files in the input folder
(_, _, filenames) = next(walk(IMPORT_FOLDER))

dataframe = pandas.DataFrame()

for filename in filenames:
    print("=============================================\n")
    print("Parsing ", filename)
    fitfile = fitparse.FitFile(IMPORT_FOLDER+filename)
    sw_version = getSwVersion(fitfile)
    file_type = getFileType(fitfile)

    if(file_type == "monitoring_b"):
        time_offset = getTimeOffset(fitfile)
        dataframe = appendToDataframe(
            "monitoring", fitfile, time_offset, dataframe)
    else:
        print("ERROR: This script converts files of type \"wellness / monitoring_b\" only.")
        print("Files of type \"", file_type, "\" can not be converted.")

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)


dataframe.index = pandas.to_datetime(
    dataframe.local_time_calculated_with_timestamp_16, format='%Y-%m-%D %H:%M:%S')
dataframe = dataframe.resample("1min").asfreq()

print("=============================================\n")
try:
    dataframe.to_excel(OUTPUT_FOLDER + "SUMMARY_RESAMPLED_1MIN" + ".xlsx")
    print("SUCCESS: The converted file is located in the \"export\" folder.")
except:
    print("ERROR: File could not be saved.")
    print("If the file to be exported already exists, please make sure that the file is not open (e.g. in Excel) when running this script.")
print("")
