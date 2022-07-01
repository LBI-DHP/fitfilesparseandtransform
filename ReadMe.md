# Python Scripts for Converting Garmin Wellness and Activity Fit Files to Excel Files

## Get Started

These scripts have been developed and tested (with python 3.9.6) for a specific use case. Therefore, they must always be tested (for your specific use case) before use.

### Install Python

This set of scripts is written in python. To run them, you must first install python.

[Download Python here: https://www.python.org/downloads](https://www.python.org/downloads/)

To check if python has been installed correctly run `python --version` in your Windows Powershell/Terminal or MacOS Terminal.

### Install Packages with pip

Pip is a package manager for python and allows convenient management of python software packages. If you installed python from source, using an installer from python.org (as recommended above), you should already have pip installed.

To check if pip is installed run `pip --version` in your Windows Powershell/Terminal or MacOS Terminal.

Software packages provide specific functionality, such as reading fit files or analyzing data. Run the following command in your Windows Powershell/Terminal or MacOS Terminal to install the required packages:

`pip install fitparse pandas openpyxl`

## Run Scripts

To run the scripts, save them somewhere on your computer. Open the path where you saved the scripts using Windows Powershell/Terminal or MacOS Terminal. The easiest way to do so is to navigate to the correct path using your file explorer and then right-click "Open in Windows-Terminal" or "Open in Terminal".

Place the fit files you want to convert in the import folder located in the same directory as the scripts. You must not move this folder to a different location than the python scripts.

### Convert Garmin Activity Fit Files to Excel Files

To convert Garmin Activity Fit files, run the following command in your Windows Powershell/Terminal or MacOS Terminal (make sure you are on the correct path as described above):

`python .\ParseActivityFitFiles.py`

Garmin records data at irregular intervals (every few seconds). To resample the data to measurements every second, run the following command (rows of seconds without source data will remain empty):

`python .\ParseActivityFitFiles.py resample`

Find the converted files in the "export" folder.

### Convert Garmin Wellness Fit Files to Excel Files

To convert Garmin Wellness Fit files, run the following command in your Windows Powershell/Terminal or MacOS Terminal (make sure you are on the correct path as described above):

`python .\ParseWellnessFitFiles.py`

Find the converted Excel files in the "export" folder.

### Convert Multiple Garmin Wellness Fit Files to one Excel Summary File (heart rate only, resampled to 1 minute)

To convert Garmin Wellness Fit files, run the following command in your Windows Powershell/Terminal or MacOS Terminal (make sure you are on the correct path as described above):

`python .\ParseWellnessHeartRateSummaryResampledTo1min.py`

Find the converted Excel files in the "export" folder.

