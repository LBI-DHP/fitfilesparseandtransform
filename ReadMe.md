# Python scripts for converting Garmin Wellness and Activity Fit files to CSV files

## Get started

### Install Python

This set of scripts is written in python. To run them, you must first install python. The scripts were tested with python 3.9.6, but should also work with newer versions.

[Download Python here: https://www.python.org/downloads](https://www.python.org/downloads/)

To check if python has been installed correctly run `python --version` in your Windwows Powershell/Terminal or MacOS Terminal.

### Install Packages with pip

Pip is a package manager for python and allows convenient management of python software packages. If you installed python from source, using an installer from python.org (as recommended above), you should already have pip installed.

To check if pip is intalled run `pip --version` in your Windwows Powershell/Terminal or MacOS Terminal.

Software packages provide specific functionality, such as reading fit files or analysing data. Run the following comand in your Windwows Powershell/Terminal or MacOS Terminal to install the required packages:

`pip install fitparse pandas openpyxl`

## Run Scripts

To run the scripts, save them somewhere on your computer. Open the path where you saved the scripts using Windwows Powershell/Terminal or MacOS Terminal. The easiest way to do so is to navigate to the correct path using your file explorer and then right-click "Open in Windows-Terminal" or "Open in Terminal".

Place the fit files you want to convert in the "import" folder. You must not move this folder to a different location than the python scripts.

### Converting Garmin Activity Fit files to CSV files

To convert Garmin Activity Fit files, run the following command in your Windwows Powershell/Terminal or MacOS Terminal (make sure you are on the correct path as described above):

`python .\ParseActivityFitFiles.py`

Garmin records data at irregular intervals (every few seconds). To resample the data to measurements every second, run the following command (rows of seconds without source data will remain empty):

`python .\ParseActivityFitFiles.py resample`

Find the converted files in the "export" folder.

### Converting Garmin Wellness Fit files to CSV files

To convert Garmin Wellness Fit files, run the following command in your Windwows Powershell/Terminal or MacOS Terminal (make sure you are on the correct path as described above):

`python .\ParseWellnessFitFiles.py`

Find the converted files in the "export" folder.
