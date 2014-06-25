access-db
=========

Analysis tool

##Pre-requisites

Before you run the script on your computer, go through the checklist below and ensure you have the pre-requisites.

1. **python 2.7.7** (python 2.7.7 is the latest release of Python 2 as of June 25, 2014)
  
  For windows, click [here](https://www.python.org/downloads/windows/ "Python 2.7.7 Installation") to install. Note that to be able to run python commands on the windows command line, you'd have to add the installation directory *Python27* to the PATH variable. Check [this](http://stackoverflow.com/questions/4621255/how-do-i-run-a-python-program-in-the-command-prompt-in-windows-7 "Stackoverflow thread") thread for help.

2. **pyodbc 3.0.7** (pyodbc 3.0.7 is the lastest release as of June 25, 2014)
  
  To connect to Access DB, the script uses this external library. To install, click [here](https://code.google.com/p/pyodbc/downloads/list). Note that if you're working on a LAM-issued computer, installing the 32-bit version is recommended.

3. **openpyxl 2.0.3** (openpyxl 2.0.3 is the lastest release as of June 25, 2014)
  
  To process Excel 2007 xlsx/xlsm files, the script uses this external library. To install, click [here](https://pypi.python.org/pypi/openpyxl).
  Caution: the documentation for openpyxl 2.0.3 is still being updated, so don't fret if the code in the tutorial section of the library website doesn't work as it is supposed to (you might need to go to `\Python27\lib\site-packages\openpyxl` and check the invidual files and figure out the correct syntax on your own).

##Running the Script

Open the command line, and type `python main.py`.

##Example

```python
> python main.py
1 Software
2 Controls...1
Path...C:\Users\LAMy\Desktop\GitCode\analysis-tool\PR-Metrics.accdb
Workbook name ('something.xlsx')...test.xlsx
Previous worksheet name ('Sheet1')...13Jun
New worksheet name ('Sheet2')...24Jun
```
