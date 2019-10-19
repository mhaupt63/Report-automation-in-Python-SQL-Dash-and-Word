# Report-automation-in-Python-SQL-DASh-and-Word
Here is a summary of a private project I did for my employer Commonwealth Bank.

I wrote a Python program that automates the creation of a printable report in Microsoft Word, and also an online interactive version in Plotly/Dash.  The report relates to technology risk Risk Appetite metrics.

The program extracts data from an SQL database, manipulates them in Python and outputs a couple of pickled dataframes.  These dataframes are then used to preare a report, including tables and graphs, in Mictorsoft Word.  The same dataframes are also used to drive an interactive online dashboard version of the report.

The program consists of 4 modules:
1. RAS_extract.py., which joins and extracts the required SQL data into a number of Python dataframes.  Uses moldule 'pyodbc.'
2. Read_RAS_mapping.py, a tiny module that reads in an Excel file with the report structure and certain reference data
3. RAS_TL.py, which transforms the data and summarises it into a dataframe ConsolidatedDf
4. RAS_print_report.py, which outputs the data from ConsolidatedDf into a Word report (docx), including trend graphs (Matplotlib)
5. app.py, which outputs the data in an interactive online dashboard version of the report.  This is done via Dash.  The module app.py calls a number of other modules:  layout.py, callbacks.py, tablelinegraphs.py (using Plotly subplots to show a complicated table with embedded time series graphs)
