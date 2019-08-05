# Report-automation-in-Python-SQL-and-Word
Here is a summary of a private project I did for my employer Commonwealth Bank.

I wrote a Python program that automates the creation of a report in Microsoft Word.  The report relates to technology risk Risk Appetite metrics.

The program extracts data from an SQL database, manipulates them in Python and outputs a report, including tables and graphs, in Mictorsoft Word.  It was decided to use Word rather than Tableau to avoid working with pdf output which is difficult to edit and incorporate with Microsoft Office.

The program consists of 4 modules:
1. RAS_extract.py., whicj joins and extracts the required SQL data into a number of Python dataframes.  Uses moldule 'pyodbc.'
2. Read_RAS_mapping.py, a tiny module that reads in an Excel file with the report structure and certain reference data
3. RAS_TL.py, which rtransforms the data and summarises it into a dataframe ConsolidatedDf
4. RAS_plot.py, which outputs the data from ConsolidatedDf into a Word report (docx), including trend graphs (Matplotlib)
