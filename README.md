# excel_reports

These are projects I created for automatic reporting at my current work (I work as an actuary in an insurance company).

Alfa_costs.ipynb contains a simple SQL query, which is stored into a pandas dataFrame. This is modified with some corrections (original dataFrame is kept unchanged so that it can be compared with the modified one). Finally, the modified version of the dataFrame is exported into an Excel template file and saved with a custom name, dependent on the month of the report.

reinsurance_report is a .py file containing methods which are called by the Reinsurance_statement.ipynb. Methods are called several times, each time with a different set of input parameters. Initially, Reinsurance_statement.ipynb was written until repeatable sections were apparent. reinsurance_report.py contains the sections that were of modular nature.

Reinsurance_statement.ipynb contains SQL queries that store results into pandas dataFrames. One of these queries contains the active portfolio of an insurance company and the reinsurance premium that was calculated for the active policies. The dataFrame is broken into four parts based on the Reinsurer; it is then further broken down into smaller pieces based on the type of risk that is being reinsured. Finally, data is exported into Excel and given proper style, specified by each Reinsurer. Since each Reinsurer has their own specifications for the report, some of the process could be universal, other parts had to be company-specific. (Disclaimer: the names of Reinsurer companies are invented.)
