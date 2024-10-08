# getDSN
cdw-getDSN to list an ETL servers ODBC DSN entries

The script may have a few bugs that still require working out and I think the best way to use this script is to have it run as a scheduled daily task where it will store the files/over-write the file if no changes detected. I also discovered if you execute the script with an administrative account it will run in any directory the script is so make sure to either place the script within it's own directory, ie, 'C:\###_coop\ETL\ODBC_DSN\' and create a sub-directory for storing the report(s) ('C:\###_coop\ETL\ODBC_DSN\Report').

If a change is detected it will save the before/after files with a timestamp and then you can provide the information to the appropriate team for verification/resolution.
