# Capture_SplitReport
Data Processing and Excel Generation Script

This Python script processes data from multiple text files and generates Excel files for each client, as well as a summary Excel file.

Prerequisites
Python 3.x
Pandas library (install via pip install pandas)
XlsxWriter library (install via pip install XlsxWriter)
Usage
Ensure that all necessary text files (AL_GE_20180203.txt, EOB_GE_20180203.txt, MonthlyReport_GE_20180203.txt, Accum_GE_20180203.txt) are present in the same directory as the script.
Run the script.
Description
The script reads data from the text files and processes them into Pandas DataFrame objects.
It generates individual Excel files for each client, containing data from the different categories (AL, EOB, Order, Accum).
Additionally, it creates a summary Excel file (totals.xlsx) containing combined data from all clients.
Important Note
Ensure that the text files contain the required data in the specified format for accurate processing.

Disclaimer
This script is provided as-is and without warranty. Use it responsibly and at your own risk.
