# Permit-Processor
This script automates the processing of PDF files containing work permit and residence form data. It is designed to work in combination with an Excel file (controlled via xlwings) and output the parsed information into a text file and an Excel table.

Features

Parses PDF filenames to extract:

Tag ID (unique identifier)

Name

Date (formatted and parsed)

Form number

Price (if present)

Avoids reprocessing of previously processed files using import date tracking

Maps form numbers to actual services and prices using an Excel sheet ("Services_Table")

Writes new entries to:

A text log file (Processed_Permits.txt)

An Excel table named master_list on Sheet 6

Folder Structure

Expected folder structure for files:

<WORK_PERMIT_PATH>/2025/Abril/*.pdf
<RESIDENCE_PATH>/2025/Abril/*.pdf

Excel Sheet Layout

Sheet 5 contains a table named Services_Table with the following columns:

i-number (form number, e.g., i-765)

Service name

Price

Sheet 6 contains a table named master_list where all processed data is appended

Setup Instructions

Install xlwings: pip install xlwings

Modify the paths at the top of the script (<YOUR_WORK_PERMIT_PATH>, etc.) to match your environment.

Open the Excel file and ensure:

Sheet 5 has the Services_Table properly defined

Sheet 6 has the master_list table created

Create a button in Excel that triggers this script via xlwings.

Filename Convention

The script expects filenames in this format:

i-765 John Doe 042524 120

Where:

i-765 is the form number

John Doe is the name

042524 is the date (MMDDYY)

120 is the price (optional)

Output

The script generates:

Processed_Permits.txt log file with CSV data

Appended data to the Excel master_list table

Notes

Make sure the script runs only once per day or manages deduplication effectively.

Locale is set to es_MX.UTF-8 to support Spanish month names.

Update your system locale if errors arise from date parsing or month formatting.
