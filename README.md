# __Depository__
Tools that I've written to help me with generic office tasks.

# __VBA__
###### How to assign a hot-key shortcut to a macro:
> Microsoft Excel ribbon > developer tab > view macros > select macro > options....

## [resize_all_columns](https://github.com/andrewdavis23/office-tools/blob/main/resize_all_columns)
- Resizes all columns in active sheet and returns the cursor to A1

## [combinations](https://github.com/andrewdavis23/office-tools/blob/main/combinations)
- Desigined for a specific sheet (you'll have to modify the cell references in VBA match your sheet)
- Returns two columns of equal length that are all combinations of two lists of arbitrary length
- Started crashing on Excel (array limit exceeded) unsolved

## [email_selection-to-table](https://github.com/andrewdavis23/office-tools/blob/main/email_selection-to-table)
- Demonstrates sending Excel data in an Outlook
- Emails a selected range

# __Python__
## [Basic ODBC Program.py](https://github.com/andrewdavis23/office-tools/blob/main/Basic%20ODBC%20Program.py)
- Run SQL and export Excel to user-selected location

## [Secondaries - File to Transactions Converter.py](https://github.com/andrewdavis23/office-tools/blob/main/Secondaries%20-%20File%20to%20Transactions%20Converter.py)
- A task specific data cleaner
- Redacted and stored for backup

## [Secondary Programs V2](https://github.com/andrewdavis23/office-tools/blob/main/secondary%20programs%20v2.py)
- Uses class structure
- Breaks out some of the functions
- Is scalable to perform ETL

## [process secondary folder](https://github.com/andrewdavis23/office-tools/blob/main/process%20secondary%20folder)
- takes the entire folder of excel files
- creates a list of classes
- classes contain all data and metadata needed for each secondary program
- saves the list as a pickle
- will need to modify to allow new drops (dump folder > process > add excel file to archive folder)

## [py-combos 3-28-2022.py](https://github.com/andrewdavis23/office-tools/blob/main/py-combos%203-28-2022.py)
- Returns combinations of two lists
- Duplicates removed
