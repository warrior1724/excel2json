# excel2json
A Python 3 app to convert .xlsx file to json that works using Pandas

## Dependencies
1. Install Pandas  `pip install pandas` 
2. Install SimpleJSON  `pip install simplejson` 

### Usage
* Make sure the data you want as json is in a Excel file with the Column name as the key value and the sheet's name is "Sheet1".
* Bring your xlsx file to the same directory as the .py file. Execute the following:
`$ excel2json.py filename.xlsx`
where filename.xlsx is the file you want to convert to json.

### Arguements
* -f Force: Will overwrite the json file, if exists
* -o : Specifies outfile (json) name
* -q : Quite (no confirmation massages on CLI)

and ofc you can always use -h or --help to get description of these 
