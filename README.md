# xlsxecute

This tool will take an Excel model (.xlsx), update any parameters as defined via command line arguments or in the parameters file and calculate all cells, resulting in an Excel spreadsheet resembling the original, but with all formula cells replaced by the calculated values.

Parameters that define how to update cells in your spreadsheet can be provided in three ways: 
JSON file, CSV file, or command line arguments

If both a config file and command line arguments are provided, the command line arguments 


### Config file formatting

Only one config file can be provided at a time. The config file can either be 

#### Command line arguments:

Command line arguments take the form of:
```bash
-f "Sheet name.Cell1=Replacement value string" -f "Sheet name.Cell2=Replacement_value_float"
```
Note: Quotation marks are not required if there no space in the parameter string.

Example:
```
xlsxecute -f "Variables.C2=red" -f Variables.C3=0.8 sample.xlsx
```

#### JSON:
```json
{
   "Sheet name.Cell1": "Replacement value string",
   "Sheet name.Cell2": Replacement_value_float
}
```

Example: params.json
```json
{
    "Variables.C2": "red",
    "Variables.C3": 0.8
}
```


#### CSV:
```csv
Sheet name.Cell1,Replacement value string
Sheet name.Cell2,Replacement value float
```

Example: params.csv
```csv
Variables.C2,red
Variables.C3,0.8
```

NOTE: Do NOT include a header row in the CSV

<br/>
<br/>


#### Executable usage:

```
usage: xlsxecute [-h] [--output_dir OUTPUT_DIR] [--run_dir RUN_DIR] [--param {sheet}.{cell}={new_value}] source_file [parameter_file]

positional arguments:
  source_file           Excel (xlsx) file that contains the model
  parameter_file        Path to json or csv parameter file

optional arguments:
  -h, --help            show this help message and exit
  --output_dir OUTPUT_DIR
                        Optional output location. (Default: output)
  --run_dir RUN_DIR     Optional directory to store intermediate files. (Default: runs)
  --param {sheet}.{cell}={new_value}, -p {sheet}.{cell}={new_value}
```
