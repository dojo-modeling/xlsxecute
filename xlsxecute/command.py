import argparse
import csv
import formulas
import json
import openpyxl
import os
import time

DESCRIPTION = """
    This tool will take an Excel model (.xlsx), update any parameters as defined via command line
    arguments or in the parameters file and calculate all cells, resulting in an Excel spreadsheet
    resembling the original, but with all formula cells replaced by the calculated values.

    The parameter file can be either JSON file or a CSV file in the following format:

    JSON:
    >  {
    >     "Sheet name.Cell1": "Replacement value string",
    >     "Sheet name.Cell2": Replacement value float
    >  }

    Example: params.json
    >  {
    >      "Variables.C2": "red",
    >      "Variables.C3": 0.8
    >  }


    CSV:
    >  Sheet name.Cell1,Replacement value string
    >  Sheet name.Cell2,Replacement value float

    Example: params.csv
    >  Variables.C2,red
    >  Variables.C3,0.8

    NOTE: Do NOT include a header row in the CSV
    """

def render(source_file, run_dir, output_dir, parameter_file=None, arg_params=None):

    source_file = source_file
    param_file = parameter_file
    base_run_dir = run_dir
    output_dir = output_dir
    run = str(int(time.time()*1000))
    run_dir = os.path.join(base_run_dir, run)
    run_file = os.path.join(run_dir, source_file)

    if not os.path.exists(run_dir):
        os.makedirs(run_dir, exist_ok=True)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    print("Parsing parameters")
    if param_file:
        if param_file.lower().endswith(".csv"):
            param_data = csv.reader(open(param_file))
            params = dict(row[0:2] for row in list(param_data))
        elif param_file.lower().endswith(".json"):
            params = json.load(open(param_file))
        else:
            raise RuntimeError(
                f'Parameter file "{param_file}" is not an accepted format. Only JSON and CSV files are allowed. Check command help for more information.'
            )
    else:
        params = {}

    for cli_param in arg_params:
        try:
            cell, value = cli_param.rsplit('=', 1)
        except ValueError as err:
            raise RuntimeError(
                f'Parameters must be in the form of "{{sheet}}.{{cell}}:{{value}}", got "{cli_param}".'
            )
        params[cell] = value

    # Update original spreadsheet, replacing values in cells based on the parameters
    print("Updating cells based on parameters")
    original_workbook = openpyxl.load_workbook(source_file)
    for location, value, *_ in params.items():
        sheet, cell = location.split(".")
        original_workbook[sheet][cell].value = value
    original_workbook.save(run_file)

    # Load the updated spreadsheet, calculate all formulas and save result to output directory
    print("Procssing formulas to resolve model")
    xl_model = formulas.ExcelModel().load(run_file).finish()
    solution = xl_model.calculate()
    xl_model.write(dirpath=output_dir)

    print("Output file(s):")
    for output_file in os.listdir(output_dir):
        os.rename(os.path.join(output_dir, output_file), os.path.join(output_dir, output_file.lower()))
        print(f"  {os.path.join(output_dir, output_file.lower())}")


def main():
    arg_parser = argparse.ArgumentParser(
        description=DESCRIPTION,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    arg_parser.add_argument(
        "source_file",
        type=str,
        help="Excel (xlsx) file that contains the model",
    )
    arg_parser.add_argument(
        "parameter_file",
        type=str,
        help="Path to json or csv parameter file",
        nargs="?",
    )
    arg_parser.add_argument(
        "--output_dir",
        type=str,
        help="Optional output location. (Default: output)",
        default="output",
    )
    arg_parser.add_argument(
        "--run_dir",
        type=str,
        help="Optional directory to store intermediate files. (Default: runs)",
        default="runs",
    )
    arg_parser.add_argument(
        "--param", "-p",
        action="append",
        type=str,
        help="",
        dest="params",
        metavar="{sheet}.{cell}={new_value}",
        default=[],
    )

    args = arg_parser.parse_args()

    render(
        source_file=args.source_file,
        parameter_file=args.parameter_file,
        arg_params=args.params,
        output_dir=args.output_dir,
        run_dir=args.run_dir,
    )


if __name__ == "__main__":
    main()

