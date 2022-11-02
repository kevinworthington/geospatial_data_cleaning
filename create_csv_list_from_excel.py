
"""
 The following takes a list of columns and creates rows for each point
 the output is unique id (row/name plus column name), lat, lng
-- this is a python program that is best run through a virtual environment as you'll need to install some libraries
-- do this open the terminal and navigate to your project folder
-- call ```python3 -m venv venv``` from the terminal
-- activate the new venv, run ```source venv/bin/activate```
-- then install the requirements, run ```pip install -r requirements.txt```
-- (Aside) conversely to export use - pip freeze > requirements.txt

-- Now you can run run the program

python create_csv_list_from_excel.py -p "output.xlsx" -s Sheet1 -c "PO1 GPS coordinates,PO2 GPS coordinates,PO3 GPS coordinates,PO4 GPS coordinates,PO5 GPS coordinates" -i "Plot code" -o output.csv
"""
import argparse
import pandas as pd
import re
import json


def main(path,sheet,cols,_index,output, verbose=0):

    # step 1 - open the excel file

    df = pd.read_excel(io=path, sheet_name=sheet)

    # step 2 create a list of columns
    cols=cols.split(",")

    # step 3 create stub geojson file
    output_df = pd.DataFrame(columns = ['id', 'lat', 'lng'])

    for index, row in df.iterrows():
        for c in cols:
            # split the location
            points= df.at[index, c].split(",")

            output_df = output_df.append({'id': row[_index]+"_"+c, 'lat': float(points[0]), 'lng': float(points[1])}, ignore_index=True)


    output_df.to_csv(output, index=False)


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("-v", "--verbose", help="increase output verbosity",
                        action="count")

    parser.add_argument("-p", "--path", help="excel file path", )
    parser.add_argument("-s", "--sheet", help="excel sheet", )
    parser.add_argument("-c", "--cols", help="The columns with points needing normalization", )
    parser.add_argument("-i", "--index", help="The column with the index", )
    parser.add_argument("-o", "--output", help="output file", )

    return parser.parse_args()

if __name__ == '__main__':
    args = parse_args()
    if not args.verbose:
        args.verbose = 0


    main(args.path,args.sheet,args.cols,args.index,args.output,args.verbose)
