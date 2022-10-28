
import argparse
import pandas as pd
import re
"""
The following looks through an excel file, takes the columns specified and finds the lat and lng equivelent

-- this is a python program that is best run through a virtual environment as you'll need to install some libraries
-- do this open the terminal and navigate to your project folder
-- call ```python3 -m venv venv``` from the terminal
-- activate the new venv, run ```source venv/bin/activate```
-- then install the requirements, run ```pip install -r requirements.txt```
-- (Aside) conversely to export use - pip freeze > requirements.txt

-- Now you can run run the program 

python normalize_points.py -p "Plot location 1PO-13PO Intern Challenge.xlsx" -s Sheet2 -c "PO1 GPS coordinates,PO2 GPS coordinates,PO3 GPS coordinates,PO4 GPS coordinates,PO5 GPS coordinates" -o output.xlsx
"""

def main(path,sheet,cols,output, verbose=0):

    # step 1 - open the locations

    df = pd.read_excel(io=path, sheet_name=sheet)
    # print(df.head(5))

    # step 2 create a list of columns
    cols=cols.split(",")

    for index, row in df.iterrows():
        for c in cols:

            n = re.findall('[0-9\.]+', df.at[index, c])

            for idx, i in enumerate(n):
                if i.count(".") > 1:
                    # note is there are two decimals, take the first decimal and split to two numbers
                    print("more than 2 decimals", i)
                    n[idx] = i.split('.', 1)
                    print("now two numbers",   n[idx])

            #flatten list
            flat_list=[]
            for i in n:
                if type(i) ==list:
                    for j in i:
                        flat_list.append(j)
                else:
                    flat_list.append(i)
            n = flat_list
            print("new list",n)
            #convert all strings to numbers
            n = list(map(float, n))


            if (len(n)==6):

                df.at[index, c] = str(n[0] + ((( n[1] * 60) + ( n[2])) / 3600))+","+str(n[3] + ((( n[4] * 60) + ( n[5])) / 3600))
            else:
                # there should be 4
                if(type(n[1])==float):
                    df.at[index, c] = str(n[0] + (n[1] / 60)) + "," + str(n[2] + (n[3] / 60))

    # export to excel
    df.to_excel(output, index=False)
    print("done")



def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("-v", "--verbose", help="increase output verbosity",
                        action="count")

    parser.add_argument("-p", "--path", help="excel file path", )
    parser.add_argument("-s", "--sheet", help="excel sheet", )
    parser.add_argument("-c", "--cols", help="The columns with points needing normalization", )
    parser.add_argument("-o", "--output", help="output file", )

    return parser.parse_args()

if __name__ == '__main__':
    args = parse_args()
    if not args.verbose:
        args.verbose = 0


    main(args.path,args.sheet,args.cols,args.output,args.verbose)
