from __future__ import print_function
import os
import os.path
import sys
import json
import time
import datetime
import pandas as pd
import yfinance as yf
from gooey import Gooey, GooeyParser


def mustBeDir(pathway):
    if os.path.isdir(pathway):
        return pathway
    else:
        raise TypeError(
            "What you provided is not a directory. Please enter a valid directory")


def mustBeFile(pathway):
    if os.path.isfile(pathway):
        return pathway
    else:
        raise TypeError(
            "What you provided is not a directory. Please enter a valid file")


@Gooey(program_name="Company Financial Data Collector",
       program_description="Speed up your grunt work process",
       menu=[{'name': 'Help', 'items': [{
           'type': 'AboutDialog',
           'menuTitle': 'About',
           'name': 'Financial Data Collector',
           'description': 'Accelerate your data research process so you can move on from the grunt work',
           'version': '1.2.0',
           'copyright': '2021 TK',
           'developer': 'Taekyu Kim'
       },
           {'type': 'MessageDialog',
            'menuTitle': 'How to use',
            'name': 'How to use',
            'message': "Input an excel file with company names and stock Tickers to retrieve financial datum of your choice. The first column must be titled 'Name', ane the second must be 'Code'."}
       ]}])
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Company Financial Data Collector')
    parser.add_argument('Input_File',
                        action='store',
                        default=stored_args.get('Input_File'),
                        widget='FileChooser',
                        help="Choose the Excel File to read in. Read 'how to use' for specifics.",
                        gooey_options=dict(
                            wildcard="Excel Files (*.xlsx)|*.xlsx", full_width=True
                        ),
                        type=mustBeFile
                        )
    parser.add_argument('Output_Directory',
                        action='store',
                        widget='DirChooser',
                        default=stored_args.get('Output_Directory'),
                        help="Output directory to save collected data file",
                        gooey_options=dict(
                            full_width=True
                        ),
                        type=mustBeDir
                        )
    parser.add_argument('Parameters',
                        action='store',
                        widget='Listbox',
                        help="Parameters you would like to retrieve",
                        nargs="+",
                        gooey_options=dict(
                            full_width=True
                        ),
                        choices=[
                            'fullTimeEmployees',
                            'city',
                            'state',
                            'country',
                            'website',
                            'market',
                            'financialCurrency',
                            'grossProfits',
                            'profitMargins',
                            'ebitda',
                            'ebitdaMargins',
                            'grossMargins',
                            'operatingMargins',
                            'operatingCashflow',
                            'revenueGrowth',
                            'bookValue',
                            'freeCashflow',
                            'targetMedianPrice',
                            'currentPrice',
                            'earningsGrowth',
                            'currentRatio',
                            'returnOnAssets',
                            'targetMeanPrice',
                            'returnOnEquity',
                            'totalCash',
                            'totalDebt',
                            'totalRevenue'
                        ],
                        default=[
                            'fullTimeEmployees',
                            'city',
                            'state',
                            'country',
                            'website',
                            'market',
                            'financialCurrency',
                            'grossProfits',
                            'profitMargins',
                            'ebitda',
                            'ebitdaMargins',
                            'grossMargins',
                            'operatingMargins',
                            'operatingCashflow',
                            'revenueGrowth',
                            'bookValue',
                            'freeCashflow',
                            'targetMedianPrice',
                            'currentPrice',
                            'earningsGrowth',
                            'currentRatio',
                            'returnOnAssets',
                            'targetMeanPrice',
                            'returnOnEquity',
                            'totalCash',
                            'totalDebt',
                            'totalRevenue',
                        ]
                        )
    parser.add_argument('-s',
                        '--Sheet_Name',
                        help="Enter the name of the sheet EXACTLY. If there is only one sheet, input 'none' or don't input anything.",
                        action='store',
                        gooey_options=dict(
                            full_width=True
                        ),
                        default="none"
                        )
    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args


def combine_files(src_directory):
    """ Read in source excel file and create
    data frame to retrieve target companies
    """
    if conf.Sheet_Name == 'none':
        all_Data = pd.DataFrame(pd.read_excel(src_directory))
    else:
        try:
            all_Data = pd.DataFrame(pd.read_excel(
                src_directory, sheet_name=conf.Sheet_Name))
        except RuntimeError as r:
            print(r.args)
    if 'Code' not in all_Data.keys():
        raise RuntimeError("this sheet does not contain the column'Code'")
    specific_Data = all_Data[all_Data['Code'].notnull()]

    return specific_Data


def calculateCagr(ticker):
    sheet = ticker.financials
    try:
        revRow = sheet.loc[['Total Revenue'], :]
    except (RuntimeError):
        raise RuntimeError(
            "there is no 'total revenue' information available.")
    revRow = revRow.loc[:, revRow.any()]
    CAGR = (revRow.iloc[:, 0][0]/revRow.iloc[:, -1][0])**(1/(int(
        revRow.columns[0].strftime('%Y'))-int(revRow.columns[len(
            revRow.columns)-1].strftime('%Y'))))-1
    CAGR *= 100
    return CAGR


def addCompany(companyName):
    newDat = {}
    shortDict = {}
    comp = yf.Ticker(companyName)
    for key, value in comp.info.items():
        newDat[key] = value
    for keys in conf.Parameters:
        if keys in newDat.keys():
            shortDict[keys] = newDat[keys]
        else:
            shortDict[keys] = 'N/A'
    try:
        shortDict['CAGR'] = calculateCagr(comp)
    except RuntimeError:
        print("An Error has occurred, couldn't calculate CAGR")
    return pd.DataFrame(shortDict, index=[companyName])


def save_results(nameDf, collected_data, output):
    """ save created financial data dataframe into selected folder for output
    """
    now = datetime.datetime.now()
    dateNTime = now.strftime("%Y%m%d_%H%M%S")
    collected_data = collected_data.reset_index()
    out_Data = pd.concat([nameDf, collected_data], axis=1)
    out_Data = out_Data.rename({'index': 'Code'}, axis=1)
    out_Data.set_index('Name', inplace=True)
    if 'profitMargins' in out_Data.keys():
        out_Data.profitMargins *= 100
    if 'grossMargins' in out_Data.keys():
        out_Data.grossMargins *= 100
    if 'operatingMargins' in out_Data.keys():
        out_Data.operatingMargins *= 100
    if 'ebitdaMargins' in out_Data.keys():
        out_Data.ebitdaMargins *= 100
    fileName = conf.Input_File
    fileName = os.path.splitext(os.path.basename(fileName))[0]
    outputFileDir = output + "/" + dateNTime + "_" + fileName + "_financials.xlsx"
    out_Data.to_excel(outputFileDir)


if __name__ == '__main__':
    try:
        conf = parse_args()
    except TypeError as e:
        print("Check your inputs' types and make sure it is an excel file and a folder.", e.args)
        sys.exit(1)
    except ValueError as v:
        print("You need to check your values to see if they are valid.", v.args)
    print("Reading file")
    try:
        sales_df = combine_files(conf.Input_File)
    except RuntimeError as r:
        print(r.args)
    codes = sales_df['Code']
    names = sales_df['Name']
    nameDf = pd.DataFrame(names).reset_index()
    nameDf = nameDf.loc[:, ['Name']]
    outputFile = pd.DataFrame(columns=conf.Parameters)
    print("Retrieving and saving requested data")
    for code in codes:
        try:
            outputFile = pd.concat([outputFile, addCompany(code)])
        except ConnectionError as c:
            print("server is not responding, check that you have the right network security access level to run this program.")
            break
        time.sleep(5)
    save_results(nameDf, outputFile, conf.Output_Directory)
    print("Done")
