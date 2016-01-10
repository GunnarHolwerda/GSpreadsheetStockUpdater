"""
This script is used to pull down the most recent stock price for the following
tickers and update the spreadsheet who's URL gets passed in as a command line
argument
"""

import authentication
import argparse
import requests
import json
import gspread
import time
from pprint import pprint
from os.path import dirname, realpath
from datetime import datetime
from urllib.parse import quote_plus

BASE_DIR = dirname(realpath(__file__))


def get_ticker_symbols(worksheet):
    """
    Gets the ticker symbols from the column holding them
    :param ws: The worksheet object to get the values from
    :type ws: gspread.models.Worksheet
    :return: a list of the ticker symbols from the column
    :rtype : list
    """
    tickers = list(set(worksheet.col_values(ticker_column)))

    if 'Company' in tickers:
        tickers.remove('Company')

    if '' in tickers:
        tickers.remove('')

    return tickers


def build_yql_query(tickers):
    """
    Builds the url to make the curl request to the Yahoo Finance URL
    :param tickers: A list of the ticker symbols of the stocks you want to update
    :type tickers: list
    :return: Encoded URL to be used in a GET request
    """
    yql_base_url = "http://query.yahooapis.com/v1/public/yql"
    end_yql_url = ("&format=json&diagnostics=true&env=store%3A%2F%2Fdatatables"
                   ".org%2Falltableswithkeys&callback=")
    yql_query = "select LastTradePriceOnly, Change from yahoo.finance.quotes where symbol in ("

    for ticker in tickers:
        yql_query += "'{0}',".format(ticker)

    # Take off trailing comma and add the end parenthesis
    yql_query = yql_query.rstrip(',')
    yql_query += ")"
    yql_query = quote_plus(yql_query)

    return yql_base_url + "?q=" + yql_query + end_yql_url


def get_price_data(query_url, ticker_symbols):
    """
    Returns the pricing info for the stocks in TICKER_SYMBOLS in dictionary format
    :param query_url: A url encoded string to be used in a requests.get call
    :type query_url: str
    :return: A dictionary of ticker to prices
    :rtype : dict
    """
    # Make curl request to Yahoo Finance URL
    request = requests.get(query_url)
    response = json.loads(request.text)
    price_info = response['query']['results']['quote']
    price_dict = {}
    daily_return_dict = {}

    for index in range(0, len(price_info)):
        price_dict[ticker_symbols[index]] = price_info[index]['LastTradePriceOnly']
        daily_return_dict[ticker_symbols[index]] = price_info[index]['Change']

    return price_dict, daily_return_dict


def store_end_of_day_value(ss, value_cell, value_col, date_col, value_worksheet=1, porfolio_worksheet=0):
    # If a spreadsheet title was specified load by title, else default to the
    # second worksheet
    portfolio_worksheet = ss.worksheet(porfolio_worksheet)
    value_worksheet = ss.worksheet(value_worksheet)

    # Get the value to copy from the portfolio spreadsheet
    end_of_day_value = portfolio_worksheet.acell(value_cell).value

    # Get all of the values in the values column
    values = value_worksheet.col_values(value_col)
    # Remove all empty values from the list
    values = [i for i in values if i != '']

    # The row to update is the length of the values array + 1
    row_of_cell_to_update = len(values) + 1
    pprint(row_of_cell_to_update)
    cur_date = datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d')

    # Update value and date cell
    value_worksheet.update_cell(row_of_cell_to_update,
                                value_col, end_of_day_value)
    value_worksheet.update_cell(row_of_cell_to_update, date_col, cur_date)

def update_portfolio_value(ss, update_column, change_update_column):
    # Authenticate with the Google API using OAuth2
    worksheet = ss.sheet1
    ticker_symbols = get_ticker_symbols(worksheet)

    update_column = args.update_column
    change_update_column = args.change_update_column

    # Get pricing data from Yahoo Finance
    price_data, daily_returns = get_price_data(build_yql_query(ticker_symbols), ticker_symbols)

    # Update cells in the Current Price Column with the pricing info from the ticker
    # in the Company column
    for cell_row in range(2, 11):
        stock_price = price_data[worksheet.cell(cell_row, ticker_column).value]
        worksheet.update_cell(cell_row, update_column, stock_price)

    # Update cells in the Change with the daily change in price info from the ticker
    # in the Company column
    for cell_row in range(2, 11):
        pprint(daily_returns)
        change = daily_returns[worksheet.cell(cell_row, ticker_column).value]
        print(change)
        worksheet.update_cell(cell_row, change_update_column, change)

    # Update the cell next to "Last updated: to the current timestamp
    cell = worksheet.find("Last updated:")
    cur_time = datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S')
    worksheet.update_cell(cell.row, cell.col + 1, cur_time)

ticker_column = 3 if not args.ticker_column else args.ticker_column
price_update_column = 6 if not args.update_column else args.update_column
change_update_column = 13 if not args.change_update_column else args.change_update_column



parser = argparse.ArgumentParser()
parser.add_argument(
    "spreadsheet_key", help="The spreadsheet key of the Google Spreadsheet to update")
parser.add_argument("-c", "--change_update_column", action="store", type="int",
                    help="The column number to place the current change in stock price")
parser.add_argument("-d",
                    "--date_column", help="The column in which the date of copy will be copied too", type=int)
parser.add_argument("-p", "--portfolio_worksheet", action="store",
                    help="The spreadsheet title which to copy the current value to")
parser.add_argument("-s", "--save_value", action="store_true",
                    help="Copy end of the day value to another location")
parser.add_argument("-t", "--ticker_column", action="store", type=int,
                    help="The column number holding the ticker symbols (A = 1 and so on)")
parser.add_argument("-u", "--update_column", action="store", type=int,
                    help="The column number to update with the prices (A = 1 and so on)")
parser.add_argument("-v", "--value_worksheet", action="store", type=int,
                    help="The spreadsheet title which to copy the portfolio lives")
parser.add_argument("-x", "--copy_column", action="store", type=int,
                    help="The column in which the spreadsheet value will be copied too")
parser.add_argument("-z", "--copy_cell", action="store",
                    help="The cell which to copy to store")
args = parser.parse_args()

# Get the URL for the spreadsheet to update from command line argument
spreadsheet_key = args.spreadsheet_key

# Authenticate and get spreadsheet object
gc = gspread.authorize(authentication.generate_oauth_credentials())
spreadsheet = gc.open_by_key(spreadsheet_key)

if parser.save_value:
    if not args.copy_column or not args.date_column or not args.copy_cell:
        print("The copy_column, date_column, and copy_cell are required options when storing the end of the day value using the -s (--save_value) option.\n")
        exit()
    else:
        #TODO: Figure out how to accomdate if worksheet titles are provided
        store_end_of_day_value(spreadsheet, args.value_cell, args.value_col, args.date_col)
else:
    if not args.update_column or not args.change_update_column:
        print("The update_column and change_update_column are required options when updating the portfolio value.\n")
        exit()
    else:
        update_portfolio_value(spreadsheet, args.update_column, args.change_update_column)
