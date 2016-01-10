"""
This script is used to pull down the most recent stock price for the following
tickers and update the spreadsheet who's URL gets passed in as a command line
argument
"""

import sys
import requests
import json
import gspread
import time
from pprint import pprint
from os.path import dirname, realpath
from optparse import OptionParser
from datetime import datetime
from oauth2client.client import SignedJwtAssertionCredentials
from urllib.parse import quote_plus

BASE_DIR = dirname(realpath(__file__))


def get_ticker_symbols(ws):
    """
    Gets the ticker symbols from the column holding them
    :param ws: The worksheet object to get the values from
    :type ws: gspread.models.Worksheet
    :return: a list of the ticker symbols from the column
    :rtype : list
    """
    tickers = list(set(ws.col_values(ticker_column)))

    if 'Company' in tickers:
        tickers.remove('Company')

    if '' in tickers:
        tickers.remove('')

    return tickers


def generate_oauth_credentials():
    """
    Returns OAuth2 credentials from credentials.json
    :rtype : SignedJwtAssertionCredentials
    :return: OAuth2 Credentials to authenticate with Google API
    """
    json_key = json.load(open(BASE_DIR + '/credentials.json'))
    scope = ['https://spreadsheets.google.com/feeds']

    return SignedJwtAssertionCredentials(json_key['client_email'],
                                         bytes(json_key['private_key'], 'utf-8'),
                                         scope
                                         )


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


def get_price_data(query_url):
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


PARSER = OptionParser()

PARSER.add_option("-t", "--ticker_column", action="store", type="int", dest="ticker_column",
                  help="The column number holding the ticker symbols (A = 1 and so on)")
PARSER.add_option("-p", "--price_update_column", action="store", type="int",
                  dest="price_update_column",
                  help="The column number to update with the prices (A = 1 and so on)")
PARSER.add_option("-c", "--change_update_column", action="store", type="int",
                  dest="change_update_column",
                  help="The column number to place the current change in stock price")
(options, args) = PARSER.parse_args()

ticker_column = 3 if not options.ticker_column else options.ticker_column
price_update_column = 6 if not options.price_update_column else options.price_update_column
change_update_column = 13 if not options.change_update_column else options.change_update_column

# Get the URL for the spreadsheet to update from command line argument
spreadsheet_key = sys.argv[1]

# Authenticate with the Google API using OAuth2
gc = gspread.authorize(generate_oauth_credentials())
worksheet = gc.open_by_key(spreadsheet_key).sheet1
ticker_symbols = get_ticker_symbols(worksheet)

# Get pricing data from Yahoo Finance
price_data, daily_returns = get_price_data(build_yql_query(ticker_symbols))

# Update cells in the Current Price Column with the pricing info from the ticker
# in the Company column
for cell_row in range(2, 11):
    stock_price = price_data[worksheet.cell(cell_row, ticker_column).value]
    worksheet.update_cell(cell_row, price_update_column, stock_price)

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
