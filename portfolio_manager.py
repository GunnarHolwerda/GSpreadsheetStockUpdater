"""
This script is used to pull down the most recent stock price for the following
tickers and update the spreadsheet who's URL gets passed in as a command line
argument
"""

import argparse
import config
import gspread
import json
import requests
import smtplib
import time
from utility_functions import construct_time_variables, remove_dollar_sign_and_commas
from pprint import pprint
from oauth2client.client import SignedJwtAssertionCredentials
from os.path import dirname, realpath
from datetime import datetime
from urllib.parse import quote_plus

BASE_DIR = dirname(realpath(__file__))


def generate_oauth_credentials():
    """
    Returns OAuth2 credentials from credentials.json
    :rtype : SignedJwtAssertionCredentials
    :return: OAuth2 Credentials to authenticate with Google API
    """
    json_key = json.load(open(BASE_DIR + '/credentials.json'))
    scope = ['https://spreadsheets.google.com/feeds']

    return SignedJwtAssertionCredentials(json_key['client_email'],
                                         bytes(
                                             json_key['private_key'], 'utf-8'),
                                         scope)


def get_cell_range(worksheet, start_col, end_col, start_row, end_row):
    alphanum_range = "" + str(chr(start_col + 96)) + str(start_row) + \
        ":" + str(chr(end_col + 96)) + str(end_row)
    return worksheet.range(alphanum_range)


def get_biggest_movers(ws, ticker_col, change_col):
    tickers = ws.col_values(ticker_col)
    tickers = [i for i in tickers if i != ''][1:]
    changes = ws.col_values(change_col)
    changes = [i for i in changes if i != ''][1:]

    changes_dict = {}
    pos = 0
    for ticker in tickers:
        changes_dict[ticker] = changes[pos]
        pos += 1

    ret_str = "Performance by Stock\n"
    for ticker, change in changes_dict.items():
        ret_str += "{}: {}\n".format(ticker, change.rjust(6))

    return ret_str


def email_end_of_day_report(to_address, from_address, prev_total, cur_total, custom_value, biggest_movers):
    day_of_week, str_month, str_day = \
        construct_time_variables(datetime.today())
    prev_total_float = remove_dollar_sign_and_commas(prev_total)
    cur_total_float = remove_dollar_sign_and_commas(cur_total)
    daily_change = (float(cur_total_float) -
                    float(prev_total_float)) / float(prev_total_float)

    msg = ("Subject: Daily Stock Report For {}, {} {}\n"
           "Daily Stock Report\n\n"
           "Yesterday's ending value: {}\n"
           "Todays's ending value: {}\n\n"
           "Overall increase/decrease: {}%\n"
           "{}: {}\n\n"
           "{}"
           "Have a great day!").format(day_of_week,
                                       str_month,
                                       str_day,
                                       prev_total,
                                       cur_total,
                                       round(100 * daily_change, 2),
                                       config.custom_value_name,
                                       custom_value,
                                       biggest_movers)
    # Send the message via our own SMTP server
    s = smtplib.SMTP('smtp.gmail.com:587')
    s.starttls()
    s.login(config.user, config.password)
    s.sendmail(from_address, to_address, msg)
    s.quit()


def get_ticker_cells(worksheet):
    """
    Gets the ticker symbols from the column holding them
    :param ws: The worksheet object to get the values from
    :type ws: gspread.models.Worksheet
    :return: a list of the ticker symbols from the column
    :rtype : list
    """
    tickers = worksheet.range(config.ticker_range)
    return tickers

def get_ticker_symbols(ticker_cells):
    symbols = []
    for cell in ticker_cells:
        symbols.append(cell.value)

    return symbols


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
    pprint(price_info)

    for index in range(0, len(price_info)):
        price_dict[ticker_symbols[index]] = price_info[index]['LastTradePriceOnly']
        daily_return_dict[ticker_symbols[index]] = price_info[index]['Change']

    return price_dict, daily_return_dict


def get_yesterdays_total(worksheet):
    cell = worksheet.find("Yesterday's Total:")
    yesterdays_total = worksheet.cell(cell.row, cell.col + 1)

    return yesterdays_total.value


def store_end_of_day_value(ss, value_worksheet="Portfolio Value over Time", porfolio_worksheet="Current Portfolio Value"):
    """
    Copies the value of the portfolio at the end of the day and copies it to a
    second spreadsheet to be stored along side the date.
    :param ss: the gspread spreadsheet object
    :type ss: gspread.Spreadsheet
    """
    # If a spreadsheet title was specified load by title, else default to the
    # second worksheet
    portfolio_worksheet = ss.worksheet(porfolio_worksheet)
    value_worksheet = ss.worksheet(value_worksheet)

    # Get the value to copy from the portfolio spreadsheet
    todays_total = portfolio_worksheet.acell(config.copy_cell).value
    yesterdays_total = get_yesterdays_total(portfolio_worksheet)
    custom_value = portfolio_worksheet.acell(config.custom_value).value

    # Get all of the values in the values column
    values = value_worksheet.col_values(config.save_column)
    # Remove all empty values from the list
    values = [i for i in values if i != '']

    biggest_movers = get_biggest_movers(portfolio_worksheet, 11, 14)

    email_end_of_day_report(config.to_addr,
                            config.from_addr,
                            yesterdays_total,
                            todays_total,
                            custom_value, biggest_movers)

    # The row to update is the length of the values array + 1
    row_of_cell_to_update = len(values) + 1
    cur_date = datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d')

    # Update value and date cell
    value_worksheet.update_cell(row_of_cell_to_update,
                                config.save_column, todays_total)
    value_worksheet.update_cell(
        row_of_cell_to_update, config.date_column, cur_date)


def update_portfolio_value(ss):
    """
    Places the current price for each stock in the config.ticker_column
    in the config.update_column. Also, stores the daily net_change in the config.change_update_column
    param ss: the gspread spreadsheet object
    :type ss: gspread.Spreadsheet
    """
    # Authenticate with the Google API using OAuth2
    cells = []
    worksheet = ss.sheet1
    ticker_cells = get_ticker_cells(worksheet)
    ticker_symbols = list(set(get_ticker_symbols(ticker_cells)))

    # Get pricing data from Yahoo Finance
    price_data, daily_returns = get_price_data(
        build_yql_query(ticker_symbols), ticker_symbols)

    cur_price_cells = get_cell_range(
        worksheet, config.price_update_column, config.price_update_column, 2, 10)

    # Update cells in the Current Price Column with the pricing info from the ticker
    # in the Company column
    for i in range(0, len(ticker_cells)):

        ticker = ticker_cells[i].value
        print("Updating price of {}".format(ticker))
        stock_price = price_data[ticker]
        cur_price_cells[i].value = stock_price
        #worksheet.update_cell(cell_row, config.price_update_column, stock_price)

    worksheet.update_cells(cur_price_cells)

    # Update cells in the Change with the daily change in price info from the ticker
    # in the Company column
    for cell_row in range(2, 11):
        change = daily_returns[worksheet.cell(
            cell_row, config.ticker_column).value]
        worksheet.update_cell(cell_row, config.net_change_update_column, change)

    # Update the cell next to "Last updated: to the current timestamp
    cell = worksheet.find("Last updated:")
    cur_time = datetime.fromtimestamp(
        time.time()).strftime('%Y-%m-%d %H:%M:%S')
    worksheet.update_cell(cell.row, cell.col + 1, cur_time)

    worksheet.update_cells(cells)

# ticker_column = 3 if not args.ticker_column else args.ticker_column
# price_update_column = 6 if not args.update_column else args.update_column
# change_update_column = 13 if not args.change_update_column else args.change_update_column

parser = argparse.ArgumentParser()
parser.add_argument("-p", "--portfolio_worksheet", action="store",
                    help="The spreadsheet title which to copy the current value to")
parser.add_argument("-s", "--save_value", action="store_true",
                    help="Copy end of the day value to another location")
parser.add_argument("-v", "--value_worksheet", action="store", type=int,
                    help="The spreadsheet title which to copy the portfolio lives")
args = parser.parse_args()

# Get the URL for the spreadsheet to update from command line argument
spreadsheet_key = config.spreadsheet_key

# Authenticate and get spreadsheet object
gc = gspread.authorize(generate_oauth_credentials())
spreadsheet = gc.open_by_key(spreadsheet_key)

if args.save_value:
    # TODO: Figure out how to accomodate if worksheet titles are provided
    store_end_of_day_value(spreadsheet)
else:
    update_portfolio_value(spreadsheet)
