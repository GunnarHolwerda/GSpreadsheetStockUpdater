"""
This script it used to copy the end portfolio value to another spreadsheet
so that it can be analyzed
"""

import json
import gspread
import time
import argparse
from pprint import pprint
from os.path import dirname, realpath
from datetime import datetime
from oauth2client.client import SignedJwtAssertionCredentials

BASE_DIR = dirname(realpath(__file__))


def generate_oauth_credentials():
    """
    Returns OAuth2 credentials from credentials.json
    :rtype : SignedJwtAssertionCredentials
    :return: OAuth2 Credentials to authenticate with Google API
    """
    json_key = json.load(open(BASE_DIR + '/credentials.json'))
    scope = ['https://spreadsheets.google.com/feeds']

    return SignedJwtAssertionCredentials(json_key['client_email'], bytes(json_key['private_key'], 'utf-8'), scope)

parser = argparse.ArgumentParser()
parser.add_argument(
    "spreadsheet_key", help="The spreadsheet key of the Google Spreadsheet to update")
parser.add_argument(
    "value_col", help="The column in which the spreadsheet value will be copied too", type=int)
parser.add_argument(
    "date_col", help="The column in which the date of copy will be copied too", type=int)
parser.add_argument("value_cell", help="Cell of the value to copy (i.e. G6)")
parser.add_argument("-p", "--portfolio_worksheet",
                    help="The spreadsheet title which to copy the current value to")
parser.add_argument("-v", "--value_worksheet",
                    help="The spreadsheet title which to copy the portfolio lives")
args = parser.parse_args()

# Authenticate with the Google API using OAuth2
gc = gspread.authorize(generate_oauth_credentials())
spreadsheet = gc.open_by_key(args.spreadsheet_key)

# If a spreadsheet title was specified load by title, else default to the
# second worksheet
if args.portfolio_worksheet:
    value_worksheet = spreadsheet.worksheet(args.spreadsheet_title)
else:
    value_worksheet = spreadsheet.get_worksheet(1)

if args.value_worksheet:
    portfolio_worksheet = spreadsheet.worksheet(args.value_worksheet)
else:
    portfolio_worksheet = spreadsheet.get_worksheet(0)

# Get the value to copy from the portfolio spreadsheet
end_of_day_value = portfolio_worksheet.acell(args.value_cell).value

# Get all of the values in the values column
values = value_worksheet.col_values(args.value_col)
# Remove all empty values from the list
values = [i for i in values if i != '']

# The row to update is the length of the values array + 1
row_of_cell_to_update = len(values) + 1
pprint(row_of_cell_to_update)
cur_date = datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d')

# Update value and date cell
value_worksheet.update_cell(row_of_cell_to_update,
                            args.value_col, end_of_day_value)
value_worksheet.update_cell(row_of_cell_to_update, args.date_col, cur_date)
