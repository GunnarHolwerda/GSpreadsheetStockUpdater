# GSpreadsheetStockUpdater

This repository contains scripts that create a small portfolio management system within a Google Spreadsheet.

## Setup
You will need to set up a project with Google in order to run these scripts. Follow the directions from [this link](http://gspread.readthedocs.org/en/latest/oauth2.html) to create the project and download the .json credentials file.

Rename the file to credentials.json and make sure that it is in the same directory as the two scripts.

These scripts run on Python 3.4 currently and you will need to install the following packages:
- oauth2client
- PyOpenSSL
- requests
- [gspread](https://github.com/burnash/gspread)

pip is the preferred method to install these: `pip install <package_name>`

## Running the scripts
### update_my_portfolio.py
`$ python3 update_my_portfolio.py <spreadsheet_key>`

You can get your `spreadsheet_key` from the URL of your spreadsheet:
![Spreadsheet Key Image](http://i.imgur.com/v666kdf.png)


Options for the script:
- `-t` or `--ticker_column`  
    - This option allows you to specify which column number holds the tickers for the stocks you want to get the prices of (A = 1, B = 2 and so on...)
- `-p` or `--price_update_column`  
    - This option allows you to specify which column number to update the price value in (A = 1, B = 2 and so on...)

Example of a portfolio:
![Example portfolio](http://i.imgur.com/axmDcE0.png)

In this example the `ticker_column` would be 3 and the `price_update_column` would be 6

The script will also attempt to find a cell with the text "Last updated:" and update the cell to the right of it with the current time that the script ran.


### store_end_of_day_value.py
`$ python3 store_end_of_day_value.py <spreadsheet_key> <value_col> <date_col> <value_cell>`  

See the documentation for the above script on how to obtain your `<spreadsheet_key>`

`<value_col>`  - is the column for which you want to copy the value to  
`<date_col>`   - is the column to place the date of the copy in  
`<value_cell>` - is the cell to get the value to copy from

Options for the script:
- `-p` or `--portfolio_worksheet`  
    - The title of the worksheet that holds the value to copy from, defaults to the first worksheet in the spreadsheet  
- `-v` or `--value_worksheet`  
    - The title of the worksheet to copy the value to, defaults to the second worksheet in the spreadsheet


Example of a value_worksheet:  
![Example value worksheet](http://i.imgur.com/vDa94LD.png)

In this example the `<value_col>` would be 2, the `<date_col>` would be 1, and the `<value_cell>` would be G11 from the worksheet in the image from update_my_portfolio.py. In my example spreadsheet the portfolio worksheet is the first worksheet, so I don't need to specify its title with the `-p` option.
