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
This may require sudo elevation or administrator elevation depending on your system

## Running the script
### portfolio_manager.py
Running the script in this manner will update the current prices only
`$ python3 portfolio_manager.py <spreadsheet_key> -u 6 -c 13 -t 3`

You can get your `spreadsheet_key` from the URL of your spreadsheet:
![Spreadsheet Key Image](http://i.imgur.com/v666kdf.png)


To see the help for the script just use the `-h` option when running the script.

Example of a portfolio:
![Example portfolio](http://i.imgur.com/axmDcE0.png)

In this example the `ticker_column` would be 3 and the `price_update_column` would be 6

The script will also attempt to find a cell with the text "Last updated:" and update the cell to the right of it with the current time that the script ran.


### store_end_of_day_value.py
`$ python3 portfolio_manager.py <spreadhseet_key> -s -x 2 -d 1 -z G11`  

The `-s` option initiates the saving function and will copy the value in the cell given to the `-z` option to the last empty sell in the column specified by the `-x` option.


Example of a value_worksheet:  
![Example value worksheet](http://i.imgur.com/vDa94LD.png)

In this example the `<value_col>` would be 2, the `<date_col>` would be 1, and the `<value_cell>` would be G11 from the worksheet in the image from update_my_portfolio.py. In my example spreadsheet the portfolio worksheet is the first worksheet, so I don't need to specify its title with the `-p` option.

## Best way to run the scripts
I highly suggest using [cron](https://en.wikipedia.org/wiki/Cron) to run these scripts on a schedule so that the values can be updated as often as possible.

An example cron might be:
```
*/15 07-15 * * 1-5 root /usr/bin/python3 portfolio_manager.py <spreadsheet_key> -u 6 -c 13 -t 3
15 15 * * 1-5 root /usr/bin/python3 /home/user/portfolio_manager/portfolio_manager.py <spreadhseet_key> -s -x 2 -d 1 -z G11
```

This cron runs the update_my_portfolio.py script every 15 minutes, from 7am to 3pm, Monday - Friday. It then runs store_end_of_day_value.py at 3:15pm, Monday - Friday.
