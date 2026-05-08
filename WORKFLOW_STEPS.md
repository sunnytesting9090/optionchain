# Option Chain Workflow

## Files in this folder

- `config.py`: stores your Zerodha API credentials.
- `access_token.txt`: stores today's access token.
- `update_option_chain.py`: pulls live data from Kite and writes the Excel sheet.
- `start_option_chain_updater.bat`: starts the Python updater by using the saved token.
- `run_option_chain.bat`: asks for today's access token, saves it, then starts the updater.
- `option_chain.xlsx`: Excel workbook with dropdowns for index and expiry.

## What the Python script does

1. Reads the API key from `config.py`.
2. Reads today's access token from `access_token.txt`.
3. Connects to Kite and checks the login.
4. Refreshes the `NFO` and `BFO` instrument caches if they are stale.
5. Builds or repairs `option_chain.xlsx` if the workbook structure is missing or outdated.
6. Opens Excel automatically.
7. Loads dropdown choices for `NIFTY`, `BANKNIFTY`, and `SENSEX`.
8. Reads the selected index and expiry from Excel.
9. Pulls live CE and PE prices around the ATM strike.
10. Calculates:
    - `Straddle Sum` for every strike = `Call LTP + Put LTP`
    - `ATM Straddle Sum`
    - `Displayed Straddle Total`
11. Writes the data back into Excel and keeps refreshing in a loop.

## Daily usage

1. Double-click `run_option_chain.bat`.
2. Paste today's access token.
3. Press `Enter`.
4. Excel opens automatically.
5. Change `Index` from the dropdown between `NIFTY`, `BANKNIFTY`, and `SENSEX`.
6. Change `Expiry` from the dropdown.
7. Keep the updater console window open for live refreshes.

## Direct usage with existing token

1. Make sure `access_token.txt` already contains a valid token.
2. Double-click `start_option_chain_updater.bat`.
3. Excel opens and the workbook starts refreshing.

## Notes

- `NIFTY` and `BANKNIFTY` are pulled from the `NFO` option segment.
- `SENSEX` is pulled from the `BFO` option segment.
- The workbook refresh interval is 10 seconds by default.
- If the token expires, run `run_option_chain.bat` again and paste a new token.
