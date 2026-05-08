# Simple Option Chain Workflow

## Files

- `simple_option_chain.xlsx`: the new simple Excel workbook.
- `update_simple_option_chain.py`: fetches live option chain data for `NIFTY`, `BANKNIFTY`, and `SENSEX`.
- `run_simple_option_chain.bat`: asks for today's access token and starts the updater.
- `start_simple_option_chain_updater.bat`: starts the updater by using the token already saved in `access_token.txt`.

## What this version does

1. Opens one Excel workbook.
2. Updates all 3 indices directly:
   - `NIFTY`
   - `BANKNIFTY`
   - `SENSEX`
3. Uses the nearest available expiry automatically.
4. Streams live ticks by WebSocket instead of waiting on a timer loop.
5. Writes the full option chain for that expiry.
6. Calculates `Straddle Sum` for every strike.
7. Shows `Spot`, `ATM Strike`, and `ATM Straddle`.
8. Adds a `RawData` sheet with one row per live instrument for Excel formulas.
9. Displays only `ATM +25` to `ATM -25` strikes on each index sheet.

## Daily use

1. Double-click `run_simple_option_chain.bat`.
2. Paste today's access token.
3. Press `Enter`.
4. `simple_option_chain.xlsx` opens automatically.
5. Leave the updater window open while you want live refreshes.

## Notes

- This version is simpler and faster because it uses a live WebSocket stream and fixed sheets.
- `NIFTY` and `BANKNIFTY` are read from `NFO`.
- `SENSEX` is read from `BFO`.
- `RawData` contains the live fields for each instrument so you can build formulas on top of it.
- The workbook autosaves every `60` seconds while it is streaming.
