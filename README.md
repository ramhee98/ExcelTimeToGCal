# ExcelTimeToGCal

Automatically create Google Calendar events from a structured Excel timesheet using Python.

## Features

- Reads time entries from one or more sheets in an Excel file.
- Filters by date range (`days_back`) and ignores future dates.
- Authenticates with Google Calendar via OAuth 2.0.
- Optionally replaces existing events with the same description.
- Fully configurable via `config.ini`.

## Requirements

- Python 3.7+
- Google Calendar API enabled
- `credentials.json` in your working directory (from Google Cloud Console)

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

Create a `config.ini` file in the same directory with the content from config.ini.template and adjust it to your needs.

## Usage

1. Ensure your Excel file has the following structure with labeled rows:
   - `Datum`: Work date (e.g. 2025-03-01 00:00:00)
   - `Start`: Start time (e.g. 07:30:00)
   - `Ist zeit`: Duration in hours (e.g. 7.5)

2. Run the script:

```bash
python excel_to_gcal.py
```

The script will:
- Load the Excel file
- Parse all sheets
- Create or replace events in your calendar

## Authentication

On the first run, the script will open a browser window to authenticate with your Google account. It will save a `token.json` for future access.

## License

MIT License
