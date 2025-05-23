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

## Google Cloud Setup

To use this script, you must enable the Google Calendar API and download credentials:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (or select an existing one).
3. Navigate to **APIs & Services > Library**.
4. Search for **Google Calendar API** and click **Enable**.
5. Go to **APIs & Services > Credentials**.
6. Click **Create Credentials > OAuth client ID**.
   - Choose **Desktop App**.
   - Name it e.g. `ExcelTimeToGCal`.
7. Download the generated `credentials.json` file and place it in your project directory.

> The script will use this to authenticate and create `token.json` on first run.

## Authentication

On the first run, the script will open a browser window to authenticate with your Google account. It will save a `token.json` for future access.

## License

MIT License
