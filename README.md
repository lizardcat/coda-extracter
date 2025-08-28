# Coda Timesheet Extractor

Extract and process timesheet data from Coda documents using Python.

## Project Overview

This tool allows you to automatically extract timesheet data from your Coda documents using the official Coda API. It processes the raw data into clean CSV files and provides logging and configuration management.

## File Structure

```
timesheet_extractor/
├── README.md
├── requirements.txt
├── config/
│   ├── __init__.py
│   └── config.py
├── src/
│   ├── __init__.py
│   ├── coda_extractor.py
│   └── data_processor.py
├── scripts/
│   └── extract_timesheet.py
├── data/
│   ├── raw/          # Raw JSON responses from Coda API
│   └── processed/    # Cleaned CSV files
├── logs/             # Extraction logs
└── .env              # Your API credentials (keep private!)
```

## Setup Instructions

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

Required packages:

- `requests>=2.31.0` - For API calls to Coda
- `pandas>=2.0.0` - For data processing
- `python-dotenv>=1.0.0` - For environment variable management

### 2. Get Your Coda API Credentials

1. **Get your API token:**

   - Go to your Coda account settings
   - Navigate to the "API" section
   - Generate a new token
   - Copy the token (you'll only see it once!)

2. **Create your `.env` file:**
   ```
   CODA_API_TOKEN=your_api_token_here
   CODA_DOC_ID=your_document_id_here
   CODA_TABLE_ID=your_table_id_here
   ```

### 3. Find Your Document and Table IDs

**Find your documents:**

```bash
python scripts/extract_timesheet.py --list-docs
```

This will show you all documents you have access to with their IDs.

**Find tables in your timesheet document:**

```bash
python scripts/extract_timesheet.py --list-tables YOUR_DOC_ID
```

Replace `YOUR_DOC_ID` with the document ID from the previous step.

### 4. Update Your Configuration

Add the correct document and table IDs to your `.env` file:

```
CODA_API_TOKEN=your_actual_token
CODA_DOC_ID=AbCdEfGhIj
CODA_TABLE_ID=table-KlMnOpQr
```

## Usage

### Basic Extraction

Extract your timesheet data with default settings:

```bash
python scripts/extract_timesheet.py
```

This will:

- Extract data from your configured timesheet
- Save raw JSON data to `data/raw/`
- Process and clean the data
- Export to CSV in `data/processed/`
- Show a summary of the extracted data

### Custom Output Filename

Specify a custom output filename:

```bash
python scripts/extract_timesheet.py --output my_timesheet_2024.csv
```

### List Available Resources

List all your Coda documents:

```bash
python scripts/extract_timesheet.py --list-docs
```

List tables in a specific document:

```bash
python scripts/extract_timesheet.py --list-tables DOC_ID_HERE
```

## Output Files

### Raw Data

- **Location:** `data/raw/timesheet_raw_YYYYMMDD_HHMMSS.json`
- **Content:** Unprocessed JSON response from Coda API
- **Purpose:** Backup and debugging

### Processed Data

- **Location:** `data/processed/timesheet_processed_YYYYMMDD_HHMMSS.csv`
- **Content:** Clean, structured CSV file
- **Purpose:** Ready for analysis in Excel, Google Sheets, or other tools

### Logs

- **Location:** `logs/extraction_YYYYMMDD.log`
- **Content:** Detailed extraction logs with timestamps
- **Purpose:** Troubleshooting and audit trail

## Data Processing Features

The tool automatically:

- Extracts data from Coda's nested JSON format
- Converts date columns to proper datetime format
- Converts hours/duration columns to numeric values
- Handles missing or malformed data gracefully
- Provides summary statistics (total hours, date range, etc.)

## Customization

### Adding Custom Data Cleaning

Edit `src/data_processor.py` in the `clean_timesheet_data()` method to add your own data cleaning rules:

```python
def clean_timesheet_data(self, df):
    cleaned_df = df.copy()

    # Your custom cleaning rules here
    if 'Project' in cleaned_df.columns:
        cleaned_df['Project'] = cleaned_df['Project'].str.strip().str.title()

    return cleaned_df
```

### Scheduling Automatic Extractions

You can set up automatic extractions using cron (Linux/Mac) or Task Scheduler (Windows):

```bash
# Run daily at 6 PM
0 18 * * * cd /path/to/timesheet_extractor && python scripts/extract_timesheet.py
```

## Troubleshooting

### Common Issues

**"Missing required environment variables"**

- Check that your `.env` file exists and has all required variables
- Make sure there are no extra spaces around the `=` signs

**"Authentication failed"**

- Verify your API token is correct
- Check that the token hasn't expired
- Ensure you have access to the specified document

**"Table not found"**

- Use `--list-tables` to verify the table ID
- Make sure you're using the table ID (starts with "table-"), not the table name

**"No data extracted"**

- Check that your timesheet table has data
- Verify you have read permissions on the document
- Look at the log files in the `logs/` directory for detailed error messages

### Getting Help

1. Check the log files in `logs/` directory
2. Run with `--list-docs` and `--list-tables` to verify your IDs
3. Test with a simple document first to verify your setup

## Security Notes

- Never commit your `.env` file to version control
- Keep your API token secure and rotate it periodically
- The tool only reads data from your Coda documents (no write access)
- All data is stored locally on your machine

## License

This project is for personal/internal use. Modify as needed for your specific timesheet format and requirements.
