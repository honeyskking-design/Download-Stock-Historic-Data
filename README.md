# Indian Stock Historical Data Downloader

A Python-based automation tool for downloading historical stock-market data for Indian companies from **NSE and BSE**.

The script reads a list of stock tickers from an Excel file, searches for each ticker on NSE first, automatically falls back to BSE when the ticker is not available on NSE, downloads the historical data as an Excel file, and consolidates all downloaded companies into a single master workbook.

## Features

* 📊 Download historical stock data for multiple companies
* 🇮🇳 NSE-first search with automatic BSE fallback
* 🔄 Automatically moves to BSE when an NSE ticker is not found
* 📅 Configurable historical date range
* 📁 Saves individual Excel files for every successfully downloaded company
* 📚 Creates a consolidated master workbook
* 📑 Stores each company on a separate worksheet in the master workbook
* 📝 Maintains a download log showing successful and failed downloads
* 🔁 Removes duplicate tickers while preserving their original order
* 🧹 Cleans invalid Windows filenames automatically
* 📋 Handles Excel worksheet-name restrictions
* ⏱️ Configurable download timeout
* 🖥️ Runs Chromium in headless/background mode
* 💾 Saves the download log after every company, reducing the risk of losing progress

## How It Works

The workflow is:

```text
CompanyTicker.xlsx
        │
        ▼
Read stock tickers from Column A
        │
        ▼
     Try NSE
        │
   ┌────┴────┐
   │         │
 Found     Not Found
   │         │
   ▼         ▼
Download    Try BSE
   │         │
   │      ┌──┴──┐
   │      │     │
   │    Found  Not Found
   │      │     │
   ▼      ▼     ▼
Save Individual Excel File
        │
        ▼
Add company data to Master Workbook
        │
        ▼
Update Download Log
```

The script checks the NSE autocomplete result before attempting a download. If no matching NSE ticker is found, it immediately proceeds to BSE instead of waiting for the full download timeout.

## Data Source

The project uses:

* **NSE Historical Data:** `https://www.nsehistoricaldata.co.in/`
* **BSE Historical Data:** `https://bse.nsehistoricaldata.co.in/`

The script attempts NSE first and uses BSE as a fallback.

> **Note:** This project automates interaction with third-party websites. Availability, website structure, download functionality, and access restrictions may change over time.

## Project Structure

A typical project folder looks like this:

```text
Company historic data - Code/
│
├── historical_data_downloader.py
├── CompanyTicker.xlsx
│
└── Company Historic Data/
    ├── Master_Historical_Data.xlsx
    ├── Download_Log.xlsx
    │
    ├── TCS_NSE_2015-04-01_to_2026-03-31.xlsx
    ├── RELIANCE_NSE_2015-04-01_to_2026-03-31.xlsx
    ├── INFY_NSE_2015-04-01_to_2026-03-31.xlsx
    └── ...
```

## Input File

The program expects an Excel file named:

```text
CompanyTicker.xlsx
```

Stock tickers should be placed in **Column A**, beginning below the header.

Example:

| Company Ticker |
| -------------- |
| TCS            |
| INFY           |
| RELIANCE       |
| HDFCBANK       |
| SBIN           |

The script reads the first column of the workbook, removes blank entries, and removes duplicate tickers while preserving their original order.

## Configuration

The main configuration can be modified at the beginning of the Python script.

### File Location

```python
BASE_FOLDER = Path(
    r"D:\Documents(H)\Company historic data - Code"
)

INPUT_FILE = BASE_FOLDER / "CompanyTicker.xlsx"

OUTPUT_FOLDER = BASE_FOLDER / "Company Historic Data"
```

Change `BASE_FOLDER` according to your own computer.

### Date Range

The current configuration downloads data from:

```python
FROM_DATE = "2015-04-01"
TO_DATE = "2026-03-31"
```

You can change these dates depending on the required historical period.

### Download Timeout

The maximum time allowed for a download to start is:

```python
DOWNLOAD_TIMEOUT = 120000
```

This represents **120 seconds / 2 minutes**.

Importantly, this is a maximum timeout rather than a mandatory two-minute delay. If the website starts the download after 20 seconds, the script continues immediately.

### Delay Between Companies

```python
DELAY_BETWEEN_COMPANIES = 2
```

This adds a two-second pause between companies.

### Headless Browser

```python
HEADLESS = True
```

When enabled, Chromium runs in the background without opening a visible browser window.

## Installation

### 1. Install Python

Python 3.9+ is recommended.

Check your Python installation:

```bash
python --version
```

or:

```bash
py --version
```

### 2. Install Required Libraries

Install the required Python packages:

```bash
pip install pandas openpyxl playwright
```

The project uses:

| Library      | Purpose                               |
| ------------ | ------------------------------------- |
| `pandas`     | Reading and processing Excel data     |
| `openpyxl`   | Creating and updating Excel workbooks |
| `playwright` | Browser automation and file downloads |
| `re`         | Filename and worksheet-name cleaning  |
| `time`       | Delays between companies              |
| `pathlib`    | File and folder management            |

The core external dependencies are `pandas`, `openpyxl`, and `playwright`.

### 3. Install Playwright Browser

After installing Playwright, install Chromium:

```bash
playwright install chromium
```

## Running the Program

Place your ticker list in:

```text
CompanyTicker.xlsx
```

Then run:

```bash
python historical_data_downloader.py
```

On Windows, you can also use:

```bash
py historical_data_downloader.py
```

The browser runs in the background because the script is configured with:

```python
HEADLESS = True
```

## Output

The script creates a folder called:

```text
Company Historic Data
```

### Individual Company Files

Each successful download is saved separately.

Example:

```text
TCS_NSE_2015-04-01_to_2026-03-31.xlsx
```

The filename identifies:

* Company ticker
* Exchange
* Historical date range

The script also sanitizes ticker names before using them as Windows filenames.

### Master Workbook

The program creates:

```text
Master_Historical_Data.xlsx
```

Each company is stored on its own worksheet.

Example:

```text
README
TCS_NSE
INFY_NSE
RELIANCE_NSE
HDFCBANK_NSE
SBIN_BSE
```

The workbook contains a `README` sheet describing the date range, data source, and workbook structure.

The company worksheets also receive basic formatting such as frozen headers and adjusted column widths.

### Download Log

The program creates:

```text
Download_Log.xlsx
```

The log records:

* Ticker
* Exchange
* Download status
* Individual file path
* Master workbook status
* Error message, if applicable

The log is saved after every processed company rather than only at the end of the program.

## NSE to BSE Fallback

The downloader follows this logic:

```text
Ticker
  │
  ▼
Search NSE
  │
  ├── Match found ──► Download from NSE
  │
  └── No match
          │
          ▼
       Search BSE
          │
          ├── Match found ──► Download from BSE
          │
          └── No match ──► Mark as FAILED
```

This is particularly useful when a company is listed on BSE but does not have a matching NSE ticker.

The script does **not** attempt an NSE download when the autocomplete search fails. Instead, it immediately proceeds to BSE.

## Error Handling

The program handles several common situations, including:

* Missing input Excel file
* Empty input workbook
* No tickers in Column A
* Ticker not found on NSE
* Ticker not found on BSE
* Download timeout
* Empty downloaded Excel file
* Failure while adding data to the master workbook
* Invalid Windows filename characters
* Duplicate Excel worksheet names

If a ticker cannot be downloaded from either exchange, it is recorded as failed in the download log instead of stopping the entire process.

## Example Console Output

A typical run looks similar to:

```text
===========================================================================
       INDIAN STOCK HISTORICAL DATA DOWNLOADER
===========================================================================

Input file:
  D:\Documents(H)\Company historic data - Code\CompanyTicker.xlsx

Date range:
  2015-04-01 -> 2026-03-31

Output folder:
  D:\Documents(H)\Company historic data - Code\Company Historic Data

Individual files:
  Yes

Master workbook:
  ...\Master_Historical_Data.xlsx

Download timeout:
  2 minutes
===========================================================================
```

At the end, the program reports:

```text
===========================================================================
PROCESS COMPLETE
===========================================================================

Total tickers : 10
Downloaded    : 9
Failed        : 1

Output folder:
...

Master workbook:
...

Download log:
...
```

The final summary is generated from the recorded processing results.

## Important Considerations

### Website Changes

Because the project relies on browser automation, changes to the source websites can break selectors or page interactions.

For example, the script searches for specific input fields, autocomplete results, date inputs, and the historical-data download button.

If the website changes its HTML structure, these selectors may need to be updated.

### Download Time

Historical data generation can take time on the source website. The current maximum download-start timeout is two minutes.

```python
DOWNLOAD_TIMEOUT = 120000
```

This should be increased if the source website regularly takes longer than two minutes to generate downloads.

### Large Number of Companies

Downloading a large number of companies may take considerable time because each ticker is processed individually.

The program intentionally adds a short delay between companies:

```python
DELAY_BETWEEN_COMPANIES = 2
```

## Use Cases

This tool can be useful for:

* Equity research
* Financial analysis
* Historical stock-price analysis
* Technical analysis
* Volatility analysis
* Return calculations
* Drawdown analysis
* Trading-volume analysis
* Delivery-volume analysis
* Research datasets
* Academic finance projects
* Building historical market-data databases

## Possible Future Improvements

Potential enhancements include:

* [ ] Automatic retry mechanism
* [ ] Resume interrupted downloads
* [ ] Progress bar
* [ ] Parallel processing with controlled concurrency
* [ ] Automatic data validation
* [ ] Detection of duplicate or incomplete downloaded files
* [ ] Summary statistics for each company
* [ ] Automated return and volatility calculations
* [ ] Corporate-action adjustment checks
* [ ] Configurable output date range from the input Excel file
* [ ] Command-line arguments for configuration
* [ ] Structured logging instead of console-only messages
* [ ] Automated data-quality report

## Disclaimer

This project is intended for **educational, research, and data-analysis purposes**.

The project does not provide investment advice or investment recommendations.

Users are responsible for complying with the terms of use, access policies, and applicable restrictions of the data sources they interact with.

Historical market data should be independently validated before being used for financial or investment decisions.

## License

Choose an appropriate license before publishing the repository.

For example, if you want a permissive open-source license, you could use the **MIT License**.

---

## Author

**Shailesh Kumar**

Financial Analyst | Valuation & Financial Modeling | Equity Research

GitHub: *Add your GitHub profile URL here*

LinkedIn: *Add your LinkedIn profile URL here*
