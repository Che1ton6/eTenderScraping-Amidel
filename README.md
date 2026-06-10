# Amidel eTender Scraper

Web scraper for extracting tender data from the South African eTenders website (etenders.gov.za), built for Amidel (Pty) Ltd.

**Version:** 2.0.0

## Features

- **Amidel-branded GUI**: Navy and orange themed desktop application — no command-line required
- **Batch week picker**: Calendar popup that automatically highlights the correct Mon–Wed (Thursday batch) or Thu–Sun (Monday batch) range when you click any date in the week
- **Batch folder structure**: Organises all output into a dated batch folder with three subfolders — `batches/`, `end product/`, and `Display Equation/`
- **Per-day Excel files**: Each scraped day is saved individually inside `batches/`
- **End product file**: Combined `RFQ_and_ICT_Checker` workbook with a `Tender Data` tab and a `Final` tab containing source-level COUNTIF formulas
- **Equation file management**: Automatically inserts the new batch at the left of `RFQ_and_ICT_Equation.xlsx`, shifts older batches right, enforces a maximum of 6 batches, and copies the updated file into the `Display Equation/` subfolder
- **Error handling**: Retry logic for stale elements and network issues
- **Duplicate prevention**: Automatic detection and filtering of duplicate tenders
- **Live logging**: Log output streamed to the GUI and written to `logs/scraper.log`

## Project Structure

```
eTenderScraping/
├── main.py                  # GUI entry point (tkinter, Amidel-branded)
├── TenderScraper.py         # Selenium scraper class
├── BatchProcessor.py        # Batch folder creation, Excel output, equation file update
├── ConfigManager.py         # Configuration management
├── Utils.py                 # Utility functions
├── config.json              # Scraper configuration
├── amidel.ico               # Application icon
├── requirements.txt         # Python dependencies
├── README.md                # This file
├── data/                    # Batch output folders
│   └── (T) 18-20 May 2026/ # Example batch folder
│       ├── batches/         # Per-day Excel files (tenders_YYYY_MM_DD.xlsx)
│       ├── end product/     # Combined RFQ_and_ICT_Checker workbook
│       └── Display Equation/# Copy of the updated equation file
└── logs/
    └── scraper.log
```

## Installation

1. **Clone or download the project files**

2. **Install Python dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```

3. **Chrome** must be installed — ChromeDriver is managed automatically by Selenium.

## Usage

Run the application:

```powershell
python main.py
```

### Picking a batch

1. Select **T** (Thursday batch — Mon to Wed) or **M** (Monday batch — Thu to Sun) using the toggle buttons.
2. Click **Pick Week** to open the calendar.
3. Click any date within the target week — the calendar will automatically highlight the full batch range.
4. Click **Confirm** to lock in the selection, then **Run Scraper**.

Chrome will open automatically and scrape each day in the range one at a time. When complete, the screen switches to a summary showing how many tenders were found.

### Output produced per run

| Location | File |
|----------|------|
| `data/(T) DD-DD Mon YYYY/batches/` | `tenders_YYYY_MM_DD.xlsx` for each scraped day |
| `data/(T) DD-DD Mon YYYY/end product/` | `RFQ_and_ICT_Checker_(DD - DD Month).xlsx` |
| `data/(T) DD-DD Mon YYYY/Display Equation/` | Updated copy of `RFQ_and_ICT_Equation.xlsx` |
| OneDrive (original) | `RFQ_and_ICT_Equation.xlsx` updated in-place |

## Configuration

`config.json` controls scraper behaviour. Dates are set automatically by the GUI.

```json
{
    "scraping": {
        "dateFrom": "2026-05-19",
        "dateTo": "2026-05-19",
        "url": "https://www.etenders.gov.za/Home/opportunities"
    },
    "browser": {
        "headless": false,
        "maximized": true,
        "disableExtensions": true,
        "disableInfobars": true
    },
    "timing": {
        "pageLoadWait": 7,
        "modalRemovalWait": 1,
        "expandRowWait": 2,
        "collapseRowWait": 1,
        "nextPageWait": 3,
        "retryDelay": 2
    },
    "retry": {
        "maxRetries": 3,
        "staleElementRetries": 3
    },
    "output": {
        "dateSpecificFile": "data/tenders_{date}.xlsx",
        "cumulativeFile": "data/master_tenders.xlsx",
        "dateFormat": "%d_%m_%Y"
    },
    "logging": {
        "level": "INFO",
        "file": "logs/scraper.log"
    }
}
```

### Timing — Important Note

> **Do not reduce `pageLoadWait` below `7`.**
>
> The eTenders website uses a jQuery DataTable that requires ~7 seconds to fully initialise and render pagination buttons. Setting this lower causes the scraper to see only a single page of results with no pagination, returning 0 tenders.

The other timing values (`expandRowWait`, `collapseRowWait`, `nextPageWait`) are safe to reduce if you want faster scraping.

## Equation File

The equation file (`RFQ_and_ICT_Equation.xlsx`) tracks NNT / ICT / RFQ counts across up to **6 batches**:

- Each new batch is inserted at **column B** (left of existing data), pushing older batches right.
- Once a 7th batch would be added, the oldest (rightmost) batch is automatically removed.
- After every update the file is copied to the current batch's `Display Equation/` folder.
- Thin borders are applied to the full data range on every save.

The equation file must be **closed in Excel** before running the scraper, otherwise the save will fail with a permission error.

## Excel Columns (Tender Data)

| Column | Description |
|--------|-------------|
| REPORT_DATE | Date the scraper was run |
| RECORD_ID | Auto-assigned ID (highest = first scraped) |
| TENDER_ID | Tender number |
| PUBLICATION_DATE | When the tender was published |
| CLOSING_DATE | Closing date |
| CLOSING_TIME | Closing time |
| TENDER_TYPE | Type of tender |
| TENDER_DESCRIPTION | Description |
| TENDER_SOURCE | Always "ETENDERS.GOV.ZA" |
| DEPARTMENT | Organ of state |
| PROVINCE | Province |
| ESUBMISSION | E-submission available |
| CATEGORY | Tender category |
| IS_THERE_A_BRIEFING_SESSION | Briefing session status |
| BRIEFING_DATE | Briefing date |
| COMPULSORY_BRIEFING | Whether briefing is compulsory |
| BRIEFING_SESSION_VENUE | Briefing venue |
| LINK | Document download link (clickable hyperlink) |
| SOE | State-owned enterprise flag |
| COST_OF_SALES_ESTIMATE | Cost estimate |
| CAPABILITY_AVAILABLE | Capability status |
| CAPABILITY_GROUP | Capability group |
| REQUIREMENTS | Requirements |

## Troubleshooting

### 0 tenders scraped
- The most common cause is `pageLoadWait` being too low — ensure it is set to `7`.
- Check `logs/scraper.log` for detail on what dates were found.

### Missing days in the batches folder
- If no tenders were published on a given date, the eTenders site returns nothing and no file is created for that day. This is expected behaviour.

### Equation file permission error
- Close `RFQ_and_ICT_Equation.xlsx` in Excel before running the scraper.

### ChromeDriver errors
- Ensure Google Chrome is installed. Selenium manages ChromeDriver automatically.

### Debug mode
```json
"logging": { "level": "DEBUG", "file": "logs/scraper.log" }
```

## License

For internal business use by Amidel (Pty) Ltd. Please respect the eTenders website's terms of service.
