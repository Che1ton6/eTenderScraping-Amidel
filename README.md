# eTenders Scraper

Web scraper for extracting tender data from the South African eTenders website (etenders.gov.za).

**Version:** 1.0.0

## Features

- **Error Handling**: Retry logic for stale elements and network issues
- **Duplicate Prevention**: Automatic detection and filtering of duplicate tenders
- **Configurable**: Configuration through JSON file
- **Logging**: Detailed logging for debugging and monitoring
- **Data Validation**: Automatic validation and cleaning of scraped data
- **Dual Output**: Generates both date-specific and cumulative Excel files


## Project Structure

```
EtenderScraping/
├── TenderScraper.py         # Main scraper class
├── Utils.py                 # Utility functions
├── ConfigManager.py         # Configuration management
├── config.json              # Configuration settings
├── main.py                  # Entry point
├── requirements.txt         # Python dependencies
├── README.md               # This file
├── data/                   # Output Excel files
│   ├── tenders_*.xlsx      # Date-specific tender data
│   └── master_tenders.xlsx # Master file with all tenders
└── logs/                   # Log files
    └── scraper.log         # Scraping logs
```

## Installation

1. **Clone or download the project files**

2. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install Chrome WebDriver:**
   - Download ChromeDriver from: https://chromedriver.chromium.org/
   - Ensure it's in your system PATH or in the project directory

## Configuration

Edit `config.json` to customize the scraper:

```json
{
    "scraping": {
        "dateFrom": "2025-06-25",
        "dateTo": "2025-06-30",
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
        "expandRowWait": 3,
        "collapseRowWait": 2,
        "nextPageWait": 4,
        "retryDelay": 2.5
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

### Configuration Options

- **dateFrom/dateTo**: Date range for scraping (YYYY-MM-DD format)
- **headless**: Run browser in headless mode (no visible window)
- **maxRetries**: Number of retry attempts for failed operations
- **pageLoadWait**: Time to wait for page loading (seconds)

## Usage

### Basic Usage

Run the scraper with default configuration:

```bash
python main.py
```

### Custom Configuration

1. Edit `config.json` with your desired settings
2. Run the scraper:
   ```bash
   python main.py
   ```

## Output Files

The scraper generates two Excel files in the `data/` folder:

1. **Date-specific file**: `data/tenders_30_06_2025.xlsx` (tenders for specific date only)
2. **Master file**: `data/master_tenders.xlsx` (all tenders combined, with duplicate prevention)

### Excel Column Structure

| Column | Description |
|--------|-------------|
| REPORT_DATE | Date when scraper was run |
| RECORD_ID | Auto-incrementing ID |
| TENDER_ID | Tender number |
| PUBLICATION_DATE | When tender was published |
| CLOSING_DATE | Closing date |
| CLOSING_TIME | Closing time |
| TENDER_TYPE | Type of tender |
| TENDER_DESCRIPTION | Description |
| TENDER_SOURCE | Always "ETENDERS.GOV.ZA" |
| DEPARTMENT | Organ of state |
| PROVINCE | Province |
| ESUBMISSION | E-submission status |
| CATEGORY | Tender category |
| IS_THERE_A_BRIEFING_SESSION | Briefing session status |
| BRIEFING_DATE | Briefing date |
| COMPULSORY_BRIEFING | Compulsory status |
| BRIEFING_SESSION_VENUE | Briefing venue |
| LINK | Document link |
| SOE | State-owned enterprise |
| COST_OF_SALES_ESTIMATE | Cost estimate |
| CAPABILITY_AVAILABLE | Capability status |
| CAPABILITY_GROUP | Capability group |
| REQUIREMENTS | Requirements |

## Logging

The scraper creates detailed logs in `logs/scraper.log` including:
- Scraping progress
- Error messages
- Performance metrics
- Data validation results

## Error Handling

The scraper includes robust error handling for:
- **Stale Elements**: Automatic retry with element refetching
- **Network Issues**: Retry logic for failed requests
- **Invalid Data**: Validation and cleaning of scraped data
- **Browser Crashes**: Graceful cleanup and recovery

## Troubleshooting

### Common Issues

1. **ChromeDriver not found**
   - Ensure ChromeDriver is installed and in PATH
   - Or place ChromeDriver in the project directory

2. **Configuration errors**
   - Check `config.json` format and required fields
   - Ensure date format is YYYY-MM-DD

3. **No data scraped**
   - Verify date range in configuration
   - Check if website structure has changed
   - Review log file for error messages

4. **Slow performance**
   - Reduce wait times in configuration
   - Enable headless mode
   - Check internet connection

### Debug Mode

Enable debug logging by changing the log level in `config.json`:

```json
{
    "logging": {
        "level": "DEBUG",
        "file": "logs/scraper.log"
    }
}
```

## Development

### Project Structure

The project follows a clean, modular structure:

- **`main.py`**: Entry point script with proper error handling
- **`TenderScraper.py`**: Main scraper class implementing OOP principles
- **`Utils.py`**: Utility functions for data processing and validation
- **`config/`**: Configuration package containing settings and management
- **`data/`**: Output directory for Excel files
- **`logs/`**: Application logs

### Code Quality

- **Object-Oriented Design**: Proper class structure with clear separation of concerns
- **Error Handling**: Comprehensive exception handling and logging
- **Configuration Management**: Centralized configuration with validation
- **Data Validation**: Input validation and data cleaning
- **Documentation**: Comprehensive docstrings and comments

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Guidelines

- Follow PEP 8 style guidelines
- Add comprehensive docstrings for all functions and classes
- Include error handling for all external operations
- Write unit tests for new functionality
- Update documentation for any configuration changes

### Code Standards

- **File/Class Names**: CamelCase (e.g., `TenderScraper.py`, `class TenderScraper`)
- **Method Names**: camelCase (e.g., `parseDate()`, `loadConfig()`)
- **Documentation**: Comprehensive docstrings for all classes and methods

### Adding Features

1. Follow the established naming conventions
2. Add proper error handling
3. Include logging for debugging
4. Update this README with new features

## License

This project is for educational and business use. Please respect the website's terms of service.

## Support

For issues and questions:
1. Check the log file for detailed error information
2. Review the troubleshooting section
3. Ensure all dependencies are properly installed 