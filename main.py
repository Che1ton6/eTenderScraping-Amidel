#!/usr/bin/env python3
"""
Main entry point for the eTenders Scraper.

This script initializes and runs the TenderScraper with proper error handling
and logging. It serves as the primary interface for running the scraping process.
"""

import sys
import logging
from datetime import datetime
from TenderScraper import TenderScraper

def main():
    """
    Main function to run the tender scraping process.
    """
    print("=" * 60)
    print("eTenders Scraper - Starting Process")
    print("=" * 60)
    print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    try:
        # Initialize the scraper
        scraper = TenderScraper()
        
        # Run the scraping process
        scraper.run()
        
        print()
        print("=" * 60)
        print("eTenders Scraper - Process Completed Successfully")
        print("=" * 60)
        print(f"End Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Total Tenders Scraped: {len(scraper.tenderData)}")
        print()
        
    except FileNotFoundError as e:
        print(f"Configuration Error: {e}")
        print("Please ensure config.json exists and is properly formatted.")
        sys.exit(1)
        
    except ValueError as e:
        print(f"Configuration Error: {e}")
        print("Please check your configuration settings in config.json.")
        sys.exit(1)
        
    except KeyboardInterrupt:
        print()
        print("Process interrupted by user.")
        print("Cleaning up resources...")
        sys.exit(0)
        
    except Exception as e:
        print(f"Unexpected Error: {e}")
        print("Check the log file for detailed error information.")
        logging.error(f"Unexpected error in main: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main() 