#!/usr/bin/env python3
"""
Test script to verify Excel formatting functionality.
"""

import pandas as pd
from datetime import datetime
import os

def test_excel_formatting():
    """Test the Excel formatting with sample data."""
    
    # Create sample data
    sample_data = [
        {
            "REPORT_DATE": "2025/01/15",
            "RECORD_ID": 1,
            "TENDER_ID": "TENDER-001",
            "PUBLICATION_DATE": "2025/01/10",
            "CLOSING_DATE": "2025/02/15",
            "CLOSING_TIME": "10:00",
            "TENDER_TYPE": "Supply",
            "TENDER_DESCRIPTION": "Test Tender Description",
            "TENDER_SOURCE": "ETENDERS.GOV.ZA",
            "DEPARTMENT": "Test Department",
            "PROVINCE": "Gauteng",
            "ESUBMISSION": "Yes",
            "CATEGORY": "Services",
            "IS_THERE_A_BRIEFING_SESSION": "No",
            "BRIEFING_DATE": "",
            "COMPULSORY_BRIEFING": "No",
            "BRIEFING_SESSION_VENUE": "",
            "LINK": "https://www.etenders.gov.za/test-link",
            "SOE": "",
            "COST_OF_SALES_ESTIMATE": "",
            "CAPABILITY_AVAILABLE": "",
            "CAPABILITY_GROUP": "",
            "REQUIREMENTS": ""
        },
        {
            "REPORT_DATE": "2025/01/15",
            "RECORD_ID": 2,
            "TENDER_ID": "TENDER-002",
            "PUBLICATION_DATE": "2025/01/12",
            "CLOSING_DATE": "2025/02/20",
            "CLOSING_TIME": "14:30",
            "TENDER_TYPE": "Construction",
            "TENDER_DESCRIPTION": "Another Test Tender",
            "TENDER_SOURCE": "ETENDERS.GOV.ZA",
            "DEPARTMENT": "Another Department",
            "PROVINCE": "Western Cape",
            "ESUBMISSION": "No",
            "CATEGORY": "Construction",
            "IS_THERE_A_BRIEFING_SESSION": "Yes",
            "BRIEFING_DATE": "2025/01/25",
            "COMPULSORY_BRIEFING": "Yes",
            "BRIEFING_SESSION_VENUE": "Cape Town Convention Centre",
            "LINK": "https://www.etenders.gov.za/another-test-link",
            "SOE": "",
            "COST_OF_SALES_ESTIMATE": "",
            "CAPABILITY_AVAILABLE": "",
            "CAPABILITY_GROUP": "",
            "REQUIREMENTS": ""
        }
    ]
    
    # Create DataFrame
    df = pd.DataFrame(sample_data)
    
    # Test the formatting function
    from TenderScraper import TenderScraper
    
    # Create a temporary scraper instance just for testing
    scraper = TenderScraper()
    
    # Test the formatting
    test_filename = "test_formatted_excel.xlsx"
    scraper._saveExcelWithFormatting(df, test_filename)
    
    print(f"Test Excel file created: {test_filename}")
    print("Please open the file and verify:")
    print("1. Dates are properly formatted as Excel dates (not text)")
    print("2. Links are clickable hyperlinks")
    print("3. Headers are formatted with blue background and white text")
    print("4. Header row is frozen")
    print("5. Column widths are auto-adjusted")
    
    # Clean up
    if os.path.exists(test_filename):
        print(f"\nTest file {test_filename} has been created successfully!")
        print("You can now open it in Excel to verify the formatting.")

if __name__ == "__main__":
    test_excel_formatting() 