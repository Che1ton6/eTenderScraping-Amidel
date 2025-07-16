import re
import logging
from datetime import datetime
from typing import Tuple, Optional, Any

class Utils:
    """
    Utility class providing helper functions for data processing, date parsing,
    text cleaning, and other common operations used by the TenderScraper.
    
    Note: All methods in this class are stateless and could be converted to static methods
    or module-level functions if desired. Currently kept as instance methods for
    consistency with existing codebase and future flexibility.
    """
    
    def parseClosingDateTime(self, dateStr: str) -> Tuple[str, str]:
        """
        Parse a closing date and time string into separate date and time components.
        
        Args:
            dateStr (str): Date string like "Thursday, 31 July 2025 - 10:00"
            
        Returns:
            Tuple[str, str]: (date in YYYY/MM/DD format, time in HH:MM format)
        """
        if not dateStr or not isinstance(dateStr, str):
            return "", ""
        
        # Pattern to match "Thursday, 31 July 2025 - 10:00"
        pattern = r"(\d{1,2})\s+(\w+)\s+(\d{4})\s*-\s*(\d{2}:\d{2})"
        match = re.search(pattern, dateStr)
        
        if match:
            day, monthName, year, timeStr = match.groups()
            try:
                monthNumber = datetime.strptime(monthName, "%B").month
                dateFmt = f"{int(year)}/{monthNumber:02d}/{int(day):02d}"
                return dateFmt, timeStr
            except ValueError:
                logging.warning(f"Could not parse month name: {monthName}")
                return dateStr, ""
        
        logging.warning(f"Could not parse closing date: {dateStr}")
        return dateStr, ""
    
    def parseDayMonthYear(self, dateStr: str) -> str:
        """
        Parse a date string containing day, month, and year into YYYY/MM/DD format.
        
        Args:
            dateStr (str): Date string like "Sunday, 15 June 2025"
            
        Returns:
            str: Date in YYYY/MM/DD format or original string if parsing fails
        """
        if not dateStr or not isinstance(dateStr, str):
            return ""
        
        # Pattern to match "Sunday, 15 June 2025"
        pattern = r"(\d{1,2})\s+(\w+)\s+(\d{4})"
        match = re.search(pattern, dateStr)
        
        if match:
            day, monthName, year = match.groups()
            try:
                monthNumber = datetime.strptime(monthName, "%B").month
                dateFmt = f"{int(year)}/{monthNumber:02d}/{int(day):02d}"
                return dateFmt
            except ValueError:
                logging.warning(f"Could not parse month name: {monthName}")
                return dateStr
        
        logging.warning(f"Could not parse date: {dateStr}")
        return dateStr
    
    def cleanText(self, text: Any) -> str:
        """
        Clean text data by removing illegal characters and normalizing whitespace.
        
        Args:
            text: Input text to clean
            
        Returns:
            str: Cleaned text string
        """
        if text is None:
            return ""
        
        if not isinstance(text, str):
            text = str(text)
        
        # Remove ASCII control characters except tab (\x09), LF (\x0A), CR (\x0D)
        text = re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]', '', text)
        
        # Normalize whitespace (replace multiple spaces/tabs with single space)
        text = re.sub(r'\s+', ' ', text)
        
        # Remove excessive line breaks
        text = re.sub(r'\n+', '\n', text)
        
        # Strip leading/trailing whitespace
        text = text.strip()
        
        return text
    
    def generateUniqueKey(self, tenderInfo: dict) -> str:
        """
        Generate a unique identifier for a tender based on TENDER_ID and PUBLICATION_DATE.
        
        Args:
            tenderInfo (dict): Dictionary containing tender information
            
        Returns:
            str: Unique key for the tender
        """
        tenderId = tenderInfo.get('TENDER_ID', '')
        pubDate = tenderInfo.get('PUBLICATION_DATE', '')
        
        # Clean the values to ensure consistent key generation
        tenderId = self.cleanText(tenderId)
        pubDate = self.cleanText(pubDate)
        
        return f"{tenderId}_{pubDate}"
    
    # On retries in the TenderScraper, some duplicates are encountered and this method helps to clean it up.
    def isDuplicate(self, tenderInfo: dict, processedTenders: set) -> bool:
        """
        Check if a tender is a duplicate based on unique key.
        
        Args:
            tenderInfo (dict): Dictionary containing tender information
            processedTenders (set): Set of already processed tender keys
            
        Returns:
            bool: True if tender is a duplicate, False otherwise
        """
        uniqueKey = self.generateUniqueKey(tenderInfo)
        return uniqueKey in processedTenders
    

    
    def validateTenderData(self, tenderInfo: dict) -> bool:
        """
        Validate tender data for required fields and data quality.
        
        Args:
            tenderInfo (dict): Dictionary containing tender information
            
        Returns:
            bool: True if data is valid, False otherwise
        """
        # Check for required fields
        requiredFields = ['TENDER_ID', 'TENDER_DESCRIPTION', 'PUBLICATION_DATE', 'CLOSING_DATE']
        
        for field in requiredFields:
            if not tenderInfo.get(field):
                logging.warning(f"Missing required field: {field}")
                return False
        
        # Validate publication date format
        pubDate = tenderInfo.get('PUBLICATION_DATE', '')
        if pubDate:
            try:
                datetime.strptime(pubDate, "%Y/%m/%d")
            except ValueError:
                logging.warning(f"Invalid publication date format: {pubDate}")
                return False
        
        # Validate closing date format (if not empty)
        closingDate = tenderInfo.get('CLOSING_DATE', '')
        if closingDate and closingDate != "":
            try:
                datetime.strptime(closingDate, "%Y/%m/%d")
            except ValueError:
                logging.warning(f"Invalid closing date format: {closingDate}")
                return False
        
        return True
    
    def cleanTenderData(self, tenderInfo: dict) -> dict:
        """
        Clean all string values in tender data.
        
        Args:
            tenderInfo (dict): Dictionary containing tender information
            
        Returns:
            dict: Dictionary with cleaned string values
        """
        cleanedData = {}
        
        for key, value in tenderInfo.items():
            if isinstance(value, str):
                cleanedData[key] = self.cleanText(value)
            else:
                cleanedData[key] = value
        
        return cleanedData
    
