import json
import logging
from datetime import datetime
from typing import Dict, Any

class ConfigManager:
    """
    Manages configuration loading, validation, and access for the TenderScraper.
    
    This class handles loading configuration from JSON files, validating settings,
    and providing easy access to configuration values throughout the application.
    """
    
    def __init__(self, configPath: str = "config.json"):
        """
        Initialize the ConfigManager with a configuration file path.
        
        Args:
            configPath (str): Path to the JSON configuration file
        """
        self.configPath = configPath
        self.config = {}
        self.loadConfig()
    
    def loadConfig(self) -> None:
        """
        Load configuration from the JSON file and validate it.
        
        Raises:
            FileNotFoundError: If the configuration file doesn't exist
            ValueError: If the configuration is invalid
        """
        try:
            with open(self.configPath, 'r') as file:
                self.config = json.load(file)
            self.validateConfig()
            logging.info(f"Configuration loaded successfully from {self.configPath}")
        except FileNotFoundError:
            raise FileNotFoundError(f"Configuration file not found: {self.configPath}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in configuration file: {e}")
    
    def validateConfig(self) -> None:
        """
        Validate the loaded configuration for required fields and data types.
        
        Raises:
            ValueError: If required fields are missing or invalid
        """
        requiredSections = ['scraping', 'browser', 'timing', 'retry', 'output', 'logging']
        
        for section in requiredSections:
            if section not in self.config:
                raise ValueError(f"Missing required configuration section: {section}")
        
        # Validate date format
        try:
            dateFrom = datetime.strptime(self.config['scraping']['dateFrom'], "%Y-%m-%d")
            dateTo = datetime.strptime(self.config['scraping']['dateTo'], "%Y-%m-%d")
            
            if dateFrom > dateTo:
                raise ValueError("dateFrom cannot be later than dateTo")
        except KeyError as e:
            raise ValueError(f"Missing required date configuration: {e}")
        except ValueError as e:
            raise ValueError(f"Invalid date format in configuration: {e}")
        
        # Validate numeric values
        numericFields = [
            ('timing', 'pageLoadWait'),
            ('timing', 'modalRemovalWait'),
            ('timing', 'expandRowWait'),
            ('timing', 'collapseRowWait'),
            ('timing', 'nextPageWait'),
            ('timing', 'retryDelay'),
            ('retry', 'maxRetries'),
            ('retry', 'staleElementRetries')
        ]
        
        for section, field in numericFields:
            try:
                value = self.config[section][field]
                if not isinstance(value, (int, float)) or value < 0:
                    raise ValueError(f"Invalid value for {section}.{field}: must be non-negative number")
            except KeyError:
                raise ValueError(f"Missing required configuration: {section}.{field}")
    
    # These getter methods could be simplified to direct access in TenderScraper.py
    # For example: self.configManager.config.get('scraping', {}) instead of self.configManager.getScrapingConfig()
    # However, they are kept for possible future purposes
    def getScrapingConfig(self) -> Dict[str, Any]:
        """
        Get scraping configuration including dates and URL.
        
        Returns:
            Dict containing scraping configuration
        """
        return self.config.get('scraping', {})
    
    def getBrowserConfig(self) -> Dict[str, Any]:
        """
        Get browser configuration for Selenium setup.
        
        Returns:
            Dict containing browser configuration
        """
        return self.config.get('browser', {})
    
    def getTimingConfig(self) -> Dict[str, Any]:
        """
        Get timing configuration for various wait periods.
        
        Returns:
            Dict containing timing configuration
        """
        return self.config.get('timing', {})
    
    def getRetryConfig(self) -> Dict[str, Any]:
        """
        Get retry configuration for error handling.
        
        Returns:
            Dict containing retry configuration
        """
        return self.config.get('retry', {})
    
    def getOutputConfig(self) -> Dict[str, Any]:
        """
        Get output configuration for file naming and formats.
        
        Returns:
            Dict containing output configuration
        """
        return self.config.get('output', {})
    
    def getLoggingConfig(self) -> Dict[str, Any]:
        """
        Get logging configuration for application logging.
        
        Returns:
            Dict containing logging configuration
        """
        return self.config.get('logging', {})

    # This function is currently not in use but might be important for automation purpose
    def updateConfig(self, updates: Dict[str, Any]) -> None:
        """
        Update configuration with new values and save to file.
        
        Args:
            updates (Dict): Dictionary containing configuration updates
        """
        # Deep merge the updates into config
        for k, v in updates.items():
            if isinstance(v, dict):
                if k not in self.config:
                    self.config[k] = {}
                for sub_k, sub_v in v.items():
                    self.config[k][sub_k] = sub_v
            else:
                self.config[k] = v
        
        # Save to file
        try:
            with open(self.configPath, 'w') as file:
                json.dump(self.config, file, indent=4)
            logging.info(f"Configuration updated and saved to {self.configPath}")
        except Exception as e:
            logging.error(f"Failed to save configuration: {e}")
            raise
    
 