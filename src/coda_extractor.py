import requests
import json
import logging
from datetime import datetime
from config.config import Config

class CodaTimesheetExtractor:
    def __init__(self):
        Config.validate_config()
        self.api_token = Config.CODA_API_TOKEN
        self.base_url = "https://coda.io/apis/v1"
        self.headers = {
            "Authorization": f"Bearer {self.api_token}",
            "Content-Type": "application/json"
        }
        
        # Set up logging
        logging.basicConfig(
            filename=f'{Config.LOGS_DIR}/extraction_{datetime.now().strftime("%Y%m%d")}.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
    
    def get_documents(self):
        """Get all documents you have access to"""
        try:
            response = requests.get(f"{self.base_url}/docs", headers=self.headers)
            response.raise_for_status()
            self.logger.info("Successfully retrieved documents list")
            return response.json()
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Error retrieving documents: {e}")
            raise
    
    def get_tables(self, doc_id):
        """Get all tables in a document"""
        try:
            response = requests.get(f"{self.base_url}/docs/{doc_id}/tables", headers=self.headers)
            response.raise_for_status()
            self.logger.info(f"Successfully retrieved tables for doc {doc_id}")
            return response.json()
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Error retrieving tables: {e}")
            raise
    
    def get_timesheet_data(self, doc_id=None, table_id=None):
        """Extract timesheet data from specified table"""
        doc_id = doc_id or Config.DOC_ID
        table_id = table_id or Config.TABLE_ID
        
        try:
            response = requests.get(
                f"{self.base_url}/docs/{doc_id}/tables/{table_id}/rows",
                headers=self.headers
            )
            response.raise_for_status()
            
            data = response.json()
            self.logger.info(f"Successfully extracted {len(data.get('items', []))} rows")
            
            # Save raw data
            self._save_raw_data(data)
            return data
            
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Error extracting timesheet data: {e}")
            raise
    
    def _save_raw_data(self, data):
        """Save raw API response as JSON"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{Config.RAW_DATA_DIR}/timesheet_raw_{timestamp}.json"
        
        os.makedirs(Config.RAW_DATA_DIR, exist_ok=True)
        
        with open(filename, 'w') as f:
            json.dump(data, f, indent=2)
        
        self.logger.info(f"Raw data saved to {filename}")