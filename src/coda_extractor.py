import requests
import json
import logging
import os
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
    
    def get_table_columns(self, doc_id, table_id):
        """Get column information for a table to map IDs to names"""
        try:
            response = requests.get(
                f"{self.base_url}/docs/{doc_id}/tables/{table_id}/columns",
                headers=self.headers
            )
            response.raise_for_status()
            
            columns_data = response.json()
            
            # Create mapping from column ID to display name
            column_mapping = {}
            for col in columns_data.get('items', []):
                column_mapping[col['id']] = col['name']
            
            self.logger.info(f"Retrieved {len(column_mapping)} column mappings")
            return column_mapping
            
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Error retrieving columns: {e}")
            raise
    
    def get_timesheet_data(self, doc_id=None, table_id=None, max_rows=None, selected_columns=None):
        """
        Extract timesheet data from specified table with pagination
        
        Args:
            doc_id: Document ID
            table_id: Table ID  
            max_rows: Maximum number of rows to retrieve (None for all)
            selected_columns: List of column names to extract (None for all)
        """
        doc_id = doc_id or Config.DOC_ID
        table_id = table_id or Config.TABLE_ID
        
        try:
            # First, get column mappings
            self.logger.info("Getting column mappings...")
            column_mapping = self.get_table_columns(doc_id, table_id)
            
            # Prepare pagination parameters
            all_rows = []
            page_token = None
            total_fetched = 0
            page_size = 500  # Coda's maximum page size
            
            while True:
                # Build request parameters
                params = {
                    'limit': min(page_size, max_rows - total_fetched if max_rows else page_size)
                }
                
                if page_token:
                    params['pageToken'] = page_token
                
                # Add column filtering if specified
                if selected_columns:
                    # Convert column names to IDs
                    reverse_mapping = {v: k for k, v in column_mapping.items()}
                    column_ids = [reverse_mapping.get(col_name) for col_name in selected_columns if col_name in reverse_mapping]
                    if column_ids:
                        params['columns'] = ','.join(column_ids)
                
                self.logger.info(f"Fetching page with {params['limit']} rows (total so far: {total_fetched})")
                
                response = requests.get(
                    f"{self.base_url}/docs/{doc_id}/tables/{table_id}/rows",
                    headers=self.headers,
                    params=params
                )
                response.raise_for_status()
                
                data = response.json()
                current_rows = data.get('items', [])
                
                if not current_rows:
                    break
                
                all_rows.extend(current_rows)
                total_fetched += len(current_rows)
                
                # Check if we've reached the maximum or if there are no more pages
                page_token = data.get('nextPageToken')
                if not page_token or (max_rows and total_fetched >= max_rows):
                    break
            
            # Combine all data
            combined_data = {
                'items': all_rows,
                'column_mapping': column_mapping
            }
            
            self.logger.info(f"Successfully extracted {len(all_rows)} total rows")
            
            # Save raw data
            self._save_raw_data(combined_data)
            return combined_data
            
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