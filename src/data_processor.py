import pandas as pd
import os
from datetime import datetime
from config.config import Config
import logging

class TimesheetProcessor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def process_raw_data(self, raw_data):
        """Process raw API data into a clean DataFrame"""
        rows = []
        
        for item in raw_data.get('items', []):
            row_data = {}
            for column, value in item.get('values', {}).items():
                # Extract the actual value from Coda's response format
                if isinstance(value, dict):
                    row_data[column] = value.get('name') or value.get('text') or str(value)
                else:
                    row_data[column] = value
            rows.append(row_data)
        
        df = pd.DataFrame(rows)
        self.logger.info(f"Processed {len(df)} rows with columns: {list(df.columns)}")
        return df
    
    def clean_timesheet_data(self, df):
        """Apply specific cleaning rules for timesheet data"""
        # Add your specific cleaning logic here
        # For example:
        # - Convert date columns to datetime
        # - Parse duration/hours columns
        # - Clean up project names, etc.
        
        cleaned_df = df.copy()
        
        # Example cleaning (adjust based on your actual columns):
        if 'Date' in cleaned_df.columns:
            cleaned_df['Date'] = pd.to_datetime(cleaned_df['Date'], errors='coerce')
        
        if 'Hours' in cleaned_df.columns:
            cleaned_df['Hours'] = pd.to_numeric(cleaned_df['Hours'], errors='coerce')
        
        self.logger.info("Applied data cleaning rules")
        return cleaned_df
    
    def export_to_csv(self, df, filename=None):
        """Export DataFrame to CSV"""
        os.makedirs(Config.PROCESSED_DATA_DIR, exist_ok=True)
        
        if filename is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"timesheet_processed_{timestamp}.csv"
        
        filepath = os.path.join(Config.PROCESSED_DATA_DIR, filename)
        df.to_csv(filepath, index=False)
        
        self.logger.info(f"Data exported to {filepath}")
        print(f"✅ Data exported to {filepath}")
        return filepath
    
    def generate_summary(self, df):
        """Generate summary statistics"""
        summary = {
            'total_rows': len(df),
            'date_range': None,
            'total_hours': None,
            'columns': list(df.columns)
        }
        
        if 'Date' in df.columns:
            dates = pd.to_datetime(df['Date'], errors='coerce').dropna()
            if len(dates) > 0:
                summary['date_range'] = f"{dates.min().date()} to {dates.max().date()}"
        
        if 'Hours' in df.columns:
            hours = pd.to_numeric(df['Hours'], errors='coerce').dropna()
            if len(hours) > 0:
                summary['total_hours'] = hours.sum()
        
        return summary