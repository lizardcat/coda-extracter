import pandas as pd
import os
import numpy as np
from datetime import datetime
from config.config import Config
import logging

class TimesheetProcessor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def process_raw_data(self, raw_data):
        """Process raw API data into a clean DataFrame with proper column names"""
        rows = []
        column_mapping = raw_data.get('column_mapping', {})
        
        for item in raw_data.get('items', []):
            row_data = {}
            for column_id, value in item.get('values', {}).items():
                # Use the display name if available, otherwise use the ID
                column_name = column_mapping.get(column_id, column_id)
                
                # Extract the actual value from Coda's response format
                if isinstance(value, dict):
                    # Handle different value types
                    if 'name' in value:
                        row_data[column_name] = value['name']
                    elif 'text' in value:
                        row_data[column_name] = value['text']
                    elif 'displayValue' in value:
                        row_data[column_name] = value['displayValue']
                    elif 'value' in value:
                        row_data[column_name] = value['value']
                    else:
                        row_data[column_name] = str(value)
                else:
                    row_data[column_name] = value
                    
            rows.append(row_data)
        
        df = pd.DataFrame(rows)
        self.logger.info(f"Processed {len(df)} rows with columns: {list(df.columns)}")
        return df
    
    def clean_timesheet_data(self, df):
        """Apply specific cleaning rules for timesheet data"""
        cleaned_df = df.copy()
        
        # Auto-detect and convert date columns
        date_keywords = ['date', 'day', 'when', 'created', 'modified', 'time']
        for col in cleaned_df.columns:
            if any(keyword in col.lower() for keyword in date_keywords):
                try:
                    cleaned_df[col] = pd.to_datetime(cleaned_df[col], errors='coerce')
                    self.logger.info(f"Converted {col} to datetime")
                except:
                    pass
        
        # Auto-detect and convert numeric columns (hours, duration, amounts)
        numeric_keywords = ['hour', 'time', 'duration', 'amount', 'cost', 'rate', 'total']
        for col in cleaned_df.columns:
            if any(keyword in col.lower() for keyword in numeric_keywords):
                try:
                    # Handle time formats like "2:30" (hours:minutes)
                    if cleaned_df[col].astype(str).str.contains(':').any():
                        cleaned_df[col] = cleaned_df[col].apply(self._convert_time_to_decimal)
                    else:
                        cleaned_df[col] = pd.to_numeric(cleaned_df[col], errors='coerce')
                    self.logger.info(f"Converted {col} to numeric")
                except:
                    pass
        
        # Clean text columns
        for col in cleaned_df.columns:
            if cleaned_df[col].dtype == 'object':
                try:
                    cleaned_df[col] = cleaned_df[col].astype(str).str.strip()
                except:
                    pass
        
        self.logger.info("Applied data cleaning rules")
        return cleaned_df
    
    def _convert_time_to_decimal(self, time_str):
        """Convert time format (e.g., '2:30') to decimal hours (e.g., 2.5)"""
        if pd.isna(time_str) or time_str == '':
            return np.nan
        
        try:
            time_str = str(time_str).strip()
            if ':' in time_str:
                parts = time_str.split(':')
                hours = float(parts[0])
                minutes = float(parts[1]) if len(parts) > 1 else 0
                return hours + (minutes / 60)
            else:
                return float(time_str)
        except:
            return np.nan
    
    def calculate_timesheet_metrics(self, df):
        """Calculate common timesheet metrics"""
        metrics = {}
        
        # Find hour columns
        hour_columns = []
        for col in df.columns:
            if any(keyword in col.lower() for keyword in ['hour', 'time', 'duration']) and pd.api.types.is_numeric_dtype(df[col]):
                hour_columns.append(col)
        
        # Find date columns
        date_columns = []
        for col in df.columns:
            if df[col].dtype == 'datetime64[ns]':
                date_columns.append(col)
        
        if hour_columns:
            primary_hour_col = hour_columns[0]
            metrics['total_hours'] = df[primary_hour_col].sum()
            metrics['average_daily_hours'] = df[primary_hour_col].mean()
            metrics['max_daily_hours'] = df[primary_hour_col].max()
            metrics['min_daily_hours'] = df[primary_hour_col].min()
            metrics['overtime_days'] = len(df[df[primary_hour_col] > 8])
            
            # Weekly totals if we can identify weeks
            if date_columns:
                primary_date_col = date_columns[0]
                df_with_week = df.copy()
                df_with_week['week'] = df_with_week[primary_date_col].dt.isocalendar().week
                weekly_hours = df_with_week.groupby('week')[primary_hour_col].sum()
                metrics['weekly_totals'] = weekly_hours.to_dict()
                metrics['average_weekly_hours'] = weekly_hours.mean()
        
        # Project breakdown if there's a project column
        project_columns = []
        for col in df.columns:
            if any(keyword in col.lower() for keyword in ['project', 'client', 'task', 'category']):
                project_columns.append(col)
        
        if project_columns and hour_columns:
            primary_project_col = project_columns[0]
            primary_hour_col = hour_columns[0]
            project_hours = df.groupby(primary_project_col)[primary_hour_col].sum()
            metrics['project_breakdown'] = project_hours.to_dict()
        
        return metrics
    
    def filter_data(self, df, filters):
        """
        Apply filters to the dataframe
        
        Args:
            df: DataFrame to filter
            filters: Dict with filter criteria
                    {'column': 'Hours', 'operator': '>', 'value': 8}
                    {'column': 'Project', 'operator': 'contains', 'value': 'Client A'}
        """
        filtered_df = df.copy()
        
        for filter_config in filters:
            column = filter_config.get('column')
            operator = filter_config.get('operator')
            value = filter_config.get('value')
            
            if column not in filtered_df.columns:
                continue
            
            if operator == '>':
                filtered_df = filtered_df[pd.to_numeric(filtered_df[column], errors='coerce') > float(value)]
            elif operator == '<':
                filtered_df = filtered_df[pd.to_numeric(filtered_df[column], errors='coerce') < float(value)]
            elif operator == '==':
                filtered_df = filtered_df[filtered_df[column].astype(str) == str(value)]
            elif operator == 'contains':
                filtered_df = filtered_df[filtered_df[column].astype(str).str.contains(str(value), case=False, na=False)]
            elif operator == 'date_range':
                if len(value) == 2:  # [start_date, end_date]
                    start_date, end_date = value
                    filtered_df = filtered_df[(filtered_df[column] >= start_date) & (filtered_df[column] <= end_date)]
        
        return filtered_df
    
    def aggregate_data(self, df, group_by_column, value_column, aggregation='sum'):
        """
        Aggregate data by a specific column
        
        Args:
            df: DataFrame to aggregate
            group_by_column: Column to group by
            value_column: Column to aggregate
            aggregation: Type of aggregation ('sum', 'mean', 'count', 'max', 'min')
        """
        if group_by_column not in df.columns or value_column not in df.columns:
            return pd.DataFrame()
        
        if aggregation == 'sum':
            result = df.groupby(group_by_column)[value_column].sum()
        elif aggregation == 'mean':
            result = df.groupby(group_by_column)[value_column].mean()
        elif aggregation == 'count':
            result = df.groupby(group_by_column)[value_column].count()
        elif aggregation == 'max':
            result = df.groupby(group_by_column)[value_column].max()
        elif aggregation == 'min':
            result = df.groupby(group_by_column)[value_column].min()
        else:
            result = df.groupby(group_by_column)[value_column].sum()
        
        return result.reset_index()
    
    def export_to_csv(self, df, filename=None):
        """Export DataFrame to CSV"""
        os.makedirs(Config.PROCESSED_DATA_DIR, exist_ok=True)
        
        if filename is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"timesheet_processed_{timestamp}.csv"
        
        filepath = os.path.join(Config.PROCESSED_DATA_DIR, filename)
        df.to_csv(filepath, index=False)
        
        self.logger.info(f"Data exported to {filepath}")
        return filepath
    
    def export_with_metrics(self, df, metrics, filename=None):
        """Export data with metrics summary"""
        os.makedirs(Config.PROCESSED_DATA_DIR, exist_ok=True)
        
        if filename is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            base_filename = f"timesheet_with_metrics_{timestamp}"
        else:
            base_filename = filename.replace('.csv', '').replace('.xlsx', '')
        
        # Export main data
        main_filepath = os.path.join(Config.PROCESSED_DATA_DIR, f"{base_filename}_data.csv")
        df.to_csv(main_filepath, index=False)
        
        # Export metrics as separate file
        metrics_filepath = os.path.join(Config.PROCESSED_DATA_DIR, f"{base_filename}_metrics.txt")
        with open(metrics_filepath, 'w') as f:
            f.write("TIMESHEET METRICS SUMMARY\n")
            f.write("=" * 30 + "\n\n")
            
            for key, value in metrics.items():
                if isinstance(value, dict):
                    f.write(f"{key.upper()}:\n")
                    for sub_key, sub_value in value.items():
                        f.write(f"  {sub_key}: {sub_value}\n")
                    f.write("\n")
                else:
                    f.write(f"{key}: {value}\n")
        
        self.logger.info(f"Data exported to {main_filepath}")
        self.logger.info(f"Metrics exported to {metrics_filepath}")
        
        return main_filepath, metrics_filepath
    
    def generate_summary(self, df):
        """Generate enhanced summary statistics"""
        summary = {
            'total_rows': len(df),
            'total_columns': len(df.columns),
            'columns': list(df.columns),
            'date_range': None,
            'total_hours': None,
            'null_percentages': {}
        }
        
        # Calculate null percentages for each column
        for col in df.columns:
            null_count = df[col].isnull().sum()
            summary['null_percentages'][col] = (null_count / len(df) * 100) if len(df) > 0 else 0
        
        # Find date columns and get date range
        for col in df.columns:
            if df[col].dtype == 'datetime64[ns]':
                dates = df[col].dropna()
                if len(dates) > 0:
                    summary['date_range'] = f"{dates.min().date()} to {dates.max().date()}"
                    summary['date_span_days'] = (dates.max() - dates.min()).days
                break
        
        # Find hour columns and calculate totals
        for col in df.columns:
            if any(keyword in col.lower() for keyword in ['hour', 'time', 'duration']) and pd.api.types.is_numeric_dtype(df[col]):
                hours = df[col].dropna()
                if len(hours) > 0:
                    summary['total_hours'] = hours.sum()
                    summary['average_hours'] = hours.mean()
                    summary['max_hours'] = hours.max()
                break
        
        return summary