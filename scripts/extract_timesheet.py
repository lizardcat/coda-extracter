#!/usr/bin/env python3
"""
Main script to extract and process timesheet data from Coda
"""

import os
import sys
import argparse
from datetime import datetime

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.coda_extractor import CodaTimesheetExtractor
from src.data_processor import TimesheetProcessor

def main():
    parser = argparse.ArgumentParser(description='Extract timesheet data from Coda')
    parser.add_argument('--output', '-o', help='Output filename (optional)')
    parser.add_argument('--list-docs', action='store_true', help='List available documents')
    parser.add_argument('--list-tables', help='List tables in specified document ID')
    
    args = parser.parse_args()
    
    # Create necessary directories
    os.makedirs('logs', exist_ok=True)
    os.makedirs('data/raw', exist_ok=True)
    os.makedirs('data/processed', exist_ok=True)
    
    try:
        extractor = CodaTimesheetExtractor()
        processor = TimesheetProcessor()
        
        # Handle list operations
        if args.list_docs:
            docs = extractor.get_documents()
            print("Available documents:")
            for doc in docs.get('items', []):
                print(f"  {doc['id']}: {doc['name']}")
            return
        
        if args.list_tables:
            tables = extractor.get_tables(args.list_tables)
            print(f"Tables in document {args.list_tables}:")
            for table in tables.get('items', []):
                print(f"  {table['id']}: {table['name']}")
            return
        
        # Extract and process data
        print("🔄 Extracting timesheet data from Coda...")
        raw_data = extractor.get_timesheet_data()
        
        print("🔄 Processing data...")
        df = processor.process_raw_data(raw_data)
        df_cleaned = processor.clean_timesheet_data(df)
        
        # Generate summary
        summary = processor.generate_summary(df_cleaned)
        print("\n📊 Summary:")
        print(f"  Total rows: {summary['total_rows']}")
        print(f"  Columns: {', '.join(summary['columns'])}")
        if summary['date_range']:
            print(f"  Date range: {summary['date_range']}")
        if summary['total_hours']:
            print(f"  Total hours: {summary['total_hours']}")
        
        # Export data
        output_file = processor.export_to_csv(df_cleaned, args.output)
        
        print(f"\n✅ Extraction complete!")
        print(f"📁 Processed data saved to: {output_file}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())