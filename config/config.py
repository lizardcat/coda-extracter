import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    CODA_API_TOKEN = os.getenv('CODA_API_TOKEN')
    DOC_ID = os.getenv('CODA_DOC_ID')
    TABLE_ID = os.getenv('CODA_TABLE_ID')
    
    # File paths
    DATA_DIR = 'data'
    RAW_DATA_DIR = os.path.join(DATA_DIR, 'raw')
    PROCESSED_DATA_DIR = os.path.join(DATA_DIR, 'processed')
    LOGS_DIR = 'logs'
    
    @classmethod
    def validate_config(cls):
        missing = []
        if not cls.CODA_API_TOKEN:
            missing.append('CODA_API_TOKEN')
        if not cls.DOC_ID:
            missing.append('CODA_DOC_ID')
        if not cls.TABLE_ID:
            missing.append('CODA_TABLE_ID')
        
        if missing:
            raise ValueError(f"Missing required environment variables: {', '.join(missing)}")