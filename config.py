# Configuration file for PDF to Markdown Converter

import os

# Flask Configuration
SECRET_KEY = 'your-secret-key-here'  # Change this in production!
DEBUG = True
HOST = '0.0.0.0'
PORT = 5000

# File Upload Configuration
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'pdf'}
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB max file size

# MinerU Configuration
MINERU_CONFIG = {
    'model_name': 'default',  # Use default MinerU model
    'device': 'auto',         # Auto-detect device (CPU/GPU)
    'batch_size': 1,          # Process one file at a time
}

# Excel Export Configuration
EXCEL_CONFIG = {
    'sheet_name': 'PDF_Results',
    'max_column_width': 100,  # Maximum column width in characters
    'auto_adjust_columns': True,
}

# Session Configuration
SESSION_CONFIG = {
    'permanent': False,
    'lifetime': 3600,  # 1 hour session lifetime
}

# Cleanup Configuration
CLEANUP_CONFIG = {
    'auto_cleanup': True,
    'cleanup_interval': 3600,  # Clean up old files every hour
    'max_file_age': 86400,     # Keep files for 24 hours max
}
