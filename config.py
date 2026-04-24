import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # Google Sheets Configuration (optional - can be provided via API)
    GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE')
    GOOGLE_CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_FILE')
    
    # OAuth2 Configuration (for browser-based authentication)
    GOOGLE_CLIENT_ID = os.getenv('GOOGLE_CLIENT_ID')
    GOOGLE_CLIENT_SECRET = os.getenv('GOOGLE_CLIENT_SECRET')
    GOOGLE_REDIRECT_URI = os.getenv('GOOGLE_REDIRECT_URI', 'http://localhost:5001/api/sheets/oauth/callback')
    
    # MyAccount Credentials (optional - can be provided via API)
    MYACCOUNT_USERNAME = os.getenv('MYACCOUNT_USERNAME')
    MYACCOUNT_PASSWORD = os.getenv('MYACCOUNT_PASSWORD')
    
    # Server Configuration
    FLASK_ENV = os.getenv('FLASK_ENV', 'development')
    FLASK_DEBUG = os.getenv('FLASK_DEBUG', 'True').lower() == 'true'
    PORT = int(os.getenv('PORT', 5001))
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')

    # Selenium: set to "true" on servers where no display is available.
    # Defaults to headless in production, visible in development.
    SELENIUM_HEADLESS = os.getenv(
        'SELENIUM_HEADLESS',
        'true' if FLASK_ENV == 'production' else 'false'
    ).lower() == 'true'

