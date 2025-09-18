import os
import sys

from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientCredential

if not load_dotenv():
    sys.exit('.env file not found')

CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT = os.getenv('TENANT')
ROOT_SITE_URL = os.getenv('ROOT_SITE_URL')
ADMIN_SITE_URL = os.getenv('ADMIN_SITE_URL')
TEST_SITE_URL = os.getenv('TEST_SITE_URL')
EMAIL = os.getenv('EMAIL')

client_credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)