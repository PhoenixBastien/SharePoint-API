import os

from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientCredential

if load_dotenv():
    CLIENT_ID = os.getenv('CLIENT_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    TENANT_URL = os.getenv('TENANT_URL')
    TEAM_SITE_URL = os.getenv('TEAM_SITE_URL')
    EMAIL = os.getenv('EMAIL')
    client_credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
else:
    print('.env file not found')
    quit()