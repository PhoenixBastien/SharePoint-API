import os

from dotenv import load_dotenv

load_dotenv()

client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
site_url = os.getenv('SITE_URL')
content_type_hub_url = os.getenv('ROOT_SITE_URL') + '/sites/ContentTypeHub'