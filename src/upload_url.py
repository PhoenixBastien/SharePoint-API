from office365.sharepoint.client_context import ClientContext, ClientCredential

from config import client_id, client_secret, site_url

client_credentials = ClientCredential(client_id, client_secret)
ctx = ClientContext(site_url).with_credentials(client_credentials)

root_folder = ctx.web.default_document_library().root_folder
url_list = [
    'learn.microsoft.com',
    'www.youtube.com',
    'www.google.com',
    'en.wikipedia.org',
    'fr.wikipedia.org',
    'de.wikipedia.org'
]

for url in url_list:
    file_name = f'{url}.url'
    content = f'[InternetShortcut]\nURL=http://{url}'.encode('utf-8')
    root_folder.upload_file(file_name, content).execute_query()