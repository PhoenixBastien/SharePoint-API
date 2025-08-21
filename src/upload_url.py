from office365.sharepoint.client_context import ClientContext

from config import TEAM_SITE_URL, client_credentials

# get client context with site url and client credentials
ctx = ClientContext(TEAM_SITE_URL).with_credentials(client_credentials)

root_folder = ctx.web.default_document_library().root_folder
links = [
    'learn.microsoft.com',
    'www.youtube.com',
    'www.google.com',
    'en.wikipedia.org',
    'fr.wikipedia.org',
    'de.wikipedia.org'
]

for link in links:
    root_folder.upload_file(
        file_name=f'{link}.url',
        content=f'[InternetShortcut]\nURL=https://{link}'.encode('utf-8')
    ).execute_query()
