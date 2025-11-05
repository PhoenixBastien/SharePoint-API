from pprint import pprint

from office365.sharepoint.client_context import ClientContext

from config import ADMIN_SITE_URL, ROOT_SITE_URL, client_credentials

# get client context with site url and client credentials
ctx = ClientContext(ADMIN_SITE_URL).with_credentials(client_credentials)

result = ctx.tenant.get_site_properties_from_sharepoint_by_filters().execute_query()

for props in result:
    if props.url.startswith(f"{ROOT_SITE_URL}/sites/"):
        pprint(props.properties)
