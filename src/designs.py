from pprint import pprint
from uuid import UUID

from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.sitedesigns.creation_info import \
    SiteDesignCreationInfo
from office365.sharepoint.sitescripts.utility import SiteScriptUtility

from config import EMAIL, TEST_SITE_URL, client_credentials

# get client context with site url and client credentials
ctx = ClientContext(TEST_SITE_URL).with_credentials(client_credentials)

# get and delete existing site scripts
result = SiteScriptUtility.get_site_scripts(ctx).execute_query()
for script in result.value:
    SiteScriptUtility.delete_site_script(ctx, script.Id).execute_query()

# create site script
site_script = {
    '$schema': 'schema.json',
    'actions': [
        {
            'verb': 'applyTheme',
            'themeName': 'Contoso Theme'
        },
        {
            'verb': 'addPrincipalToSPGroup',
            'principal': EMAIL,
            'group': 'Visitors'
        }
    ],
    'bindata': {},
    'version': 1
}
result = SiteScriptUtility.create_site_script(
    ctx, 'Test script', '', site_script
).execute_query()
script_id = result.value.Id
print(f'Site script created with ID {script_id}')

# get and delete existing site designs
result = SiteScriptUtility.get_site_designs(ctx).execute_query()
for design in result.value:
    SiteScriptUtility.delete_site_design(ctx, design.Id).execute_query()

# create site design
info = SiteDesignCreationInfo(
    title='Contoso customer tracking',
    description='Creates customer list and applies standard theme',
    site_script_ids=[UUID(script_id)],
    web_template='64'
)
result = SiteScriptUtility.create_site_design(ctx, info).execute_query()
design_id = result.value.Id
print(f'Site design created with ID {design_id}')

# add site design task
result = SiteScriptUtility.add_site_design_task(
    ctx, TEST_SITE_URL, design_id
).execute_query()
pprint(result.value.to_json())
