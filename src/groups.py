from pprint import pprint

from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.permissions.base_permissions import BasePermissions

from config import TEST_SITE_URL, client_credentials

# get client context with site url and client credentials
ctx = ClientContext(TEST_SITE_URL).with_credentials(client_credentials)

owners = ctx.web.associated_owner_group.get().execute_query()
visitors = ctx.web.associated_visitor_group.get().execute_query()
members = ctx.web.associated_member_group.get().execute_query()

root_folder = ctx.web.default_document_library().root_folder
folder_item = root_folder.list_item_all_fields
role_defs = ctx.web.role_definitions.get().execute_query()
for role_def in role_defs:
    assert isinstance(role_def.base_permissions, BasePermissions)
    print(role_def)
    pprint(role_def.base_permissions.permission_levels, indent=10)
