from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.permissions.base_permissions import BasePermissions
from office365.sharepoint.permissions.kind import PermissionKind
from office365.sharepoint.principal.users.collection import UserCollection

from config import EMAIL, SITE_URL, client_credentials

# get client context with site url and client credentials
ctx = ClientContext(SITE_URL).with_credentials(client_credentials)

# check if new group name already exists
group_name = "Alert Manager Group"
groups = ctx.web.site_groups
result = groups.filter(f"Title eq '{group_name}'").get().execute_query()

# delete group if name already exists
if result:
    groups.remove_by_id(group_id=result[0].id).execute_query()

# add new group
group = groups.add(group_name).execute_query()
print(f"{group} created")

# check if new role definiton name already exits
role_name = "Alert Manager Role"
role_defs = ctx.web.role_definitions
result = role_defs.filter(f"Name eq '{role_name}'").get().execute_query()

# delete role definition if name already exists
if result:
    role_defs[0].delete_object().execute_query()

# set base permissions for new role definition
perms = BasePermissions()
perms.set(PermissionKind.CreateAlerts)
perms.set(PermissionKind.ManageAlerts)
perms.set(PermissionKind.ApplyStyleSheets)

# BasePermissions' Low and High variables are swapped
# issue: https://github.com/vgrem/office365-rest-python-client/issues/959
if perms.to_json() == {"Low": perms.High, "High": perms.Low}:
    perms.Low, perms.High = perms.High, perms.Low

# add new role definition
role_def = role_defs.add(perms, role_name).execute_query()
print(f"{role_name} created with permission levels: {perms.permission_levels}")

# get folder by relative url
folder = (
    ctx.web.get_folder_by_server_relative_url(
        "Shared Documents/LINES OF BUSINESS - (NCSB-NSPD-PPP)"
    )
    .get()
    .execute_query()
)
folder_item = folder.list_item_all_fields
print(f"{folder} found")

# ensure user is a member of site by login name and add if not
user = ctx.web.ensure_user(f"i:0#.f|membership|{EMAIL}").execute_query()
print(f"{user} is a member of {SITE_URL}")

# add user to group
assert isinstance(group.users, UserCollection)
group.users.add_user(user).execute_query()
print(f"{user} added to {group}")

# break role inheritance on folder
folder_item.break_role_inheritance().execute_query()
print(f"Role inheritance broken on {folder}")

# add group to role assignment on folder
folder_item.role_assignments.add_role_assignment(group.id, role_def.id).execute_query()
print(f"{role_def} assigned to {group} on {folder}")

# # get user permissions on folder
# result = folder_item.get_user_effective_permissions(user).execute_query()
# print(f'{user} has permission levels: {result.value.permission_levels}')
