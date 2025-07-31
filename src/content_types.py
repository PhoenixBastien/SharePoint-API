import csv

from office365.sharepoint.client_context import ClientContext, ClientCredential
from office365.sharepoint.contenttypes.content_type import ContentType
from office365.sharepoint.fields.choice import FieldChoice
from office365.sharepoint.fields.collection import FieldCollection

from config import client_id, client_secret, content_type_hub_url

client_credentials = ClientCredential(client_id, client_secret)
ctx = ClientContext(content_type_hub_url).with_credentials(client_credentials)

with (
    open('out/Content Types + Site Columns.csv', 'w+', newline='') as all_file,
    open('out/Content Types.csv', 'w+', newline='') as content_type_file
):
    fieldnames = [
        'Name',
        'Description',
        'Category',
        'Parent',
        'Content Type ID',
        'Column Name',
        'Column Type',
        'Column Required'
    ]

    all_writer = csv.DictWriter(all_file, fieldnames=fieldnames)
    all_writer.writeheader()

    content_type_writer = csv.DictWriter(content_type_file, fieldnames=fieldnames[:5])
    content_type_writer.writeheader()

    content_types = (
        ctx.web.content_types.get()
        .filter('Group ne \'_Hidden\'')
        .order_by('Name').execute_query()
    )

    for content_type in content_types:
        assert isinstance(content_type.parent, ContentType)
        parent = content_type.parent.get().execute_query()

        row = {
            'Name': content_type.name,
            'Description': content_type.description,
            'Category': content_type.group,
            'Parent': parent.name,
            'Content Type ID': content_type.id
        }
        content_type_writer.writerow(row)

        assert isinstance(content_type.fields, FieldCollection)
        fields = (
            content_type.fields.get()
            .filter('Hidden eq false and TypeDisplayName ne \'Computed\'')
            .order_by('Title').execute_query()
        )
        
        for field in fields:
            row.update({
                'Column Name': field.title,
                'Column Type': field.type_display_name,
                'Column Required': field.properties['Required']
            })
            all_writer.writerow(row)
            row = {}

with open('out/Site Columns.csv', 'w+', newline='') as field_file:
    fieldnames = [
        'Name',
        'Type',
        'Required',
        'Choices'
    ]
    field_writer = csv.DictWriter(field_file, fieldnames=fieldnames)
    field_writer.writeheader()

    fields = (
        ctx.web.fields.get()
        .filter('Hidden eq false and TypeDisplayName ne \'Computed\'')
        .order_by('Title').execute_query()
    )

    for field in fields:
        field_writer.writerow({
            'Name': field.title,
            'Type': field.type_display_name,
            'Required': field.properties['Required'],
            'Choices': '; '.join(field.choices)
            if isinstance(field, FieldChoice) else ''
        })