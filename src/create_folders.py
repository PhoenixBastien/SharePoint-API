from pathlib import Path

import pandas as pd
from bigtree import Node, dataframe_to_tree
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.folders.folder import Folder

from config import TEST_SITE_URL, client_credentials


def read_file(path: Path) -> pd.DataFrame:
    # read csv or excel file
    if path.suffix == ".csv":
        df = pd.read_csv(path)
    elif path.suffix == ".xlsx":
        df = pd.read_excel(path, "Mandate", usecols="B:J")
        df = df.dropna(how="all", ignore_index=True)
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True).rename_axis(None, axis=1)


def create_folders(
    ctx: ClientContext,
    node: Node,
    parent_folder: Folder,
    prefix: str = "",
    is_last: bool = True,
) -> None:
    # create folder on SharePoint
    folder = parent_folder.add(node.name).execute_query()

    # print node with connector and update prefix for next level
    if node.is_root:
        print(node.name)
    else:
        connector = "└── " if is_last else "├── "
        print(f"{prefix}{connector}{node.name}")
        prefix += "    " if is_last else "│   "

    # iterate over children with index to identify last child
    child_count = len(node.children)
    for i, child in enumerate(node.children):
        is_last = i == child_count - 1
        create_folders(ctx, child, folder, prefix, is_last)


def hierarchical_fill(df: pd.DataFrame) -> pd.DataFrame:
    df_filled = df.copy()
    levels = df.columns.tolist()

    # fill top level unconditionally
    df_filled[levels[0]] = df_filled[levels[0]].ffill()

    # fill each subsequent level conditionally
    for i in range(1, len(levels)):
        parent_col, current_col = levels[i - 1], levels[i]
        last_value, last_parent = None, None

        for idx, row in df_filled.iterrows():
            parent, current = row[parent_col], row[current_col]

            if parent != last_parent:
                last_value = None

            if pd.notna(current):
                last_value = current
            elif parent == last_parent and last_value is not None:
                df_filled.at[idx, current_col] = last_value

            last_parent = parent

    return df_filled.fillna("")


# get client context with site url and client credentials
ctx = ClientContext(TEST_SITE_URL).with_credentials(client_credentials)

# read csv or excel file
path = Path("test/folder_structure.csv")
# path = Path('test/CIS - (NCSB-NSPD-PPP) - PASSENGER PROTECT PROGRAM.xlsx')
df = read_file(path)

# hierarchically fill df with full paths
df = hierarchical_fill(df)
paths = df.apply(lambda row: Path("/".join(row)).as_posix(), axis=1)
root = dataframe_to_tree(paths.to_frame())

# get root folder and create folders
root_folder = ctx.web.default_document_library().root_folder
create_folders(ctx, root, root_folder)
