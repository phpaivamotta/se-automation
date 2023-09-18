from create_assembly_folders import create_assembly_folders
from rev_manage import rev_manage
from rename_drafts_in_folders import rename_drafts_in_folders
from modify_properties import modify_properties

# create_assembly_folders
draft_folder_paths = [
    r"C:\Users\phpai\Desktop\Templates\S7\S7-8",
    r"C:\Users\phpai\Desktop\Templates\S7\S7-10",
    r"C:\Users\phpai\Desktop\Templates\S8\S8-8",
    r"C:\Users\phpai\Desktop\Templates\S8\S8-10",
]
job_folder_path = r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings"

created_folders, drafts_found = create_assembly_folders(
    draft_folder_paths, job_folder_path)

if isinstance(created_folders, list):
    print("\nCreated folders:")
    for folder in created_folders:
        print(f"  - {folder}")

    print("\nDraft files found:")
    for draft in drafts_found:
        print(f"  - {draft}")
else:
    print(created_folders)

# rev_manage
rev_manage(drafts_found, created_folders)

# rename_drafts_in_folders
renamed_files = rename_drafts_in_folders(created_folders)

# modify_properties
summary_information = {
    "Subject": "Kansas City, MO",
    # "Author": "",
}

document_summary_information = {
    "Manager": "JDE",
    "Company": "MAVERICK APPLIED SCIENCE, INC."
}

project_information = {
    "Document Number": "3031",
    "Revision": "C",
    "Project Name": "Brine Tank"
}

custom_information = {
    "CHECKED BY:": "JF",
    "DATE": "09/15/23",
    "COMPANY ADDRESS": "PALMETTO, FL",
    "DRAWN BY:": "PM",
    "APPROVED BY:": "DJM",
    "Customer P.O.:": "12356",
    "Customer W.O.:": "-",
    "Drawing Number:": "",
    # "CLIENT": "",
    "CLIENT ADDRESS": "MOBILE, AL",
    "MSS Job No:": "20233031",
}

modify_properties(renamed_files, summary_information,
                  document_summary_information, project_information, custom_information)
