import pickle
from set_properties import set_properties

# Load the cached renamed_files
with open('renamed_files_cache.pkl', 'rb') as cache_file:
    renamed_files = pickle.load(cache_file)

# properties to be updated
summary_information = {
    "Subject": "Kansas City, MO",
    # "Author": "",
}

document_summary_information = {
    "Manager": "JDE",
    "Company": "VAMSHI APPLIED TANKS, INC."
}

project_information = {
    "Document Number": "3031",
    "Revision": "C",
    "Project Name": "Brine Tank"
}

custom_information = {
    "CHECKED BY:": "JF",
    "DATE": "09/18/23",
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

set_properties(renamed_files, summary_information,
                  document_summary_information, project_information, custom_information)