import win32com.client
import os


def modify_properties(file_paths, summary_information, document_summary_information, project_information, custom_information):
    # Connect to Solid Edge using early binding
    application = win32com.client.gencache.EnsureDispatch(
        "SolidEdge.Application")

    # Make Solid Edge visible
    application.Visible = True

    for file_path in file_paths:
        # Open the draft document
        draft_document = application.Documents.Open(file_path)

        # # Print of properties for debugging
        # summary_properties_obj = draft_document.Properties("Custom")
        # print("Properties in Custom:")
        # for prop in summary_properties_obj:
        #     print(prop.Name)

        # Access and modify properties in the "SummaryInformation" section
        summary_properties_obj = draft_document.Properties(
            "SummaryInformation")
        for prop_name, value in summary_information.items():
            if prop_name in [prop.Name for prop in summary_properties_obj]:
                prop = summary_properties_obj.Item(prop_name)
                prop.Value = value
            else:
                print(
                    f"Property '{prop_name}' not found in SummaryInformation section.")

        # Access and modify properties in the "DocumentSummaryInformation" section
        document_summary_properties_obj = draft_document.Properties(
            "DocumentSummaryInformation")
        for prop_name, value in document_summary_information.items():
            if prop_name in [prop.Name for prop in document_summary_properties_obj]:
                prop = document_summary_properties_obj.Item(prop_name)
                prop.Value = value
            else:
                print(
                    f"Property '{prop_name}' not found in DocumentSummaryInformation section.")

        # Access and modify properties in the "ProjectInformation" section
        project_properties_obj = draft_document.Properties(
            "ProjectInformation")
        for prop_name, value in project_information.items():
            if prop_name in [prop.Name for prop in project_properties_obj]:
                prop = project_properties_obj.Item(prop_name)
                prop.Value = value
            else:
                print(
                    f"Property '{prop_name}' not found in ProjectInformation section.")

        # Access and modify properties in the "Custom" section
        custom_properties_obj = draft_document.Properties("Custom")
        for prop_name, value in custom_information.items():
            if prop_name in [prop.Name for prop in custom_properties_obj]:
                prop = custom_properties_obj.Item(prop_name)
                if prop_name == "Drawing Number:":
                    prop.Value = os.path.basename(file_path).split('-')[1]
                else:
                    prop.Value = value
            else:
                print(f"Property '{prop_name}' not found in Custom section.")

        # Save and close the draft document
        draft_document.Save()
        draft_document.Close()


# # Example usage
# summary_information = {
#     "Subject": "Kansas City, MO",
#     # "Author": "",
# }

# document_summary_information = {
#     "Manager": "JDE",
#     "Company": "MAVERICK APPLIED SCIENCE, INC."
# }

# project_information = {
#     "Document Number": "3031",
#     "Revision": "C",
#     "Project Name": "Brine Tank"
# }

# custom_information = {
#     "CHECKED BY:": "JF",
#     "DATE": "09/12/23",
#     "COMPANY ADDRESS": "PALMETTO, FL",
#     "DRAWN BY:": "PM",
#     "APPROVED BY:": "DJM",
#     "Customer P.O.:": "12356",
#     "Customer W.O.:": "-",
#     "Drawing Number:": "", #
#     # "CLIENT": "",
#     "CLIENT ADDRESS": "MOBILE, AL",
#     "MSS Job No:" : "20233031",
# }

# file_paths = [
#  r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0101-S7-8\3031-0101-S7-8.dft",
#  r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0102-S7-10\3031-0102-S7-10.dft",
#  r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0103-S8-8\3031-0103-S8-8.dft",
#  r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0104-S8-10\3031-0104-S8-10.dft",
# ]

# modify_properties(file_paths, summary_information,
#                   document_summary_information, project_information, custom_information)
