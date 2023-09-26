import win32com.client
import os



def pdf_print(working_drawings_folder):
    '''
    PDF Printing function
    '''

    dft_files = []  # List to store paths of .dft files

    # Loop through the directory tree
    for dirpath, dirnames, filenames in os.walk(working_drawings_folder):

        # Skip checks for the main directory itself
        if dirpath == working_drawings_folder:
            continue

        # Check if the directory is empty
        if not dirnames and not filenames:
            print(f"⚠️ WARNING: Directory '{dirpath}' is empty! Operation could not be performed.")
            continue  # Skip further processing for this directory

        # Count .dft files in the current directory
        count_dft = sum(1 for filename in filenames if filename.endswith('.dft'))

        # If no .dft files exist in the sub-directory
        if count_dft == 0:
            print(f"⚠️ WARNING: No .dft files found in '{dirpath}'. Operation will be skipped.")
            continue  # Skip further processing for this directory

        # If more than two .dft files exist in the sub-directory
        if count_dft > 2:
            print(f"⚠️ WARNING: There are more than two .dft files in '{dirpath}'. Operation will be skipped.")
            continue  # Skip further processing for this directory

        # Otherwise, append .dft files to the list
        for filename in filenames:
            if filename.endswith('.dft'):
                dft_files.append(os.path.join(dirpath, filename))

    # Connect to Solid Edge
    try:
        application = win32com.client.Dispatch("SolidEdge.Application")
        application.Visible = True
        # SolidEdgeFramework = win32com.client.Dispatch("SolidEdgeFramework.Application")
        SolidEdgeConstants = win32com.client.Dispatch("SolidEdge.ConstantsModule")

        # Set the global setting to save all sheets to PDF
        # application.SetGlobalParameter(
        #     SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSheetOptions,
        #     SolidEdgeConstants.DraftSaveAsPDFSheetOptionsConstants.seDraftSaveAsPDFSheetOptionsConstantsAllSheets
        # )
    except Exception as e:
        print(f"Error trying to connect to Solid Edge: {e}")

    # Loop through each .dft file and convert it
    for dft_file in dft_files:
        # Compute the new file path in "Fab Latest Revision" folder
        base_filename = os.path.basename(dft_file)
        pdf_filename = os.path.splitext(base_filename)[0] + '.pdf'
        
        # Construct the output directory: ".../Submittal/Fab Latest Revision/"
        output_dir = os.path.join(os.path.dirname(working_drawings_folder), "Submittal", "Fab Latest Revision")
        output_file = os.path.join(output_dir, pdf_filename)

        # Print the intended output path (for debugging purposes)
        print(f"Attempting to save: {output_file}")

        # Create output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Open the .dft file
        document = application.Documents.Open(dft_file)
        
        # Save as PDF
        document.SaveAs(output_file, 0x27)


# Example usage
working_drawings_folder = r'C:\Users\pmotto\OneDrive - Maverick Applied Science\Desktop\Jobs\20233031\Working Drawings'
pdf_print(working_drawings_folder)