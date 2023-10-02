import win32com.client
import os



def pdf_print_all(templates_dir_path, output_dir_path):
    '''
    This function saves the drafts in the assembly subfolders inside of the Working Drawings folder as corresponding PDF files in the Submittal\Fab Latest Revision
    and Submittal\GA Latest Revision. All draft sheets are saved in a corresponding PDF file in the Fab Latest Revision folder, while
    only the two first draft sheets are saved in a corresponding PDF file in the GA Latest Revision folder.
    '''

    folders_not_included = [
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\2010 Catalog Supports (Old Catalog Models)',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\2010 In House Catalog & Guidelines',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\2013 In House Catalog & Guidelines',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\2014 SCS Support Templates',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\D4G',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HP17',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HP18',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HP19',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HP20',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HS9B',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HS12B',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HS13B',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HS14',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\HS15B',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\NOTES & MISC',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\P8LL',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\P9LL',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\P13-P14',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\S3(special)',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\S12 Scallops',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\S12B Scallops',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\S13 Scallops',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\S13B Scallops',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\S14 Scallops',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\XL12',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\XL13',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\XL14',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\S2\S2WP_BO',
        r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates\S2\S2-FRP',
    ]

    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir_path):
        return print(f'Output directory {output_dir_path} does NOT exist. Aborting operation.')

    # List to store paths of .dft files
    dft_files = []

    # Loop through the directory tree
    for dirpath, dirnames, filenames in os.walk(templates_dir_path):

        # Skip main directory itself
        if dirpath == templates_dir_path:
            continue

        # Skip unwanted folders and all their subfolders
        if any(dirpath.startswith(excluded_folder) for excluded_folder in folders_not_included):
            del dirnames[:]  # this line will prevent os.walk from traversing into the subdirectories
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
        # All PDF pages
        application.SetGlobalParameter(172, 1)
        # # Fist 2 PDF pages
        # application.SetGlobalParameter(172, 2)
        # application.SetGlobalParameter(173, "1,2")
    except Exception as e:
        print(f"Error trying to connect to Solid Edge: {e}")

    # Loop through each .dft file and convert it to pdf
    for dft_file in dft_files:
        # Compute the new file path in folder
        base_filename = os.path.basename(dft_file)
        pdf_filename = os.path.splitext(base_filename)[0] + '.pdf'
        
        # Construct the output directory
        output_file = os.path.join(output_dir_path, pdf_filename)

        # Print the destination output path
        print(f"Saving: {output_file}")

        # Open the .dft file
        document = application.Documents.Open(dft_file)

        # Save as PDF (Note: 0x27 stands for PDF)
        document.SaveAs(output_file, 0x27)
        document.Close()



# Example usage
templates_dir_path = r'F:\MAS\Solid Edge Parts Library\Pipe Support Templates'
output_dir_path = r'F:\MAS\Pipe Supports\Quotes\2020- Support Prices\9-29-23 PDF for Pricing - Auto'

pdf_print_all(templates_dir_path, output_dir_path)