import os
import win32com.client



def reorder_supports(reordered_supports_dirs):
    '''
    Reorder supports
    '''

    drawing_number = 101

    try:
        application = win32com.client.Dispatch("SolidEdge.Application")
        application.Visible = True
    except Exception as e:
        print(f"Error trying to connect to Solid Edge: {e}")

    for support_dir in reordered_supports_dirs:
        # Check support folder path exists
        if not os.path.exists(support_dir):
            print(f"The support folder path '{support_dir}' does not exist.")
            return
        
        base_name = os.path.basename(support_dir)
        chunks = base_name.split('-')
        
        new_folder_name = f"{chunks[0]}-{drawing_number:04}-{chunks[2]}-{chunks[3]}"
        new_folder_path = os.path.join(os.path.dirname(support_dir), new_folder_name)

        if support_dir == new_folder_path:
            drawing_number += 1
            continue
        
        os.rename(support_dir, new_folder_path)

        # Search for the draft file in the folder
        draft_file = [f for f in os.listdir(new_folder_path) if f.endswith('.dft')]

        if len(draft_file) != 1:
            print(
                f"Unexpected number of draft files ({len(draft_file)}) in {new_folder_path}. Skipping...")
            continue

        draft_file = draft_file[0]
        old_file_path = os.path.join(new_folder_path, draft_file)

        new_file_path = os.path.join(new_folder_path, os.path.basename(new_folder_path)) + '.dft'

        os.rename(old_file_path, new_file_path)

        try:
            draft = application.Documents.Open(new_file_path)
        except Exception as e:
            print(f"Error trying to open file '{new_file_path}' in Solid Edge: {e}")

        draft.Properties("Custom").Item("Drawing Number:").Value = f"{drawing_number:04}"

        drawing_number += 1

#example usage
reordered_supports_dirs = [
    r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0101-S7-8",
    r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0102-S7-10",
    r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0101-S8-8",
    r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0102-S8-10",
    ]
reorder_supports(reordered_supports_dirs)
