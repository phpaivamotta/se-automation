import win32com.client
import os


def rev_manage(draft_files, created_folders):
    """
    Automate the process of saving draft files to new locations using Solid Edge's Design Manager.

    Parameters:
    - draft_files (list): List of full paths to the draft files to be managed.
    - created_folders (list): List of full paths to the folders where the drafts should be saved.

    Returns:
    - If successful, the function will not return anything since application.PerformAction() throws a type error
    """

    results = []

    # Ensure the number of draft files matches the number of created folders
    if len(draft_files) != len(created_folders):
        print("Mismatch between number of draft files and created folders.")
        return
    
    # Create an instance of the DesignManager application once
    application = win32com.client.Dispatch("DesignManager.Application")
    application.Visible = True

    # Define the action value (2 = SaveAllAs)
    action = 2

    for file_path, new_folder_path in zip(draft_files, created_folders):
        try:
            # Check if there are already files inside the new_folder_path
            if os.listdir(new_folder_path):
                print(f"Folder {new_folder_path} is not empty. Skipping...")
                results.append(None)
                continue

            # Open the file in DesignManager
            document = application.OpenFileInDesignManager(file_path)

            # Set action for all files
            application.SetActionForAllFiles(action, new_folder_path)

            # Print the paths for clarity
            print(f"\nOriginal draft: {file_path}")
            print(f"Saved to folder: {new_folder_path}")
            print("-" * 50)  # Separator for readability

            # Perform the action
            result = application.PerformAction()
            results.append(result)

        except TypeError as te:
            # Ignore the specific 'int' object is not callable exception
            if "'int' object is not callable" in str(te):
                pass
            else:
                print(f"Error processing {file_path}: {te}")
                results.append(None)
        except Exception as ex:
            print(f"Error processing {file_path}: {ex}")
            results.append(None)

    return results


# # Example usage
# draft_files = [
#     r"C:\Users\phpai\Desktop\Templates\S7\S7-8\S7-8.dft",
#     r"C:\Users\phpai\Desktop\Templates\S7\S7-10\S7-10.dft",
#     r"C:\Users\phpai\Desktop\Templates\S8\S8-8\S8-8.dft",
#     r"C:\Users\phpai\Desktop\Templates\S8\S8-10\S8-10.dft",
# ]
# created_folders = [
#     r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0101-S7-8",
#     r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0102-S7-10",
#     r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0103-S8-8",
#     r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0104-S8-10",
# ]

# results = rev_manage(draft_files, created_folders)
# for file, result in zip(draft_files, results):
#     print(f"Processing {file} resulted in: {result}")
