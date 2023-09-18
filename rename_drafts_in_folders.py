import os


def rename_drafts_in_folders(created_folders):
    """
    Rename the draft files in the destination folders to have the same name as the folder.

    Parameters:
    - created_folders (list): List of full paths to the folders where the drafts were saved.

    Returns:
    - List of new paths to the renamed draft files.
    """

    renamed_files = []

    for folder_path in created_folders:
        # List all files in the folder
        files_in_folder = os.listdir(folder_path)

        # Check if the folder is empty
        if not files_in_folder:
            print(f"Folder {folder_path} is empty. Skipping...")
            continue

        # Filter out only the draft files
        draft_files = [f for f in files_in_folder if f.endswith('.dft')]

        if not draft_files:
            print(f"No draft files found in {folder_path}. Skipping...")
            continue
        elif len(draft_files) > 1:
            print(
                f"Unexpected number of draft files in {folder_path}. Skipping...")
            continue

        old_file_name = draft_files[0]
        old_file_path = os.path.join(folder_path, old_file_name)

        # Extract the desired new name from the folder name
        new_file_name = os.path.basename(folder_path) + ".dft"
        new_file_path = os.path.join(folder_path, new_file_name)

        # Rename the file
        os.rename(old_file_path, new_file_path)
        renamed_files.append(new_file_path)

        # Print the renaming action for clarity
        print(
            f"Renamed:\n  From: {old_file_name}\n  To:   {new_file_name}\n  In:   {folder_path}\n{'-' * 60}")

    return renamed_files


# # Example usage
# created_folders = [
#     r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0101-S7-8",
#     r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0102-S7-10",
#     r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0103-S8-8",
#     r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings\3031-0104-S8-10",
# ]

# renamed_files = rename_drafts_in_folders(created_folders)
