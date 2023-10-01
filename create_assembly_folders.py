import os


def create_assembly_folders(draft_folder_paths, job_folder_path):
    """
    Create folders for the corresponding draft assembly files inside the job folder.

    Parameters:
    - draft_folder_paths (list): List of paths to folders containing the draft files.
    - job_folder_path (str): The full path to the job folder.

    Returns:
    - List of paths to the created folders or an alert if a path does not exist.
    - List of full paths to the draft files found inside the draft folders.
    """

    # Check if job folder path exists
    if not os.path.exists(job_folder_path):
        print(f"The job folder path '{job_folder_path}' does not exist.")
        return

    # Ensure "Working Drawings" folder exists
    if not job_folder_path.endswith("Working Drawings"):
        job_folder_path = os.path.join(job_folder_path, "Working Drawings")
        if not os.path.exists(job_folder_path):
            os.makedirs(job_folder_path)

    # Extract the job number from the folder name
    parent_folder_path = os.path.dirname(job_folder_path)
    job_folder_name = os.path.basename(parent_folder_path)
    job_number = job_folder_name[4:]

    # Create folders for each draft file inside the job folder
    created_folders = []
    draft_files_found = []
    drawing_number = 101  # Start from 0101

    for folder_path in draft_folder_paths:
        # Check if draft folder path exists
        if not os.path.exists(folder_path):
            return f"The draft folder path '{folder_path}' does not exist.", []

        # Search for the draft file in the folder
        draft_files = [f for f in os.listdir(folder_path) if f.endswith('.dft')]

        if len(draft_files) != 1:
            print(
                f"Unexpected number of draft files ({len(draft_files)}) in {folder_path}. Skipping...")
            continue

        draft_file = draft_files[0]
        draft_files_found.append(os.path.join(folder_path, draft_file))

        # Extract the name without the extension
        draft_name = os.path.splitext(draft_file)[0]

        # Create the folder name
        folder_name = f"{job_number}-{drawing_number:04}-{draft_name}"

        # Create the full path to the folder
        new_folder_path = os.path.join(job_folder_path, folder_name)

        # Create the folder
        if not os.path.exists(new_folder_path):
            os.makedirs(new_folder_path)
            created_folders.append(new_folder_path)

        # Increment the drawing number
        drawing_number += 1

    return created_folders, draft_files_found


# # Example usage
# draft_folder_paths = [
#     r"C:\Users\phpai\Desktop\Templates\S7\S7-8",
#     r"C:\Users\phpai\Desktop\Templates\S7\S7-10",
#     r"C:\Users\phpai\Desktop\Templates\S8\S8-8",
#     r"C:\Users\phpai\Desktop\Templates\S8\S8-10",
# ]
# job_folder_path = r"C:\Users\phpai\Desktop\Jobs\20233031\Working Drawings"

# created_folders, drafts_found = create_assembly_folders(
#     draft_folder_paths, job_folder_path)

# if isinstance(created_folders, list):
#     print("\nCreated folders:")
#     for folder in created_folders:
#         print(f"  - {folder}")

#     print("\nDraft files found:")
#     for draft in drafts_found:
#         print(f"  - {draft}")
# else:
#     print(created_folders)
