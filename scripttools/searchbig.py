import os
import shutil

def find_files_recurse(folder_path, partial_name):
    matches = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if partial_name.lower() in file.lower():
                matches.append(os.path.join(root, file))
    return matches

def find_files_by_partial_name(folder_path, partial_name):
    matches = []
    try:
        with os.scandir(folder_path) as entries:
            for entry in entries:
                if entry.is_file() and partial_name.lower() in entry.name.lower():
                    matches.append(entry.path)
    except FileNotFoundError:
        print(f"Folder not found: {folder_path}")
    return matches



def confirm_and_copy(files, destination):
    if not files:
        print("No matching files found.")
        return
    for file in files:
        print(file)
    confirm = input(f"\nDo you want to copy these files to {destination}? (yes/no): ").strip().lower()
    if confirm == 'yes':
        os.makedirs(destination, exist_ok=True)
        for file in files:
            shutil.copy(file, destination)
        print(f"\nFiles successfully copied to {destination}")
    else:
        print("\nCopy canceled.")

# Example usage
folder = r'J:\RecvArchive'  # r'J:\RecvArchive' r'S:\Recv' 
    
search_term = 'FW4.0226.SUPP.CSV'                  # Change this
destination_folder = r'C:\Temp'

matches = find_files_by_partial_name(folder, search_term)
confirm_and_copy(matches, destination_folder)
