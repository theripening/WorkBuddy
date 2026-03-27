import os
import argparse

def search_in_files(folder_path, search_term, filename_filter=None):
    if not os.path.isdir(folder_path):
        print(f"Error: '{folder_path}' is not a valid directory.")
        return

    print(f"Searching for '{search_term}' in folder: {folder_path}")
    if filename_filter:
        print(f"Filtering files with: '{filename_filter}'")

    for root, _, files in os.walk(folder_path):
        for file in files:
            if filename_filter and filename_filter not in file:
                continue

            file_path = os.path.join(root, file)
            try:
                print(f'trying {file_path}')
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    for line_num, line in enumerate(f, start=1):
                        if search_term in line:
                            print(f"[{file_path}] Line {line_num}: {line.strip()}")
            except Exception as e:
                print(f"Could not read file '{file_path}': {e}")

def main():
    parser = argparse.ArgumentParser(description="Search for a term in files within a folder.")
    parser.add_argument("folder", help="Path to the folder to search")
    parser.add_argument("term", help="Search term or phrase")
    parser.add_argument("--filter", help="Optional filename filter (e.g., .txt, .log)", default=None)

    args = parser.parse_args()
    search_in_files(args.folder, args.term, args.filter)

if __name__ == "__main__":
    main()
