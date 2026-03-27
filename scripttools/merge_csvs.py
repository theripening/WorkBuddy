#USAGE
# python merge_csvs.py C:\\TEMP\\SRG_auths --filter authorization
import os
import pandas as pd

def merge_csvs_in_folder(folder, output_filename="merged.csv", filter_str=None):
    """
    Merges all CSV files in a folder into a single output file.
    Optionally filters files by a substring in the filename.

    Args:
        folder (str): Path to the folder containing CSV files.
        output_filename (str): Name of the merged output CSV file.
        filter_str (str, optional): Substring to filter filenames.
    """
    merged_df = pd.DataFrame()

    for file in os.listdir(folder):
        if file.lower().endswith('.txt') and (filter_str is None or filter_str in file):
            file_path = os.path.join(folder, file)
            print(f"Merging: {file_path}")
            df = pd.read_csv(file_path)
            merged_df = pd.concat([merged_df, df], ignore_index=True)

    output_path = os.path.join(folder, output_filename)
    merged_df.to_csv(output_path, index=False)
    print(f"✅ Merged file saved as: {output_path}")

# Command-line interface
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Merge CSV files in a folder into one.")
    parser.add_argument("folder", help="Path to the folder containing CSV files")
    parser.add_argument("--output", help="Name of the output merged CSV file", default="merged.csv")
    parser.add_argument("--filter", help="Substring to filter filenames", default=None)

    args = parser.parse_args()
    merge_csvs_in_folder(args.folder, args.output, args.filter)
