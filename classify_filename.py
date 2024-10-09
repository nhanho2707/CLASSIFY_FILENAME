import os
import shutil
import pandas as pd

# Load the Excel file
df = pd.read_excel('ID_LIST.xlsx')  # Replace with your actual file name

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    ID = str(row['InstanceID'])  # Ensure it's a string for matching
    folder_code = str(row['M2_Label'])  # Folder name for classification

    # Get the folder path based on the Code
    target_folder = os.path.join(r'./dest_folder', folder_code)  # Replace with your base path

    # Create the folder if it doesn't exist
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    # Search for the file that matches the ID
    for file_name in os.listdir(r'./images'):  # Replace with your files directory
        if ID in file_name:  # If the ID is part of the file name
            # Move the file to the appropriate folder
            source_path = os.path.join(r'./images', file_name)
            destination_path = os.path.join(target_folder, file_name)
            
            shutil.move(source_path, destination_path)
            print(f"Moved {file_name} to {target_folder}")
