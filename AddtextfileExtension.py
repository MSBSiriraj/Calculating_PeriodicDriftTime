import os

folder_path = os.path.dirname(os.path.abspath(__file__))
prefix = "Lipid_"

# Get a list of all files in the folder
files = os.listdir(folder_path)

for file in files:
    if file.startswith(prefix) and '.' not in file:
        # Construct the new filename with the '.txt' extension
        new_name = file + ".txt"

        # Build the full paths for the old and new filenames
        old_path = os.path.join(folder_path, file)
        new_path = os.path.join(folder_path, new_name)

        # Rename the file
        os.rename(old_path, new_path)

print("Files renamed successfully!")
