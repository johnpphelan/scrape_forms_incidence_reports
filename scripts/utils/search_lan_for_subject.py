#%%
import os
import shutil
import extract_msg

folder_path = "//SFP.IDIR.BCGOV/S140/S40203/RSD_ FISH & AQUATIC HABITAT BRANCH/General/2 SCIENCE - Invasives/SPECIES/5_Incidental Observations/Aquatic Reports/Aquatic Reports pending to be processed/Unsorted"
target_subject = "RI-And Smartphone App Report Submission"
destination_folder = "./messages/"
count = 0

# Iterate over all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".msg"):  # Only process .msg files
        file_path = os.path.join(folder_path, filename)
        
        try:
            # Extract content from the .msg file
            msg = extract_msg.Message(file_path)
            msg_subject = msg.subject  # Get the subject of the email
            
            # Check if the subject matches the target subject
            if msg_subject == target_subject:
                count += 1  # Increment the counter if it matches
                destination_path = os.path.join(destination_folder, filename)
                shutil.copy(file_path, destination_path)  # Copy the file
        except Exception as e:
            continue

# Output the count
print(f"Number of .msg files with subject '{target_subject}': {count}")



# %%
