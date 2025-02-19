#%%


import extract_msg
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import argparse
import re

msg_file_path = "msg/2024-08-26_Freshwater-jellyfish_Matsuda_no-action.msg"

save_directory = "/output"

msg = extract_msg.openMsg(msg_file_path)

print('Sender: {}' .format(msg.sender))

print('Sent On: {}'.format(msg.date))

print('Subject: {}'.format(msg.subject))

print('Body: {}'.format(msg.body))
# %%

# open attachments

def saveAttachments(msgfile):
    with extract_msg.openMsg(msgfile) as msg:
        for attachment in msg.attachments:
            file_path = save_directory + attachment.longFilename
            
            attachment.save(customPath.msgfile)
            print(f'Saved attachment to {msgfile}')
        
    
#%%
msgbody = msg.body
#msgbody = msgbody.splitlines()

# Extract Date and Species using regular expressions

# Fields to extract
fields = ["Date", "Species", "Area", "Location", "Name", "Email", "Phone", "Comments"]

# Regex template
field_regex = {field: re.search(fr"{field}+(.+?)(?=\n\w|$)", msgbody, re.DOTALL) for field in fields}

# Special handling for Location to strip hyperlink
if field_regex["Location"]:
    location_text = re.sub(r"<.*?>", "", field_regex["Location"].group(1)).strip()
    field_regex["Location"] = location_text
else:
    location_text = "Location not found"

# Extracted values
results = {field: (location_text if field == "Location" else match.group(1).strip() if match else f"{field} not found") for field, match in field_regex.items()}

# Display results
for field, value in results.items():
    print(f"{field}: {value}")

""" date_match = re.search(r"Date\s+([\d-]+\s[\d:]+)", msgbody)
species_match = re.search(r"Species\s+([^\n]+)", msgbody)
area_match  = re.search(r"Area\s+([^\n]+)", msgbody)
location_match = re.search(r"Location\s+([^\n]+)", msgbody)
Name_match = re.search(r"Name\s+([^\n]+)", msgbody)
Email_match = re.search(r"Email\s+([^\n]+)", msgbody)
Phone_match = re.search(r"Phone\s+([^\n]+)", msgbody)
comments_match = re.search(r"Comments\s+(.+)", msgbody, re.DOTALL)

# Check and print extracted values
date = date_match.group(1) if date_match else "Date not found"
species = species_match.group(1) if species_match else "Species not found"
area = area_match.group(1) if area_match else "Area not found"
location = location_match.group(1) if area_match else "Location not found"

print(f"Date: {date}")
print(f"Species: {species}") """


# %%
