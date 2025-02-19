#%%


import extract_msg
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import argparse
import re
import requests
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
import scripts.utils.android_defs as msgFuncs
from collections import defaultdict 

msg_file_path = "msg/2024-08-26_Freshwater-jellyfish_Matsuda_no-action.msg"

lan_record_loc = "//SFP.IDIR.BCGOV/S140/S40203/RSD_ FISH & AQUATIC HABITAT BRANCH/General/2 SCIENCE - Invasives/SPECIES/5_Incidental Observations/"

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
""" msgbody = msg.body


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


# Extract date with only the day portion
if field_regex["Date"]:
  date_match = field_regex["Date"].group(1).strip()
  # Use datetime to extract date only (YYYY-MM-DD)
  from datetime import datetime
  date_object = datetime.strptime(date_match, "%Y-%m-%d %H:%M:%S")  # Assuming format is YYYY-MM-DD HH:MM:SS
  date_only = date_object.strftime("%Y-%m-%d")
  field_regex["Date"] = date_only
else:
  field_regex["Date"] = "Date not found" """


 # Use defaultdict for handling missing keys

def parse_email_body(msgbody):
  """Parses an email body to extract specific fields using regular expressions.

  Args:
      msgbody: The email body content as a string.

  Returns:
      A dictionary containing extracted field values, with default messages
      for missing fields. Keys are field names, and values are the extracted data
      or default messages.
  """

  # Fields to extract
  fields = ["Date", "Species", "Area", "Location", "Name", "Email", "Phone", "Comments"]

  # Regular expression template
  field_regex = {field: re.search(fr"{field}+(.+?)(?=\n\w|$)", msgbody, re.DOTALL) for field in fields}

  # Special handling for Location to strip hyperlink
  if field_regex["Location"]:
      location_text = re.sub(r"<.*?>", "", field_regex["Location"].group(1)).strip()
      field_regex["Location"] = location_text
  else:
      location_text = "Location not found"  # Default for missing Location

  # Extract values and handle missing fields using defaultdict
  results = defaultdict(lambda field: f"{field} not found")
  for field, match in field_regex.items():
      if match:
          value = location_text if field == "Location" else match.group(1).strip()
      else:
          # No match, use default value
          pass  # defaultdict takes care of this

      results[field] = value

  return results

# Example usage (assuming you have the email body in msgbody)
extracted_data = parse_email_body(msgbody)
for field, value in extracted_data.items():
  print(f"{field}: {value}")




# %%
def extract_last_url(text):
  """Extracts the last URL from a given text.

  Args:
    text: The text to search for URLs.

  Returns:
    The last URL found in the text, or None if no URL is found.
  """

  url_pattern = r'https?://\S+'
  urls = re.findall(url_pattern, text)
  return urls[-1] if urls else None

def download_image(url, filename):
  """Downloads an image from the given URL and saves it to a file.

  Args:
    url: The URL of the image.
    filename: The desired filename for the image.
  """

  response = requests.get(url)
  response.raise_for_status()  # Raise an exception for error HTTP statuses

  with open(filename, 'wb') as f:
    f.write(response.content)

def save_url_to_file(url, filename):
  """Saves the given URL to a text file.

  Args:
    url: The URL to be saved.
    filename: The desired filename for the text file.
  """

  with open(filename, 'w') as f:
    f.write(url)
    
last_url = extract_last_url(msgbody)

date_string = extracted_data["Date"] 
date_object = datetime.strptime(date_string, "%Y-%m-%d %H:%M:%S")
date_only = date_object.strftime("%Y-%m-%d")
reporter_name = extracted_data["Name"]
last_name = reporter_name.split()[-1]
photo_name = f".{save_directory}/{extracted_data["Species"]}_{date_only}_{last_name}.jpg"
download_image(last_url, photo_name)


""" if last_url:
  # Save the URL to a file
  
  save_url_to_file(last_url, filename)
  print(f"URL saved to {filename}")
else:
  print("No URL found in the email body.")    """
# %%


file_path = lan_record_loc+"Master Incidence Report Records.xlsx"
inc_reports = pd.read_excel(io=file_path, sheet_name = "Aquatic Reports")


names_spp = inc_reports[['Submitted_Common_Name', 'Submitted_Scientific_Name']]
names_spp = names_spp.drop_duplicates()   

# %%

spp_common_name = names_spp.loc[names_spp['Submitted_Scientific_Name'] == extracted_data["Species"], 'Submitted_Common_Name'].squeeze().replace(' ', '_')
reporter_name = extracted_data["Name"]
last_name = reporter_name.split()[-1]
report_name = f"{date_only}_{spp_common_name}_{last_name}"
report_date = date_object.strftime("%d-%m-%Y")
report_sender = extracted_data["Name"]
report_email = extracted_data["Email"]
report_phone = extracted_data["Phone"]
report_source = "App-Android"
report_photo_avail = "Y" if last_url else "N" 
report_photo_loc = "LAN Folder" if last_url else "NA" 
report_sub_com = spp_common_name
report_sub_sci = extracted_data["Species"]

report_lat = extracted_data["Location"].split(",")[0]
report_lon = extracted_data["Location"].split(",")[1]




report_frame = {'Incidence_Report_Name': report_name, 'Date': report_date, 'Sender': report_sender,
                'Email_Address': report_email, 'Phone_Number': report_phone, 'Source': report_source,
                'Photo_Available': report_photo_avail, 'Photo_Location': report_photo_loc,
                'Submitted_Common_Name': report_sub_com, 'Submitted_Scientific_Name': report_sub_sci,
                'Latitude': report_lat, 'Longitude': report_lon}
df = pd.DataFrame(report_frame, index=[0])


#wb = pd.read_excel("./output/report_form.xlsx")
df_existing = pd.read_excel("./output/report_form.xlsx", sheet_name='Sheet1')
# Assuming the column names match



df['Photo_Location'] = df['Photo_Location'].apply(
    lambda x: f'=HYPERLINK("file:///{os.path.abspath(photo_name).replace("\\", "/")}", "LAN Folder")' if pd.notna(x) else x
)
 
next_row = len(df_existing) + 1  # +2 because Excel is 1-indexed, and we want to start at row 2

        
output_file = './output/report_form.xlsx'

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df_existing.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Write the new data to the next row (start at next_row)
    for col_idx, column in enumerate(df.columns, start=1):
        worksheet.write(next_row, col_idx - 1, df[column].values[0])

    # Write the hyperlink formula for the 'Photo_Location' column (if not already done)
    photo_loc = df['Photo_Location'].values[0]
    worksheet.write(f'H{next_row + 1}', photo_loc)  # Write the actual hyperlink formula

# %%
