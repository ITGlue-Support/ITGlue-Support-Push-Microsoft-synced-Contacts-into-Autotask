# ITGlue-Support-Push-Microsoft-synced-Contacts-into-Autotask
Author: Bart Jozwiak
This script will scan IT Glue account for organizations actively syncing with Microsoft and Autotask. From those organizations, it will filter out contacts that are actively syncing with Microsoft but not with Autotask. The script will create contacts inside Autotask in the corresponding companies.
Once the script is done, a full account sync is required to pick up the match
Script will prompt for IT Glue API key and Autotask credentials
There will be a prompt to choose between licensed and unlicensed contacts and a prompt to exclude any organizations
The final prompt will be displayed to confirm the total number of contacts to be created. 
Script will skip any contacts if the email already exists inside Autotask to avoid duplicates.
Any contact that doesn't have the required fields will also be skipped (First Name, Last Name and Email)

Prerequisites:
IT Glue API key
API User credentials generated inside Autotask (Security Level: APIUser System): Username, password, tracking identifier

<img width="515" height="547" alt="image" src="https://github.com/user-attachments/assets/ac91ece7-eff2-4f11-93f2-05aad62cfc8d" />

Python3.6+
The latest version of Python can be downloaded here: https://www.python.org/downloads/
The following packages installed: requests, tqdm

To avoid authorisation errors (401 and 403), inspect the Base URLs so they point to the correct Instance/Region

IT Glue URLs are pointing to the EU region by default. Those can be updated inside the following definitions:
1. def get_all_contact_ids
2. def fetch_contact_details
3. def get_autotask_syncing_orgs


Autotask Base URL can be found inside the main definition and by default is pointing to: "https://webservices15.autotask.net/atservicesrest/v1.0"
The webservices number has to be updated to match the UI inside Autotask:

<img width="1285" height="629" alt="image" src="https://github.com/user-attachments/assets/907334a8-952a-48c2-aea6-4e0d738ed4b2" />
