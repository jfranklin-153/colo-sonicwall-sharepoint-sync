from dotenv import load_dotenv
from office365.graph_client import GraphClient
import os
import time
# load environment file
load_dotenv(override = True)

# app credentials
SHAREPOINT_APPLICATION_ID = os.getenv("SHAREPOINT_APPLICATION_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
SHAREPOINT_TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
SHAREPOINT_ROOT_URL = os.getenv("SHAREPOINT_ROOT_URL")
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")

# set the site url
site_url = f"{SHAREPOINT_ROOT_URL}/sites/{SHAREPOINT_SITE_NAME}"

client = GraphClient(tenant=SHAREPOINT_TENANT_ID).with_client_secret(
    SHAREPOINT_APPLICATION_ID, SHAREPOINT_CLIENT_SECRET
)

# Fetch the site by URL
site = client.sites.get_by_url(site_url).get().execute_query()

# Access the default document library (e.g., "Documents" or "Shared Documents")
document_library = site.drive.root  # Access the root of the default document library
subdirectory_path = os.getenv("SHAREPOINT_SUBDIRECTORY")
subdirectory = document_library.get_by_path(subdirectory_path).get().execute_query()
# get monday of this week
today = time.localtime()
monday = today.tm_mday - today.tm_wday
# get the date of monday in YYYY-MM-DD format
monday_date = time.strftime("%Y-%m-%d", time.localtime(time.mktime((today.tm_year, today.tm_mon, monday, 0, 0, 0, 0, 0, 0))))

# get current script path
script_path = os.path.abspath(__file__)
script_dir = os.path.dirname(script_path)
# get all file names that end in.csv in the script directory
csv_files = [f for f in os.listdir(script_dir) if f.endswith('.csv')]

def check_directory_exists(monday):
    # look for a subdirectory with the name of the current week
    subdirectory_name = f"Week {monday}"
    subdirectory = document_library.get_by_path(subdirectory_name).get().execute_query()
    if subdirectory:
        print(f"Subdirectory '{subdirectory_name}' already exists.")
    else:
        # create the subdirectory
        subdirectory = document_library.add_folder(subdirectory_name).execute_query()
        print(f"Subdirectory '{subdirectory_name}' created.")
        return subdirectory

def main():
    final_directory = check_directory_exists(monday_date)
    file_path = "local_file.csv"
    file_name = os.path.basename(file_path)

    with open(file_path, "rb") as file_content:
        final_directory.upload(file_name, file_content).execute_query()

    print(f"File '{file_name}' uploaded successfully to the document library.")

if __name__ == "__main__":
    main()