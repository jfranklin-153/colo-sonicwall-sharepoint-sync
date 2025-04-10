from dotenv import load_dotenv
from office365.graph_client import GraphClient
import os
# load environment file
load_dotenv(override = True)

# app credentials
SHAREPOINT_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
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

# Upload a local file into the document library
file_path = "local_file.csv"
file_name = os.path.basename(file_path)

with open(file_path, "rb") as file_content:
    document_library.upload(file_name, file_content).execute_query()

print(f"File '{file_name}' uploaded successfully to the document library.")