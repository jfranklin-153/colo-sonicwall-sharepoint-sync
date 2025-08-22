from dotenv import load_dotenv
from office365.graph_client import GraphClient
import os
import time
import math
# load environment file
load_dotenv(override=True)

# app credentials
SHAREPOINT_APPLICATION_ID = os.getenv("SHAREPOINT_APPLICATION_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
SHAREPOINT_TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
SHAREPOINT_ROOT_URL = os.getenv("SHAREPOINT_ROOT_URL")
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")

UPLOAD_SPEED_LIMIT_KBPS = int(
    os.getenv("UPLOAD_SPEED_LIMIT_KBPS", "100"))  # Default 100KB/s

# set the site url
site_url = f"{SHAREPOINT_ROOT_URL}/sites/{SHAREPOINT_SITE_NAME}"
print(SHAREPOINT_APPLICATION_ID)
print(SHAREPOINT_CLIENT_SECRET)
client = GraphClient(tenant=SHAREPOINT_TENANT_ID).with_client_secret(
    SHAREPOINT_APPLICATION_ID, SHAREPOINT_CLIENT_SECRET
)

# Fetch the site by URL
site = client.sites.get_by_url(site_url).get().execute_query()

# Access the default document library (e.g., "Documents" or "Shared Documents")
# Access the root of the default document library
document_library = site.drive.root
subdirectory_path = os.getenv("SHAREPOINT_SUBDIRECTORY")
subdirectory = document_library.get_by_path(
    subdirectory_path).get().execute_query()
# get monday of this week
today = time.localtime()
monday = today.tm_mday - today.tm_wday
# get the date of monday in YYYY-MM-DD format
monday_date = time.strftime("%Y-%m-%d", time.localtime(
    time.mktime((today.tm_year, today.tm_mon, monday, 0, 0, 0, 0, 0, 0))))

# get current script path
script_path = os.path.abspath(__file__)
LOCAL_UPLOAD_DIRECTORY = os.getenv('LOCAL_UPLOAD_DIRECTORY')
# get all file paths that end in .csv in the LOCAL_UPLOAD_DIRECTORY
csv_files = [os.path.join(LOCAL_UPLOAD_DIRECTORY, f) for f in os.listdir(
    LOCAL_UPLOAD_DIRECTORY) if f.endswith('.csv')]


def check_directory_exists(monday):
    # look for a subdirectory with the name of the current week
    final_directory_name = f"Week {monday}"
    try:
        final_directory = subdirectory.get_by_path(
            final_directory_name).get().execute_query()
    except Exception as e:
        print(f"Error checking for subdirectory '{final_directory_name}': {e}")
        if e.response.status_code == 404:
            # create the subdirectory
            final_directory = subdirectory.create_folder(
                final_directory_name).execute_query()
            print(f"Subdirectory '{final_directory_name}' created.")
        else:
            final_directory = None

    return final_directory


def throttled_upload(final_directory, file_path, file_name, speed_limit_kbps):
    chunk_size = speed_limit_kbps * 1024  # bytes per second
    with open(file_path, "rb") as f:
        file_size = os.path.getsize(file_path)
        total_chunks = math.ceil(file_size / chunk_size)
        print(
            f"Uploading '{file_name}' in {total_chunks} chunks at {speed_limit_kbps}KB/s...")
        # Read and simulate upload in chunks
        uploaded = 0
        for chunk_num in range(total_chunks):
            chunk = f.read(chunk_size)
            time.sleep(1)  # Sleep 1 second per chunk
            uploaded += len(chunk)
            print(f"Read {uploaded}/{file_size} bytes...")
        # After throttling, upload the file
        f.seek(0)
        final_directory.upload(file_name, f).execute_query()


def main():
    final_directory = check_directory_exists(monday_date)
    if not final_directory:
        exit(1)

    for csv_file in csv_files:
        file_path = os.path.join(LOCAL_UPLOAD_DIRECTORY, csv_file)
        file_name = os.path.basename(file_path)
        throttled_upload(final_directory, file_path,
                         file_name, UPLOAD_SPEED_LIMIT_KBPS)
        print(
            f"File '{file_name}' uploaded successfully to the document library.")


if __name__ == "__main__":
    main()
