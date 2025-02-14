from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

def invoke(Arguments, go_Session, orchestrator_connection: OrchestratorConnection): 
    import requests
    import uuid
    import re
    import html
    import unicodedata
    import pandas as pd
    from datetime import datetime
    import os
    import json
    import os
    import time
    from office365.runtime.auth.authentication_context import AuthenticationContext
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
    from office365.runtime.auth.user_credential import UserCredential
    from urllib.parse import quote
    import re

    # Helper function for sanitizing sagstitel
    def sanitize_sagstitel(sagstitel):
        try:
            sagstitel = html.unescape(sagstitel)
            sagstitel = unicodedata.normalize('NFKC', sagstitel)
            sagstitel = sagstitel.replace('"', '')
            sagstitel = re.sub(r'[.:>#<*\?/%&{}\[\]\$!"@+\|\'=€]+', '', sagstitel)
            sagstitel = sagstitel.replace('\n', '').replace('\r', '')
            sagstitel = re.sub(r'[^a-zA-Z0-9ÆØÅæøå ]', '', sagstitel)
            sagstitel = re.sub(r' {2,}', ' ', sagstitel)
            sagstitel = sagstitel.strip()
                   # Check length and truncate if necessary
            if len(sagstitel) > 49:
                sagstitel = sagstitel[:50].strip()
            
            return sagstitel
        except Exception as e:
            raise Exception(f"Error during sanitazation: {str(e)}")

    # Initialize variables
    RobotUserName = Arguments.get("in_RobotUserName")
    RobotPassword = Arguments.get("in_RobotPassword")
    Sagsnummer = Arguments.get("in_Sagsnummer")
    GeoSag = Arguments.get("in_GeoSag")  
    NovaSag = Arguments.get("in_NovaSag")
    KMD_access_token = Arguments.get("KMD_access_token")
    KMDNovaURL = Arguments.get("KMDNovaURL")
    SharePointUrl = Arguments.get("in_SharePointUrl")
    sagstitel = ""  # Default value if no title is retrieved
    Overmappe = Arguments.get("in_Overmappe")
    Undermappe = Arguments.get("in_Undermappe")


    # --- Check if it's a Geo-sag ---
    if GeoSag:
        orchestrator_connection.log_info("Sagen er en Geo-sag, henter derfor sagstitel i GO")
        url = f"https://ad.go.aarhuskommune.dk/_goapi/Cases/Metadata/{Sagsnummer}"

        try:
            response = go_Session.get(url)
            print(f"Geo API Response Status Code: {response.status_code}")

            response_data = response.json()
            metadata = response_data.get("Metadata")

            if metadata:
                sagstitel = metadata.split('ows_Title="')[1].split('"')[0]
                print("Sagstitel (Geo):", sagstitel)
            else:
                print("Metadata field is missing in the response.")
        except Exception as e:
            raise Exception("Failed to extract Sagstitel (Geo):", str(e))


    # --- Check if it's a Nova-sag ---
    elif NovaSag:
        orchestrator_connection.log_info("Sagen er en Novasag, henter Sagstitel i NOVA")
        TransactionID = str(uuid.uuid4())
        url = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"

        headers = {
            "Authorization": f"Bearer {KMD_access_token}",
            "Content-Type": "application/json"
        }

        payload = {
            "common": {
                "transactionId": TransactionID
            },
            "paging": {
                "startRow": 1,
                "numberOfRows": 100
            },
            "caseAttributes": {
                "userFriendlyCaseNumber": Sagsnummer
            },
            "caseGetOutput": {
                "caseAttributes": {
                    "title": True,
                    "userFriendlyCaseNumber": True
                }
            }
        }

        try:
            response = requests.put(url, headers=headers, json=payload)
            print("Nova API Response:", response.status_code, response.text)

            if response.status_code == 200:
                sagstitel = response.json()['cases'][0]['caseAttributes']['title']
                print("Sagstitel (Nova):", sagstitel)
            else:
                print("Failed to fetch Sagstitel from NOVA. Status Code:", response.status_code)
        except Exception as e:
            raise Exception("Failed to fetch Sagstitel (Nova):", str(e))

    # Sanitize sagstitel regardless of source or failure
    sagstitel = sanitize_sagstitel(sagstitel)
    print(f"Final Sanitized Sagstitel: {sagstitel}")
    
    

        # ---- Henter dokumentlisten fra Sharepoint ----

    # Inputs
    site_relative_path = "/Teams/tea-teamsite10506/Delte Dokumenter"
    download_path = os.path.join("C:\\Users", os.getlogin(), "Downloads")


    # SharePoint authentication and client setup
    def sharepoint_client(RobotUserName, RobotPassword, SharePointUrl) -> ClientContext:
        try:
            credentials = UserCredential(RobotUserName, RobotPassword)
            ctx = ClientContext(SharePointUrl).with_credentials(credentials)

            # Load the SharePoint web to test the connection
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()

            return ctx
        except Exception as e:
            raise Exception(f"Authentication failed: {e}")


    # File downloading logic from SharePoint
    def download_file_from_sharepoint(client: ClientContext, sharepoint_file_url: str) -> str:
        """
        Downloads a file from SharePoint and returns the local file path.
        """
        file_name = sharepoint_file_url.split("/")[-1]  # Extract file name from URL
        local_file_path = os.path.join(download_path, file_name)  # Define local path

        try:
            # Ensure the download directory exists
            if not os.path.exists(download_path):
                os.makedirs(download_path)

            # Download the file
            with open(local_file_path, "wb") as local_file:
                client.web.get_file_by_server_relative_path(sharepoint_file_url).download(local_file).execute_query()
            return local_file_path
        except Exception as e:
            raise Exception(f"Error downloading file from SharePoint: {e}")


    # Main logic
    # Authenticate to SharePoint
    client = sharepoint_client(RobotUserName, RobotPassword, SharePointUrl)

    # Construct paths for Overmappe and Undermappe without over-encoding
    overmappe_url = f"{site_relative_path}/Dokumentlister/{Overmappe}"
    print(f"Overmappe URL: {overmappe_url}")
    overmappe_folder = client.web.get_folder_by_server_relative_url(overmappe_url)
    client.load(overmappe_folder)
    client.execute_query()


    undermappe_url = f"{overmappe_url}/{Undermappe}"
    undermappe_folder = client.web.get_folder_by_server_relative_url(undermappe_url)
    client.load(undermappe_folder)
    client.execute_query()

    # Fetch files in the Undermappe folder
    print("Fetching files from the folder...")
    files = undermappe_folder.files
    client.load(files)
    client.execute_query()

    # Print and process file names
    data_table = []  # To store file information with dates

    for file in files:
        file_name = file.properties["Name"]

        dokument_date = None  # Initialize dokument_date

        if "_" in file_name:
            try:
                # Extract the part after the first underscore
                date_part = file_name.split("_")[1]
                date_str = date_part.split(".")[0]  # Part before the first dot
                dokument_date = datetime.strptime(date_str, "%d-%m-%Y")
            except (IndexError, ValueError):
                print(f"  -> Error parsing date from: {file_name}. Defaulting to 01-01-2023")
                dokument_date = datetime.strptime("01-01-2023", "%d-%m-%Y")
        else:
            print(f"  -> No underscore found in: {file_name}. Defaulting to 01-01-2023")
            dokument_date = datetime.strptime("01-01-2023", "%d-%m-%Y")

        data_table.append({
            "FileName": file_name,
            "DocumentDate": dokument_date.strftime('%d-%m-%Y')
        })

    # Sort files by date in descending order
    data_table = sorted(
        data_table, 
        key=lambda x: datetime.strptime(x["DocumentDate"], "%d-%m-%Y"), 
        reverse=True
    )

    for entry in data_table:
        orchestrator_connection.log_info(f" - {entry['FileName']} (Date: {entry['DocumentDate']})")

    # Download the newest file if available
    if data_table:
        newest_file = data_table[0]
        newest_file_name = newest_file["FileName"]
        DokumentlisteDatoString = newest_file["DocumentDate"] 
        sharepoint_file_url = f"{undermappe_url}/{newest_file_name}"

        local_file_path = download_file_from_sharepoint(client, sharepoint_file_url)
    if local_file_path.endswith('.xlsx'):
        # Read Excel file into a Pandas DataFrame
        dt_DocumentList = pd.read_excel(local_file_path)
        os.remove(local_file_path)
        # Return the DataFrame to be used later
    else:
        raise Exception(f"Downloaded file is not an Excel file: {local_file_path}")

    return {
    "sagstitel": sagstitel,
    "dt_DocumentList": dt_DocumentList,
    "out_DokumentlisteDatoString": DokumentlisteDatoString
    }