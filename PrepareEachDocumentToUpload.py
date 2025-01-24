def invoke_PrepareEachDocumentToUpload(Arguments_PrepareEachDocumentToUpload):
    import pandas as pd
    import re
    import requests
    from requests_ntlm import HttpNtlmAuth
    import json
    import os
    import time
    from datetime import datetime
    from msal import PublicClientApplication
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.sharing.links.kind import SharingLinkKind
    from office365.sharepoint.webs.web import Web
    from office365.runtime.auth.user_credential import UserCredential
    import json
    from SendSMTPMail import send_email, EmailAttachment
    import shutil
    import uuid
    from SharePointUploader import upload_file_to_sharepoint
    import re

    # henter in_argumenter:
    dt_DocumentList = Arguments_PrepareEachDocumentToUpload.get("in_dt_Documentlist")
    CloudConvertAPI = Arguments_PrepareEachDocumentToUpload.get("in_CloudConvertAPI")
    MailModtager = Arguments_PrepareEachDocumentToUpload.get("in_MailModtager")
    UdviklerMailAktbob = Arguments_PrepareEachDocumentToUpload.get("in_UdviklerMail")
    RobotUserName = Arguments_PrepareEachDocumentToUpload.get("in_RobotUserName")
    RobotPassword = Arguments_PrepareEachDocumentToUpload.get("in_RobotPassword")
    FilarkivCaseID = Arguments_PrepareEachDocumentToUpload.get("in_FilarkivCaseID")
    SharePointAppID = Arguments_PrepareEachDocumentToUpload.get("in_SharePointAppID")
    SharePointTenant = Arguments_PrepareEachDocumentToUpload.get("in_SharePointTenant")
    SharePointURL = Arguments_PrepareEachDocumentToUpload.get("in_SharePointUrl")
    Overmappe = Arguments_PrepareEachDocumentToUpload.get("in_Overmappe")
    Undermappe = Arguments_PrepareEachDocumentToUpload.get("in_Undermappe")
    Sagsnummer = Arguments_PrepareEachDocumentToUpload.get("in_Sagsnummer")
    GeoSag = Arguments_PrepareEachDocumentToUpload.get("in_GeoSag")
    NovaSag = Arguments_PrepareEachDocumentToUpload.get("in_NovaSag")
    FilarkivURL = Arguments_PrepareEachDocumentToUpload.get("in_FilarkivURL")
    Filarkiv_access_token = Arguments_PrepareEachDocumentToUpload.get("in_Filarkiv_access_token")
    KMDNovaURL = Arguments_PrepareEachDocumentToUpload.get("in_KMDNovaURL")
    KMD_access_token = Arguments_PrepareEachDocumentToUpload.get("in_NovaToken")
    GoUsername = Arguments_PrepareEachDocumentToUpload.get("in_GoUsername")
    GoPassword = Arguments_PrepareEachDocumentToUpload.get("in_GoPassword")

    # Define the structure of the data table
    dt_AktIndex = {
        "Akt ID": pd.Series(dtype="int32"),
        "Filnavn": pd.Series(dtype="string"),
        "Dokumentkategori": pd.Series(dtype="string"),
        "Dokumentdato": pd.Series(dtype="datetime64[ns]"),
        "Dok ID": pd.Series(dtype="string"),
        "Bilag til Dok ID": pd.Series(dtype="string"),
        "Bilag": pd.Series(dtype="string"),
        "Omfattet af aktindsigt?": pd.Series(dtype="string"),
        "Gives der aktindsigt?": pd.Series(dtype="string"),
        "Begrundelse hvis Nej/Delvis": pd.Series(dtype="string"),
        "IsDocumentPDF": pd.Series(dtype="bool"),
    }

    #Functions: 
    def sanitize_title(Titel):
        # 1. Replace double quotes with an empty string
        Titel = Titel.replace("\"", "")

        # 2. Remove special characters with regex
        Titel = re.sub(r"[.:>#<*\?/%&{}\$!\"@+\|'=]+", "", Titel)

        # 3. Remove any newline characters
        Titel = Titel.replace("\n", "").replace("\r", "")

        # 4. Trim leading and trailing whitespace
        Titel = Titel.strip()

        # 5. Remove non-alphanumeric characters except spaces and Danish letters
        Titel = re.sub(r"[^a-zA-Z0-9ÆØÅæøå ]", "", Titel)

        # 6. Replace multiple spaces with a single space
        Titel = re.sub(r" {2,}", " ", Titel)

        return Titel
    
    def calculate_available_title_length(base_path, Overmappe, Undermappe, AktID, DokumentID, Titel, max_path_length=400):
        overmappe_length = len(Overmappe)
        undermappe_length = len(Undermappe)
        aktID_length = len(str(AktID))
        dokID_length = len(str(DokumentID))

        fixed_length = len(base_path) + overmappe_length + undermappe_length + aktID_length + dokID_length + 7
        available_title_length = max_path_length - fixed_length

        if len(Titel) > available_title_length:
            return Titel[:available_title_length]
        
        return Titel

    def upload_to_filarkiv(FilarkivURL, FilarkivCaseID, Filarkiv_access_token, AktID, DokumentID, Titel, FilePath):
        print("Uploader til Filarkiv")
        DoesFolderExists = False
        Filarkiv_DocumentID = None  # Ensure it is initialized
        FileName = f"{AktID:04} - {DokumentID} - {Titel}"
        print(f"FilarkivCaseID: {FilarkivCaseID}")
        url = f"{FilarkivURL}/Documents/CaseDocumentOverview?caseId={FilarkivCaseID}&pageIndex=1&pageSize=500"

        headers = {
            "Authorization": f"Bearer {Filarkiv_access_token}",
            "Content-Type": "application/xml"
        }
        
        try:
            response = requests.get(url, headers=headers)
            print("FilArkiv respons:", response.status_code)
            
            if response.status_code == 200:
                response_json = response.json()
                
                if not response_json:
                    print("Der findes ingen dokumenter på sagen")
                    DocumentDate = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                    DocumentNumber = 1
                    data = {
                        "caseId": FilarkivCaseID,
                        "securityClassificationLevel": 0,
                        "title": FileName,
                        "documentNumber": DocumentNumber,
                        "documentDate": DocumentDate,
                        "direction": 0
                    }
                    response = requests.post("https://core.filarkiv.dk/api/v1/Documents", headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, data=json.dumps(data))  
                    if response.status_code in [200, 201]:
                        Filarkiv_DocumentID = response.json().get("id")
                    else:
                        print("Failed to create document. Response:", response.text)
                else:
                    for current_item in response_json:
                        if FileName in current_item.get("title", ""):
                            print("Dokument Mappen er oprettet")
                            Filarkiv_DocumentID = current_item.get("id")
                            DoesFolderExists = True
                            break  # Exit loop once found
                    if not DoesFolderExists:
                        print("Finder det nye dokumentnummer")
                        HighestDocumentNumber = max((int(i.get("documentNumber", 0)) for i in response_json), default=1)
                        DocumentNumber = HighestDocumentNumber + 1
                        DocumentDate = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                        data = {
                            "caseId": FilarkivCaseID,
                            "securityClassificationLevel": 0,
                            "title": FileName,
                            "documentNumber": DocumentNumber,
                            "documentDate": DocumentDate,
                            "direction": 0
                        }
                        response = requests.post("https://core.filarkiv.dk/api/v1/Documents", headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, data=json.dumps(data))
                        if response.status_code in [200, 201]:
                            Filarkiv_DocumentID = response.json().get("id")
                            print("Anvender følgende Filarkiv_DocumentID:", Filarkiv_DocumentID)
                        else:
                            print("Failed to create document. Response:", response.text)
            else:
                print("Failed to fetch data, status code:", response.status_code)
        except Exception as e:
            print("Kunne ikke hente dokumentinformation:", str(e))
        
        if Filarkiv_DocumentID is None:
            print("Fejl: Filarkiv_DocumentID blev ikke genereret. Afbryder processen.")
            return
        
        if not DoesFolderExists:
            extension = os.path.splitext(FilePath)[1]
            mime_type = {
                ".txt": "text/plain", ".pdf": "application/pdf", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
                ".gif": "image/gif", ".doc": "application/msword", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xls": "application/vnd.ms-excel", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".csv": "text/csv",
                ".json": "application/json", ".xml": "application/xml"
            }.get(extension, "application/octet-stream")
            FileName += extension
            print(f"Anvender følgende dokumentID: {Filarkiv_DocumentID}")
            response = requests.post("https://core.filarkiv.dk/api/v1/Files", headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, json={"documentId": Filarkiv_DocumentID, "fileName": FileName, "sequenceNumber": 0, "mimeType": mime_type})
            if response.status_code in [200, 201]:
                FileID = response.json().get('id')
                print(f"FileID: {FileID}")
            else:
                print("Failed to create file metadata.", response.text)
            
            url = f"https://core.filarkiv.dk/api/v1/FileIO/Upload/{FileID}"
            if not os.path.exists(FilePath):
                print(f"Error: File not found at {FilePath}")
            else:
                with open(FilePath, 'rb') as file:
                    files = [('file', (FileName, file, mime_type))]
                    response = requests.post(url, headers={"Authorization": f"Bearer {Filarkiv_access_token}"}, files=files)
                    if response.status_code in [200, 201]:
                        print("File uploaded successfully.")
                        print(response.json())
                    else:
                        print(f"Failed to upload file. Status Code: {response.status_code}")

    # Create an empty DataFrame with the defined structure
    dt_AktIndex = pd.DataFrame(dt_AktIndex)
    

    # ---- If-statement som tjekker om det er en GeoSag eller NovaSag ----
    if GeoSag == True:
        #Sagen er en geo sag 
        dt_DocumentList['Dokumentdato'] = pd.to_datetime(dt_DocumentList['Dokumentdato'], errors='coerce')
        for index, row in dt_DocumentList.iterrows():
            # Convert items to strings unless they are explicitly integers
            Omfattet = str(row["Omfattet af ansøgningen? (Ja/Nej)"])
            DokumentID = str(row["Dok ID"])
            
            # Handle AktID conversion
            AktID = row['Akt ID']
            if isinstance(AktID, str):  
                AktID = int(AktID.replace('.', ''))
            elif isinstance(AktID, int):  
                AktID = AktID

            Titel = str(row["Dokumenttitel"])
            BilagTilDok = str(row["Bilag til Dok ID"])
            DokBilag = str(row["Bilag"])
            Dokumentkategori = str(row["Dokumentkategori"])
            Aktstatus = str(row["Gives der aktindsigt i dokumentet? (Ja/Nej/Delvis)"])
            Begrundelse = str(row["Begrundelse hvis nej eller delvis"])
            Dokumentdato = datetime.strptime(row["Dokumentdato"], "%d-%m-%Y").strftime("%d-%m-%Y")
            IsDocumentPDF = True
            print(f"AktID til debug: {AktID}")

            # Declare the necessary variables
            base_path = "Teams/tea-teamsite10506/Delte dokumenter/Aktindsigter/"

            # Sanitize the title
            Titel = sanitize_title(Titel)

            Titel = calculate_available_title_length(base_path, Overmappe, Undermappe, AktID, DokumentID, Titel)

            print("Adjusted Title:", Titel)

            if (("ja" in Aktstatus.lower() or "delvis" in Aktstatus.lower()) 
                and DokumentID != "" 
                and "ja" in Omfattet.lower()):
                
                print("Dokumentet er omfattet i ansøgningen")

                    # Define the variables
                # Construct the URL
                url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/Data/{DokumentID}"

                # Set up the NTLM authentication
                auth = HttpNtlmAuth(GoUsername, GoPassword)

                # Make the GET request
                response = requests.get(url, headers={"Content-Type": "application/json"}, auth=auth)

                # Get the response content
                DocumentData = response.text  # Or response.content for raw bytes
                # Extract the `ItemProperties` field, which contains XML-like data
                data = json.loads(DocumentData)
                item_properties = data.get("ItemProperties", "")

                # Use regex to extract `_File_x0020_Type` and `_UIVersionString`
                file_type_match = re.search(r'ows_File_x0020_Type="([^"]+)"', item_properties)
                version_ui_match = re.search(r'ows__UIVersionString="([^"]+)"', item_properties)

                # Extract values or set default if not found
                DokumentType = file_type_match.group(1) if file_type_match else "Not found"
                VersionUI = version_ui_match.group(1) if version_ui_match else "Not found"
                Feedback = " "
                FilePath = os.path.join(
                    "C:\\Users",
                    os.getenv("USERNAME"),  # Fetches the current username
                    "Downloads",
                    f"{AktID:04} - {DokumentID} - {Titel}"
                    )

                # Tjekker om Goref-fil
                if ".goref" in FilePath:
                    url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/DocumentBytes/{DokumentID}"

                    # Initialize the session
                    session = requests.Session()
                    session.auth = HttpNtlmAuth(GoUsername, RobotPassword)

                    # Create the request
                    headers = {"Content-Type": "application/json"}
                    response = session.get(url, headers=headers)

                    # Check the response
                    if response.status_code == 200:
                        # Get the raw bytes
                        ByteResult = response.content
                    else:
                        print(f"Request failed with status code: {response.status_code}")

                    # Write bytes to a file
                    with open(FilePath, "wb") as file:
                        file.write(ByteResult)

                    # Read the file content
                    with open(FilePath, "r", encoding="utf-8") as file:
                        RefDokument = file.read()

                    # Process RefDokument to extract refdocument and DokumentID
                    refdocument = RefDokument.split("?docid=")[1]
                    DokumentID = refdocument.split('"')[0]

                    # Delete the file
                    if os.path.exists(FilePath):
                        os.remove(FilePath)
                        print("File deleted.")
                    else:
                        print("The file does not exist.")

                    #Henter dokument data
                    url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/Data/{DokumentID}"

                    # Set up the NTLM authentication
                    auth = HttpNtlmAuth(GoUsername, GoPassword)

                    # Make the GET request
                    response = requests.get(url, headers={"Content-Type": "application/json"}, auth=auth)

                    # Get the response content
                    DocumentData = response.text  # Or response.content for raw bytes
                    # Extract the `ItemProperties` field, which contains XML-like data
                    data = json.loads(DocumentData)
                    item_properties = data.get("ItemProperties", "")

                    # Use regex to extract `_File_x0020_Type` and `_UIVersionString`
                    file_type_match = re.search(r'ows_File_x0020_Type="([^"]+)"', item_properties)
                    version_ui_match = re.search(r'ows__UIVersionString="([^"]+)"', item_properties)

                    # Extract values or set default if not found
                    DokumentType = file_type_match.group(1) if file_type_match else "Not found"
                    VersionUI = version_ui_match.group(1) if version_ui_match else "Not found"
                    Feedback = ""
                    FilePath = os.path.join(
                        "C:\\Users",
                        os.getenv("USERNAME"),  # Fetches the current username
                        "Downloads",
                        f"{AktID:04} - {DokumentID} - {Titel}"
                        )

                if DokumentType.lower() == "pdf":
                    print("Allerede PDf - downloader")
                    # Create the URL
                    url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/DocumentBytes/{DokumentID}"

                    # Retry mechanism
                    max_retries = 30
                    retry_interval = 5  # Seconds

                    ByteResult = None
                    for attempt in range(max_retries):
                        try:
                            # Make the API request
                            response = requests.get(
                                url,
                                auth=HttpNtlmAuth(GoUsername, GoPassword),
                                headers={"Content-Type": "application/json"},
                                timeout=60  # Timeout for the request
                            )

                            # Check if the response is successful
                            if response.status_code == 200:
                                ByteResult = response.content  # Extract the file bytes
                                break  # Exit the retry loop
                            else:
                                print(f"Attempt {attempt + 1}: Failed with status code {response.status_code}")
                        except Exception as e:
                            print(f"Attempt {attempt + 1}: Exception occurred - {e}")

                        # Wait before retrying
                        time.sleep(retry_interval)
                    else:
                        print("Max retries reached. File download failed.")

                    # ByteResult will contain the file bytes if successful
                    if ByteResult:
                        print(f"File size: {len(ByteResult)} bytes")
                    else:
                        print("No file was downloaded.")
                
                #Dokumentet er ikke en pdf - forsøger at konverterer
                else: 
                    print("Konverterer filen til PDF")
                    # Construct the URL
                    url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/ConvertToPDF/{DokumentID}/{VersionUI}"

                    # Make the request
                    try:
                        response = requests.get(
                            url,
                            auth=HttpNtlmAuth(GoUsername, GoPassword),
                            headers={"Content-Type": "application/json"},
                            timeout=None  # Equivalent to client.Timeout = -1
                        )
                        
                        # Error message
                        if response.status_code != 200:
                            print(f"Error Message: {response.text}")
                        
                        # Feedback and byte result
                        Feedback = response.text
                        ByteResult = response.content
                        
                        # Check if ByteResult is empty
                        if len(ByteResult) == 0:
                            print(f"Status Code: {response.status_code}")
                        else:
                            print("ByteResult received successfully.")

                    except Exception as e:
                        print(f"An exception occurred: {e}")
                    
                    
                    # tjekker om go-conversion lykkedes 
                    if "Document could not be converted" in Feedback or len(ByteResult) == 0:
                        print("GO conversion failed, attempting cloudconvert")

                        # Create the API URL
                        url = f"https://api.cloudconvert.com/v2/convert/formats?filter[input_format]={DokumentType}&filter[output_format]=pdf&filter[operation]=convert"

                        # Add the authorization header
                        headers = {
                            "Authorization": CloudConvertAPI
                        }

                        # Initialize conversionPossible
                        conversionPossible = False

                        # Execute the request
                        response = requests.get(url, headers=headers)

                        # Process the response
                        if response.status_code == 200 and response.text.strip():
                            # Parse the response content to a dictionary
                            jsonResponse = json.loads(response.text)
                            
                            # Check if the data array contains any elements
                            data = jsonResponse.get("data", [])
                            if data:
                                # Iterate through each conversion option
                                for item in data:
                                    operation = item.get("operation")
                                    inputFormatCheck = item.get("input_format")
                                    outputFormat = item.get("output_format")
                                    
                                    # Check if it matches the desired conversion
                                    if operation == "convert" and inputFormatCheck == DokumentType and outputFormat == "pdf":
                                        conversionPossible = True
                                        break
                        
                        if not conversionPossible:
                            print(f"Skipping cause CloudConvert doesn't support: {DokumentType}->PDF")
                            ByteResult = bytes()                  
                        else:
                            print("Conversion is supported!")
                            url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/DocumentBytes/{DokumentID}"

                            # Make the request
                            try:
                                response = requests.get(
                                    url,
                                    auth=HttpNtlmAuth(GoUsername, GoPassword),
                                    headers={"Content-Type": "application/json"},
                                    timeout=None  # No timeout (similar to client.Timeout = -1)
                                )

                                # Check the response
                                if response.status_code == 200:
                                    ByteResult = response.content  # Equivalent to response.RawBytes in C#
                                else:
                                    print(f"Request failed with status code {response.status_code}")

                            except Exception as e:
                                print(f"An error occurred: {e}")

                            # Variables
                            FilnavnFørPdf = f"Output.{DokumentType}"

                            # Write ByteResult to a file
                            ByteResult = b""  # Replace with the actual byte content
                            with open(FilnavnFørPdf, "wb") as file:
                                file.write(ByteResult)

                            # Step 1: Create the job with an import/upload task
                            create_job_url = "https://api.cloudconvert.com/v2/jobs"
                            create_job_headers = {
                                "Authorization": CloudConvertAPI,
                                "Content-Type": "application/json",
                            }
                            json_body = {
                                "tasks": {
                                    "import_1": {
                                        "operation": "import/upload"
                                    },
                                },
                                "tag": f"Aktbob-{DokumentID}-{time.strftime('%H-%M-%S')}",
                            }
                            response = requests.post(create_job_url, headers=create_job_headers, json=json_body)
                            job_response_data = response.json()

                            # Extract upload URL and parameters
                            upload_url = job_response_data["data"]["tasks"][0]["result"]["form"]["url"]
                            upload_parameters = job_response_data["data"]["tasks"][0]["result"]["form"]["parameters"]

                            # Step 2: Perform file upload
                            upload_data = {param["name"]: param["value"] for param in upload_parameters}
                            upload_files = {"file": open(FilnavnFørPdf, "rb")}
                            upload_response = requests.post(upload_url, data=upload_data, files=upload_files)
                            os.remove(FilnavnFørPdf)

                            upload_task_id = job_response_data["data"]["tasks"][0]["id"]

                            # Step 3: Create convert and export tasks
                            convert_export_body = {
                                "tasks": {
                                    "convert_1": {
                                        "operation": "convert",
                                        "input": [upload_task_id],
                                        "input_format": DokumentType,
                                        "output_format": "pdf",
                                    },
                                    "export_1": {
                                        "operation": "export/url",
                                        "input": ["convert_1"],
                                    }
                                },
                                "tag": f"Aktbob-{DokumentID}-{time.strftime('%H-%M-%S')}",
                            }
                            convert_export_response = requests.post(
                                create_job_url, headers=create_job_headers, json=convert_export_body
                            )
                            convert_export_response_data = convert_export_response.json()
                            print(convert_export_response.text)

                            if "INVALID_CONVERSION_TYPE" not in convert_export_response.text:
                                export_task_id = convert_export_response_data["data"]["tasks"][1]["id"]

                                # Step 4: Check export task status
                                while True:
                                    status_check_url = f"https://api.cloudconvert.com/v2/tasks/{export_task_id}"
                                    status_check_response = requests.get(status_check_url, headers=create_job_headers)
                                    status_check_data = status_check_response.json()

                                    task_status = status_check_data["data"]["status"]

                                    if task_status == "finished":
                                        # Extract the download URL
                                        download_url = status_check_data["data"]["result"]["files"][0]["url"]

                                        # Download the file
                                        file_path = "Output.pdf"
                                        with requests.get(download_url, stream=True) as r:
                                            with open(file_path, "wb") as file:
                                                for chunk in r.iter_content(chunk_size=8192):
                                                    file.write(chunk)

                                        print("File downloaded successfully.")

                                        # Read the file into ByteResult
                                        with open(file_path, "rb") as file:
                                            ByteResult = file.read()

                                        os.remove(file_path)
                                        break
                                    elif task_status not in ["waiting", "processing"]:
                                        print("An error occurred:", status_check_response.text)
                                        ByteResult = b""
                                        break

                                    time.sleep(5)  # Wait for 5 seconds before checking again
                            
                            

                            if len(ByteResult) == 0:
                                print("ByteResult is empty.")
                            else:
                                Feedback = "CloudConvert lykkedes"
                                print(Feedback)


                if "Document could not be converted" in Feedback or len(ByteResult) == 0:
                    print(f"Could not be converted, uploading as {FilePath}.{DokumentType}")
                    # Construct the URL
                    url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/DocumentBytes/{DokumentID}"

                    # Retry mechanism
                    max_retries = 60
                    retry_interval = 5  # seconds
                    ByteResult = None

                    for attempt in range(max_retries):
                        try:
                            # Make the API request
                            response = requests.get(
                                url,
                                auth=HttpNtlmAuth(GoUsername, GoPassword),
                                headers={"Content-Type": "application/json"},
                                timeout=60  # Timeout for the request
                            )
                            
                            if response.status_code == 200:  # Check if request was successful
                                ByteResult = response.content  # Raw bytes from the response
                                print(f"Success! ByteResult size: {len(ByteResult)} bytes")
                                break  # Exit the retry loop
                            else:
                                print(f"Attempt {attempt + 1}: Failed with status code {response.status_code}")
                        except Exception as e:
                            print(f"Attempt {attempt + 1}: Exception occurred - {e}")
                        
                        # Wait before retrying
                        time.sleep(retry_interval)
                    else:
                        print("Max retries reached. File download failed.")

                    # Final ByteResult check
                    if ByteResult:
                        print("File downloaded successfully.")
                    else:
                        print("ByteResult is empty.")

                
                    FilePath = (f"{FilePath}.{DokumentType}")

                else: 
                    FilePath = (f"{FilePath}.pdf")   

                print("Gemmer fil")
                # Variables

                try:
                    # Step 1: Attempt to write ByteResult to the file
                    with open(FilePath, "wb") as file:
                        file.write(ByteResult)
                    print("File written successfully.")
                    file.close
                except Exception as initial_exception:
                    print(f"Failed, trying from URL: {DokumentID} Path: {FilePath}")
                    print(initial_exception)

                    # Assign ByteResult to an empty byte array
                    ByteResult = bytes()

                    try:
                        # Step 2: Retry scope
                        max_retries = 2
                        for attempt in range(max_retries):
                            try:
                                # Fetch metadata to retrieve the document URL
                                metadata_url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/MetadataWithSystemFields/{DokumentID}"
                                metadata_response = requests.get(
                                    metadata_url,
                                    auth=HttpNtlmAuth(GoUsername, GoPassword),
                                    headers={"Content-Type": "application/json"},
                                    timeout=60
                                )
                                
                                # Parse the document URL
                                content = metadata_response.text
                                DocumentURL = content.split("ows_EncodedAbsUrl=")[1].split('"')[1]
                                DocumentURL = DocumentURL.split("\\")[0].replace("go.aarhus", "ad.go.aarhus")
                                print(f"Document URL: {DocumentURL}")
                                
                                # Download the file
                                handler = requests.Session()
                                handler.auth = HttpNtlmAuth(GoUsername, GoPassword)
                                with handler.get(DocumentURL, stream=True) as download_response:
                                    download_response.raise_for_status()  # Ensure the request was successful
                                    with open(FilePath, "wb") as file:
                                        for chunk in download_response.iter_content(chunk_size=8192):
                                            file.write(chunk)

                                print("File downloaded successfully.")
                                break  # Exit the retry loop after success
                            except Exception as retry_exception:
                                print(f"Retry {attempt + 1} failed: {retry_exception}")
                                if attempt == max_retries - 1:
                                    raise RuntimeError(
                                        f"Failed to download file after {max_retries} retries. "
                                        f"DokumentID: {DokumentID}, Path: {FilePath}"
                                    )
                                time.sleep(5)  # Wait before the next retry
                    except RuntimeError as nested_exception:
                        # Step 3: Catch the error from the nested retry scope
                        print(f"An unrecoverable error occurred: {nested_exception}")
                        raise nested_exception  # Re-raise the error to propagate

                if ".pdf" in FilePath:
                    upload_to_filarkiv(FilarkivURL,FilarkivCaseID, Filarkiv_access_token, AktID, DokumentID,Titel, file_path)

                
                else: # Filtypen er ikke understøttet, uploader til Sharepoint
                    print("FilTypen er ikke understøttet, uploader til Sharepoint")
                    IsDocumentPDF = False
                    site_url = SharePointURL

                    if os.path.getsize(FilePath) > 262143999:
                        print("Filen er større end 262 mb")
                        upload_file_to_sharepoint(
                            site_url=SharePointURL,
                            overmappe=Overmappe,
                            undermappe=Undermappe,
                            file_path=file_path,
                            sharepoint_app_id=SharePointAppID,
                            sharepoint_tenant=SharePointTenant,
                            robot_username=RobotUserName,
                            robot_password=RobotPassword
                        )

                    else:
                        print("Filen er mindre end 250 MB")

                        try:
                            # Input variables
                            sharepoint_folder_path = f"/Aktindsigter/{Overmappe}/{Undermappe}"
                            file_name = os.path.basename(FilePath)

                            # Normalize the site URL
                            if site_url.startswith("https://"):
                                site_url = site_url[8:]

                            site_url = site_url.replace(".sharepoint.com", ".sharepoint.com:")

                            scopes = ["https://graph.microsoft.com/.default"]

                            # Create the MSAL app for token generation
                            msal_app = PublicClientApplication(
                                client_id=SharePointAppID,
                                authority=f"https://login.microsoftonline.com/{SharePointTenant}"
                            )

                            print("Getting access token...")
                            token_response = msal_app.acquire_token_by_username_password(
                                username=RobotUserName,
                                password=RobotPassword,
                                scopes=scopes
                            )

                            if "access_token" not in token_response:
                                raise Exception(f"Failed to acquire token: {token_response}")

                            access_token = token_response["access_token"]
                            headers = {"Authorization": f"Bearer {access_token}"}

                            # Get the site ID
                            site_request_url = f"https://graph.microsoft.com/v1.0/sites/{site_url}"


                            site_response = requests.get(site_request_url, headers=headers)
                            site_response.raise_for_status()
                            site_json = site_response.json()

                            if "id" not in site_json:
                                raise Exception("Key 'id' not found in site response")

                            site_id = site_json["id"]

                            # Get the drive ID
                            drive_request_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
                            drive_response = requests.get(drive_request_url, headers=headers)
                            drive_response.raise_for_status()
                            drive_json = drive_response.json()

                            if "id" not in drive_json:
                                raise Exception("Key 'id' not found in drive response")

                            drive_id = drive_json["id"]

                            # Upload the file
                            drive_item_request_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_folder_path}/{file_name}:/content"

                            with open(FilePath, "rb") as file_stream:
                                upload_response = requests.put(
                                    drive_item_request_url,
                                    headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/octet-stream"},
                                    data=file_stream
                                )

                            if not upload_response.ok:
                                raise Exception(f"File upload failed: {upload_response.text}")

                            print("Upload complete")

                        except Exception as ex:
                            print(f"Error: {ex}")
                            print("Filen kunne ikke overføres, prøver chunk upload")
                            upload_file_to_sharepoint(
                                site_url=SharePointURL,
                                overmappe=Overmappe,
                                undermappe=Undermappe,
                                file_path=file_path,
                                sharepoint_app_id=SharePointAppID,
                                sharepoint_tenant=SharePointTenant,
                                robot_username=RobotUserName,
                                robot_password=RobotPassword
                            )
                
                Titel = FilePath.split("\\Downloads\\")[1]

            else:
                print("Dokumentet skal ikke med i ansøgningen")
                Titel = f"{AktID:04} - {DokumentID} - {Titel}"
                

            
            # Parse and prepare data for the row
            row_to_add = {
                "Akt ID": int(AktID),
                "Filnavn": Titel,
                "Dokumentkategori": Dokumentkategori,
                "Dokumentdato": datetime.strptime(Dokumentdato, "%d-%m-%Y"),
                "Dok ID": DokumentID,
                "Bilag til Dok ID": BilagTilDok,
                "Bilag": DokBilag,
                "Omfattet af aktindsigt?": Omfattet,
                "Gives der aktindsigt?": Aktstatus,
                "Begrundelse hvis Nej/Delvis": Begrundelse,
                "IsDocumentPDF": IsDocumentPDF,
            }

            # Append the row to the DataFrame
            dt_AktIndex = pd.concat([dt_AktIndex, pd.DataFrame([row_to_add])], ignore_index=True)


        # Sort the DataFrame by the column "Akt ID" in ascending order
        dt_AktIndex = dt_AktIndex.sort_values(by="Akt ID", ascending=True)

        # Reset the index (optional, to clean up the index after sorting)
        dt_AktIndex = dt_AktIndex.reset_index(drop=True)


        # Initialize an empty list
        ListOfNonPDFDocs = []

        # Iterate through the DataFrame rows
        for _, row in dt_AktIndex.iterrows():  # Assuming dt_AktIndex is the DataFrame
            if row["IsDocumentPDF"] is not True:  # Check if the row's "IsDocumentPDF" is False
                # Add the "Filnavn" to the list if "IsDocumentPDF" is False
                ListOfNonPDFDocs.append(row["Filnavn"])

        # Check if ListOfNonPDFDocs is empty or not
        if not ListOfNonPDFDocs:  # This checks if the list is None or has no elements
            print("Listen er tom")
        else:
            # Initialize FinalString
            FinalString = ""

            # Iterate through the list and format the rows
            for currentText in ListOfNonPDFDocs:
                FormattedRow = currentText + "<br><br>"  # Format each item
                FinalString += FormattedRow  # Concatenate to FinalString

            #Henter delingslink til Sharepoint
            credentials = UserCredential(RobotUserName, RobotPassword)
            ctx = ClientContext(SharePointURL).with_credentials(credentials)

            # Define the server-relative URL of the folder or file
            folder_or_file_url = f"/Teams/tea-teamsite10506/Delte Dokumenter/Aktindsigter/{Overmappe}/{Undermappe}"  
            target_item = ctx.web.get_folder_by_server_relative_url(folder_or_file_url)  # Use get_file_by_server_relative_url for files

            try:
                # Share a folder or file link (Organization-only access with View permissions)
                result = target_item.share_link(SharingLinkKind.OrganizationView).execute_query()
                print("Sharing link created successfully!")
                link_url = result.value.sharingLinkInfo.Url

                # Verify the sharing link type
                result = Web.get_sharing_link_kind(ctx, link_url).execute_query()
                sharing_kind = result.value
                sharing_messages = {
                    2: "Organization view access link",
                    3: "Organization edit access link"
                }
                print(sharing_messages.get(sharing_kind, "Unknown sharing link kind"))

                # Optional: Unshare the link
                # Uncomment this if you want to remove the sharing link later
                # target_item.unshare_link(SharingLinkKind.OrganizationView).execute_query()
                # print("Sharing link unshared successfully!")

            except Exception as e:
                print(f"Error: {e}")

                # ---- Send mail til sagsansvarlig ----
        
            # Define email details
            sender = "Aktbob<rpamtm001@aarhus.dk>" # Replace with actual sender
            subject = f"Fil kan ikke konverteres til PDF - {Sagsnummer}"
            body = (
                "Kære Sagsbehandler,<br><br>"
                "Følgende dokumenter kunne ikke konverteres til PDF:<br><br>"
                f"{FinalString}"
                "Dokumenterne er blevet uploaded til sharepoint mappen: "
                f'<a href="{link_url}">SharePoint</a><br><br>'
                "Kontroller venligst manuelt dokumenterne.<br><br>"
                "Med venlig hilsen<br><br>"
                "Teknik & Miljø<br><br>"
                "Digitalisering<br><br>"
                "Aarhus Kommune"
            )
            smtp_server = "smtp.adm.aarhuskommune.dk"   # Replace with your SMTP server
            smtp_port = 25                    # Replace with your SMTP port

            # Call the send_email function
            send_email(
                receiver=UdviklerMailAktbob,
                sender=sender,
                subject=subject,
                body=body,
                smtp_server=smtp_server,
                smtp_port=smtp_port,
                html_body=True
            )

        #Fjerner IsdocumentPDF fra datatabellen
        dt_AktIndex = dt_AktIndex.drop('IsDocumentPDF',axis=1)
    
        # Base path (replace with your actual path)
        base_path = os.path.join("C:\\", "Users", os.getlogin(), "Downloads")

        # Iterate through the rows of the DataFrame and delete the files
        for _, row in dt_AktIndex.iterrows():
            file_name = row['Filnavn']
            file_path = os.path.join(base_path, file_name)

            try:
                if os.path.exists(file_path):
                    if os.path.isfile(file_path):  # Check if it's a file
                        os.remove(file_path)
                        print(f"Deleted file: {file_path}")
                    elif os.path.isdir(file_path):  # Check if it's a directory
                        shutil.rmtree(file_path, ignore_errors=True)
                        print(f"Deleted directory: {file_path}")
                else:
                    print(f"File not found: {file_path}")
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

    #Det er en nova sag
    else:
        #Det er en Nova sag
        print("Det er en Nova sag")
        for index, row in dt_DocumentList.iterrows():
            # Convert items to strings unless they are explicitly integers
            Omfattet = str(row["Omfattet af ansøgningen? (Ja/Nej)"])
            DokumentID = str(row["Dok ID"])
            
            # Handle AktID conversion
            AktID = row['Akt ID']
            if isinstance(AktID, str):  
                AktID = int(AktID.replace('.', ''))
            elif isinstance(AktID, int):  
                AktID = AktID

            Titel = str(row["Dokumenttitel"])
            BilagTilDok = str(row["Bilag til Dok ID"])
            DokBilag = str(row["Bilag"])
            Dokumentkategori = str(row["Dokumentkategori"])
            Aktstatus = str(row["Gives der aktindsigt i dokumentet? (Ja/Nej/Delvis)"])
            Begrundelse = str(row["Begrundelse hvis nej eller delvis"])
            #Skal slettes: 
                # Check if it's a Timestamp and convert it correctly
            Dokumentdato =row['Dokumentdato']
            if isinstance(Dokumentdato, pd.Timestamp):
                Dokumentdato = Dokumentdato.strftime("%d-%m-%Y")
            else:
                Dokumentdato = datetime.strptime(Dokumentdato, "%Y-%m-%d").strftime("%d-%m-%Y")
            #Dokumentdato = datetime.strptime(row["Dokumentdato"], "%d-%m-%Y").strftime("%d-%m-%Y")
            IsDocumentPDF = True
            print(Dokumentdato)
            print(f"AktID til debug: {AktID}")

            # Declare the necessary variables
            base_path = "Teams/tea-teamsite10506/Delte dokumenter/Aktindsigter/"

            # Sanitize the title
            Titel = sanitize_title(Titel)

            Titel = calculate_available_title_length(base_path, Overmappe, Undermappe, AktID, DokumentID, Titel)

            print("Adjusted Title:", Titel)

            if (("ja" in Aktstatus.lower() or "delvis" in Aktstatus.lower()) 
                and DokumentID != "" 
                and "ja" in Omfattet.lower()):
                
                print("Henter dokument information")
                print(DokumentID)
                TransactionID = str(uuid.uuid4())
                url = f"{KMDNovaURL}/Document/GetList?api-version=2.0-Case"

                headers = {
                    "Authorization": f"Bearer {KMD_access_token}",
                    "Content-Type": "application/json"
                }

                payload = {
                    "common": {
                        "transactionId": TransactionID,
                        #"uuid": DokumentID ## skal hente dokument id og ikke dokumentnr, find ud af hvornår den skal hentes.
                    },
                    "paging": {
                        "startRow": 1,
                        "numberOfRows": 100
                    },
                    "documentNumber": DokumentID,
                    "caseNumber": Sagsnummer,
                    "getOutput": {
                        "documentDate": True,
                        "title": True,
                        "fileExtension": True
                        }
                    }

                try:
                    response = requests.put(url, headers=headers, json=payload)
                    if response.status_code == 200:
                        print(response.status_code)
                    else:
                        print("Failed to fetch Sagstitel from NOVA. Status Code:", response.status_code)
                except Exception as e:
                    print("Failed to fetch Sagstitel (Nova):", str(e))

                DokumentType = response.json()["documents"][0]["fileExtension"]
                FileSize = response.json()["documents"][0]["fileSize"]
                DocumentUuid = response.json()["documents"][0]["documentUuid"]
                print(DokumentType)
                
                #Downloader file
                
                TransactionID = str(uuid.uuid4())
                url = f"{KMDNovaURL}/Document/GetFile?api-version=2.0-Case"
                file_path = os.path.join("C:\\Users", os.getlogin(), "Downloads", f"{DokumentID}.{DokumentType}")

                headers = {
                    "Authorization": f"Bearer {KMD_access_token}",
                    "Content-Type": "application/json"
                }

                payload = {
                    "common": {
                        "transactionId": TransactionID,
                        "uuid": DocumentUuid
                    }
                }

                try:
                    # Send request to API (Use GET if API expects it; otherwise, use POST)
                    response = requests.put(url, headers=headers, json=payload)

                    if response.status_code == 200:
                        # Save the entire file directly without chunking
                        with open(file_path, "wb") as file:
                            file.write(response.content)
                        
                        print(f"File successfully saved at: {file_path}")
                    else:
                        print("Failed to fetch file from NOVA. Status Code:", response.status_code)
                        print("Response:", response.text)  # Print error message from API

                except Exception as e:
                    print("Failed to fetch file from NOVA:", str(e))
                
                
                CanDocumentBeConverted = False
                conversionPossible = False

                # List of supported file extensions
                supported_extensions = [
                    "bmp", "csv", "doc", "docm", "dwf", "dwg", "dxf", "emf", "eml",
                    "epub", "fodt", "gif", "htm", "html", "ico", "jpeg", "jpg", "msg",
                    "odp", "ods", "odt", "pdf", "png", "pos", "pps", "ppt", "pptx", "psd",
                    "rtf", "tif", "tiff", "tsv", "txt", "vdw", "vdx", "vsd", "vss", "vst",
                    "vsx", "vtx", "webp", "wmf", "xls", "xlsm", "xlsx", "xltx", "heic","docx"
                ]
                # Check if the input file extension exists in the list
                if DokumentType.lower() in supported_extensions:
                    CanDocumentBeConverted = True
                else:
                    CanDocumentBeConverted = False

                if CanDocumentBeConverted:
                    print("Filen skal ikke konverteres")

                else:
                    print("Filen skal konverteres - attempting cloudconvert")
    
                    # Create the API URL
                    url = f"https://api.cloudconvert.com/v2/convert/formats?filter[input_format]={DokumentType}&filter[output_format]=pdf&filter[operation]=convert"

                    # Add the authorization header
                    headers = {
                        "Authorization": CloudConvertAPI
                    }

                    # Initialize conversionPossible
                    conversionPossible = False

                    # Execute the request
                    response = requests.get(url, headers=headers)

                    # Process the response
                    if response.status_code == 200 and response.text.strip():
                        # Parse the response content to a dictionary
                        jsonResponse = json.loads(response.text)
                        
                        # Check if the data array contains any elements
                        data = jsonResponse.get("data", [])
                        if data:
                            # Iterate through each conversion option
                            for item in data:
                                operation = item.get("operation")
                                inputFormatCheck = item.get("input_format")
                                outputFormat = item.get("output_format")
                                
                                # Check if it matches the desired conversion
                                if operation == "convert" and inputFormatCheck == DokumentType and outputFormat == "pdf":
                                    conversionPossible = True
                                    break
                    
                    if not conversionPossible:
                        print(f"Skipping cause CloudConvert doesn't support: {DokumentType}->PDF")
                        ByteResult = bytes()                  
                        #Skal der sættes en bolean value?
                    else:
                        print("Conversion is supported!")
                                                
                        # Step 1: Create the job with an import/upload task
                        create_job_url = "https://api.cloudconvert.com/v2/jobs"
                        create_job_headers = {
                            "Authorization": CloudConvertAPI,
                            "Content-Type": "application/json",
                        }
                        json_body = {
                            "tasks": {
                                "import_1": {
                                    "operation": "import/upload"
                                },
                            },
                            "tag": f"Aktbob-{DokumentID}-{time.strftime('%H-%M-%S')}",
                        }
                        # Send the request to create the job
                        response = requests.post(create_job_url, headers=headers, json=json_body)
                        job_response_data = response.json()


                        # Extract upload URL and parameters
                        tasks = job_response_data.get("data", {}).get("tasks", [])
                        if not tasks:
                            print("Error: No tasks found in job creation response.")
                            exit(1)

                        # Find the first task containing result.form
                        upload_url = None
                        upload_parameters = None
                        upload_task_id = None

                        for task in tasks:
                            if task.get("operation") == "import/upload" and "result" in task:
                                form = task["result"].get("form")
                                if form:
                                    upload_url = form.get("url")
                                    upload_parameters = form.get("parameters", {})  # Ensure it's a dictionary
                                    upload_task_id = task["id"]
                                    break  # Found the required task

                        # Validate extracted data
                        if not upload_url or not upload_parameters:
                            print("Error: Could not retrieve upload URL or parameters.")
                            exit(1)

                        # ------------------- STEP 2: UPLOAD FILE -------------------

                        # Convert upload_parameters dictionary to key-value pairs
                        upload_data = {key: value for key, value in upload_parameters.items()} 

                        # Perform file upload to AWS S3 (CloudConvert Storage)
                        with open(file_path, "rb") as file:
                            upload_files = {"file": file}
                            upload_response = requests.post(upload_url, data=upload_data, files=upload_files)

                        # Check upload response
                        if upload_response.status_code == 201:  # AWS S3 responds with 201 on success
                            print("File uploaded successfully!")
                        else:
                            print(f"Upload failed: {upload_response.status_code} - {upload_response.text}")
                            exit(1)

                        # Remove the original file after upload
                        os.remove(file_path)

                        # Step 3: Create convert and export tasks
                        convert_export_body = {
                            "tasks": {
                                "convert_1": {
                                    "operation": "convert",
                                    "input": [upload_task_id],
                                    "input_format": DokumentType,
                                    "output_format": "pdf",
                                },
                                "export_1": {
                                    "operation": "export/url",
                                    "input": ["convert_1"],
                                }
                            },
                            "tag": f"Aktbob-{DokumentID}-{time.strftime('%H-%M-%S')}",
                        }
                        convert_export_response = requests.post(
                            create_job_url, headers=create_job_headers, json=convert_export_body
                        )
                        convert_export_response_data = convert_export_response.json()


                        if "INVALID_CONVERSION_TYPE" not in convert_export_response.text:
                            export_task_id = convert_export_response_data["data"]["tasks"][1]["id"]

                            # Step 4: Check export task status
                            while True:
                                status_check_url = f"https://api.cloudconvert.com/v2/tasks/{export_task_id}"
                                status_check_response = requests.get(status_check_url, headers=create_job_headers)
                                status_check_data = status_check_response.json()

                                task_status = status_check_data["data"]["status"]

                                if task_status == "finished":
                                    # Extract the download URL
                                    download_url = status_check_data["data"]["result"]["files"][0]["url"]

                                    # Download the file
                                    file_path = os.path.join("C:\\Users", os.getlogin(), "Downloads", f"{DokumentID}.pdf")
                                    #file_path = "Output.pdf"
                                    with requests.get(download_url, stream=True) as r:
                                        with open(file_path, "wb") as file:
                                            for chunk in r.iter_content(chunk_size=8192):
                                                file.write(chunk)

                                    print("File downloaded successfully.")

                                    # Read the file into ByteResult
                                    with open(file_path, "rb") as file:
                                        ByteResult = file.read()

                                    break
                                elif task_status not in ["waiting", "processing"]:
                                    print("An error occurred:", status_check_response.text)
                                    ByteResult = b""
                                    break

                                time.sleep(5)  # Wait for 5 seconds before checking again


                if conversionPossible or CanDocumentBeConverted:
                    
                    upload_to_filarkiv(FilarkivURL,FilarkivCaseID, Filarkiv_access_token, AktID, DokumentID,Titel, file_path)


                else: 
                    print("Could not be converted or uploaded - uploading directly to SharePoint")
                    IsDocumentPDF = False 
                    upload_file_to_sharepoint(
                        site_url=SharePointURL,
                        overmappe=Overmappe,
                        undermappe=Undermappe,
                        file_path=file_path,
                        sharepoint_app_id=SharePointAppID,
                        sharepoint_tenant=SharePointTenant,
                        robot_username=RobotUserName,
                        robot_password=RobotPassword
                    )
                Titel = f"{AktID:04} - {DokumentID} - {Titel}"
            
            else:
                print("Dokumentet skal ikke med i ansøgningen")
                Titel = f"{AktID:04} - {DokumentID} - {Titel}"

    

            # Parse and prepare data for the row
            row_to_add = {
                "Akt ID": int(AktID),
                "Filnavn": Titel,
                "Dokumentkategori": Dokumentkategori,
                "Dokumentdato": datetime.strptime(Dokumentdato, "%d-%m-%Y"),
                "Dok ID": DokumentID,
                "Bilag til Dok ID": BilagTilDok,
                "Bilag": DokBilag,
                "Omfattet af aktindsigt?": Omfattet,
                "Gives der aktindsigt?": Aktstatus,
                "Begrundelse hvis Nej/Delvis": Begrundelse,
                "IsDocumentPDF": IsDocumentPDF,
            }

            # Append the row to the DataFrame
            dt_AktIndex = pd.concat([dt_AktIndex, pd.DataFrame([row_to_add])], ignore_index=True)


        # Sort the DataFrame by the column "Akt ID" in ascending order
        dt_AktIndex = dt_AktIndex.sort_values(by="Akt ID", ascending=True)

        # Reset the index (optional, to clean up the index after sorting)
        dt_AktIndex = dt_AktIndex.reset_index(drop=True)


        # Initialize an empty list
        ListOfNonPDFDocs = []

        # Iterate through the DataFrame rows
        for _, row in dt_AktIndex.iterrows():  # Assuming dt_AktIndex is the DataFrame
            if row["IsDocumentPDF"] is not True:  # Check if the row's "IsDocumentPDF" is False
                # Add the "Filnavn" to the list if "IsDocumentPDF" is False
                ListOfNonPDFDocs.append(row["Filnavn"])

        # Check if ListOfNonPDFDocs is empty or not
        if not ListOfNonPDFDocs:  # This checks if the list is None or has no elements
            print("Listen er tom")
        else:
            # Initialize FinalString
            FinalString = ""

            # Iterate through the list and format the rows
            for currentText in ListOfNonPDFDocs:
                FormattedRow = currentText + "<br><br>"  # Format each item
                FinalString += FormattedRow  # Concatenate to FinalString

            #Henter delingslink til Sharepoint
            credentials = UserCredential(RobotUserName, RobotPassword)
            ctx = ClientContext(SharePointURL).with_credentials(credentials)

            # Define the server-relative URL of the folder or file
            folder_or_file_url = f"/Teams/tea-teamsite10506/Delte Dokumenter/Aktindsigter/{Overmappe}/{Undermappe}"  
            target_item = ctx.web.get_folder_by_server_relative_url(folder_or_file_url)  # Use get_file_by_server_relative_url for files

            try:
                # Share a folder or file link (Organization-only access with View permissions)
                result = target_item.share_link(SharingLinkKind.OrganizationView).execute_query()
                print("Sharing link created successfully!")
                link_url = result.value.sharingLinkInfo.Url

                # Verify the sharing link type
                result = Web.get_sharing_link_kind(ctx, link_url).execute_query()
                sharing_kind = result.value
                sharing_messages = {
                    2: "Organization view access link",
                    3: "Organization edit access link"
                }
                print(sharing_messages.get(sharing_kind, "Unknown sharing link kind"))

                # Optional: Unshare the link
                # Uncomment this if you want to remove the sharing link later
                # target_item.unshare_link(SharingLinkKind.OrganizationView).execute_query()
                # print("Sharing link unshared successfully!")

            except Exception as e:
                print(f"Error: {e}")

                # ---- Send mail til sagsansvarlig ----
        
            # Define email details
            sender = "Aktbob<rpamtm001@aarhus.dk>" # Replace with actual sender
            subject = f"Fil kan ikke konverteres til PDF - {Sagsnummer}"
            body = (
                "Kære Sagsbehandler,<br><br>"
                "Følgende dokumenter kunne ikke konverteres til PDF:<br><br>"
                f"{FinalString}"
                "Dokumenterne er blevet uploaded til sharepoint mappen: "
                f'<a href="{link_url}">SharePoint</a><br><br>'
                "Kontroller venligst manuelt dokumenterne.<br><br>"
                "Med venlig hilsen<br><br>"
                "Teknik & Miljø<br><br>"
                "Digitalisering<br><br>"
                "Aarhus Kommune"
            )
            smtp_server = "smtp.adm.aarhuskommune.dk"   # Replace with your SMTP server
            smtp_port = 25                    # Replace with your SMTP port

            # Call the send_email function
            send_email(
                receiver=UdviklerMailAktbob,
                sender=sender,
                subject=subject,
                body=body,
                smtp_server=smtp_server,
                smtp_port=smtp_port,
                html_body=True
            )

        #Fjerner IsdocumentPDF fra datatabellen
        dt_AktIndex = dt_AktIndex.drop('IsDocumentPDF',axis=1)
    
        # Base path (replace with your actual path)
        base_path = os.path.join("C:\\", "Users", os.getlogin(), "Downloads")

        # Iterate through the rows of the DataFrame and delete the files
        for _, row in dt_AktIndex.iterrows():
            file_name = row['Filnavn']
            file_path = os.path.join(base_path, file_name)

            try:
                if os.path.exists(file_path):
                    if os.path.isfile(file_path):  # Check if it's a file
                        os.remove(file_path)
                        print(f"Deleted file: {file_path}")
                    elif os.path.isdir(file_path):  # Check if it's a directory
                        shutil.rmtree(file_path, ignore_errors=True)
                        print(f"Deleted directory: {file_path}")
                else:
                    print(f"File not found: {file_path}")
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")



    return {
    "out_dt_AktIndex": dt_AktIndex,
    }
