from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

def invoke_PrepareEachDocumentToUpload(Arguments_PrepareEachDocumentToUpload, orchestrator_connection: OrchestratorConnection):
    import pandas as pd
    import re
    import requests
    from requests_ntlm import HttpNtlmAuth
    import json
    import os
    import time
    from datetime import datetime
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.sharing.links.kind import SharingLinkKind
    from office365.sharepoint.webs.web import Web
    from office365.runtime.auth.user_credential import UserCredential
    import json
    from SendSMTPMail import send_email
    import shutil
    import uuid
    from SharePointUploader import upload_file_to_sharepoint
    import re
    import mimetypes


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
    DeskProID = Arguments_PrepareEachDocumentToUpload.get("in_DeskProID")
    DeskProTitel = Arguments_PrepareEachDocumentToUpload.get("in_DeskProTitel")
    tenant = Arguments_PrepareEachDocumentToUpload.get("tenant")
    client_id = Arguments_PrepareEachDocumentToUpload.get("client_id")
    thumbprint = Arguments_PrepareEachDocumentToUpload.get("thumbprint")
    cert_path = Arguments_PrepareEachDocumentToUpload.get("cert_path")

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

    def upload_to_filarkiv(FilarkivURL, FilarkivCaseID, Filarkiv_access_token, AktID, DokumentID, Titel, file_path, DokumentType):
        DoesFolderExists = False
        Filarkiv_DocumentID = None  # Ensure it is initialized
        FileName = f"{AktID:04} - {DokumentID} - {Titel}"
        url = f"{FilarkivURL}/Documents/CaseDocumentOverview?caseId={FilarkivCaseID}&pageIndex=1&pageSize=500"

        headers = {
            "Authorization": f"Bearer {Filarkiv_access_token}",
            "Content-Type": "application/xml"
        }
        
        try:
            response = requests.get(url, headers=headers)
            orchestrator_connection.log_info(f"FilArkiv respons: {response.status_code}")
            
            if response.status_code == 200:
                response_json = response.json()
                
                if not response_json:
                    orchestrator_connection.log_info("Der findes ingen dokumenter på sagen")
                    DocumentDate = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                    DocumentNumber = 1
                    data = {
                        "caseId": FilarkivCaseID,
                        "securityClassificationLevel": 0,
                        "title": FileName,
                        "documentNumber": DocumentNumber,
                        "documentDate": DocumentDate,
                        "direction": 0,
                        "documentReference": DokumentID
                    }
                    response = requests.post("https://core.filarkiv.dk/api/v1/Documents", headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, data=json.dumps(data))  
                    if response.status_code in [200, 201]:
                        Filarkiv_DocumentID = response.json().get("id")
                    else:
                        orchestrator_connection.log_info(f"Failed to create document. Response:{response.text}")
                else:
                    for current_item in response_json:
                        if FileName in current_item.get("title", ""):
                            orchestrator_connection.log_info("Dokument Mappen er oprettet")
                            Filarkiv_DocumentID = current_item.get("id")
                            DoesFolderExists = True
                            break  # Exit loop once found
                    if not DoesFolderExists:
                        orchestrator_connection.log_info("Finder det nye dokumentnummer")
                        HighestDocumentNumber = max((int(i.get("documentNumber", 0)) for i in response_json), default=1)
                        DocumentNumber = HighestDocumentNumber + 1
                        DocumentDate = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                        data = {
                            "caseId": FilarkivCaseID,
                            "securityClassificationLevel": 0,
                            "title": FileName,
                            "documentNumber": DocumentNumber,
                            "documentDate": DocumentDate,
                            "direction": 0,
                            "documentReference": DokumentID
                        }
                        response = requests.post("https://core.filarkiv.dk/api/v1/Documents", headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, data=json.dumps(data))
                        if response.status_code in [200, 201]:
                            Filarkiv_DocumentID = response.json().get("id")
                            orchestrator_connection.log_info(f"Anvender følgende Filarkiv_DocumentID: {Filarkiv_DocumentID}")
                        else:
                            orchestrator_connection.log_info(f"Failed to create document. Response: {response.text}")
            else:
                orchestrator_connection.log_info(f"Failed to fetch data, status code: {response.status_code}")
        except Exception as e:
            raise Exception("Kunne ikke hente dokumentinformation:", str(e))
        
        if Filarkiv_DocumentID is None:
            orchestrator_connection.log_info("Fejl: Filarkiv_DocumentID blev ikke genereret. Afbryder processen.")
            return
        
        if not DoesFolderExists:
            extension = f".{DokumentType}"
            mime_type = {
                ".txt": "text/plain", ".pdf": "application/pdf", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
                ".gif": "image/gif", ".doc": "application/msword", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xls": "application/vnd.ms-excel", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".csv": "text/csv",
                ".json": "application/json", ".xml": "application/xml"
            }.get(extension, "application/octet-stream")
            FileName += extension
            orchestrator_connection.log_info(f"Anvender følgende dokumentID: {Filarkiv_DocumentID}")
            response = requests.post("https://core.filarkiv.dk/api/v1/Files", headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, json={"documentId": Filarkiv_DocumentID, "fileName": FileName, "sequenceNumber": 0, "mimeType": mime_type})
            if response.status_code in [200, 201]:
                FileID = response.json().get('id')
                orchestrator_connection.log_info(f"FileID: {FileID}")
            else:
                orchestrator_connection.log_info(f"Failed to create file metadata. {response.text}")
                return False
            
            url = f"https://core.filarkiv.dk/api/v1/FileIO/Upload/{FileID}"
            if not os.path.exists(file_path):
                orchestrator_connection.log_info(f"Error: File not found at {file_path}")
            else:
                with open(file_path, 'rb') as file:
                    files = [('file', (FileName, file, mime_type))]
                    response = requests.post(url, headers={"Authorization": f"Bearer {Filarkiv_access_token}"}, files=files)
                    if response.status_code in [200, 201]:
                        orchestrator_connection.log_info("File uploaded successfully.")
                    else:
                        orchestrator_connection.log_info(f"Failed to upload file. Status Code: {response.status_code}")
                        return False

                    #Sætter den høje prioritet på dokumentet
                    url = f"https://core.filarkiv.dk/api/v1/FileProcess/UpdatePriority"

                    data = {
                            "fileId": FileID,
                            "priority": 10000
                    }
                    response = requests.post(url, headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, data=json.dumps(data))
                    if response.status_code in [200, 201]:
                        orchestrator_connection.log_info("Det lykkedes at opdaterer prioriteten")
                    else:
                        orchestrator_connection.log_info(f"Fejlede i prioritering: {response.text}")
            return True
                


    def check_conversion_possible(dokument_type, cloudconvert_api):
        orchestrator_connection.log_info("Filen skal konverteres - attempting CloudConvert")
        
        url = f"https://api.cloudconvert.com/v2/convert/formats?filter[input_format]={dokument_type}&filter[output_format]=pdf&filter[operation]=convert"
        
        headers = {
            "Authorization": cloudconvert_api
        }
        
        conversion_possible = False
        
        try:
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200 and response.text.strip():
                json_response = json.loads(response.text)
                
                data = json_response.get("data", [])
                if data:
                    for item in data:
                        if (item.get("operation") == "convert" and
                            item.get("input_format") == dokument_type and
                            item.get("output_format") == "pdf"):
                            conversion_possible = True
                            break
        except requests.RequestException as e:
            raise Exception(f"Error during API request: {e}")
        
        return conversion_possible

    def convert_file_to_pdf(CloudConvertAPI, file_path, DokumentID, DokumentType,Titel, AktID):
        orchestrator_connection.log_info("Conversion is supported!")
        
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
        
        tasks = job_response_data.get("data", {}).get("tasks", [])
        if not tasks:
            orchestrator_connection.log_info("Error: No tasks found in job creation response.")
            return None
        
        upload_url, upload_parameters, upload_task_id = None, None, None
        for task in tasks:
            if task.get("operation") == "import/upload" and "result" in task:
                form = task["result"].get("form")
                if form:
                    upload_url = form.get("url")
                    upload_parameters = form.get("parameters", {})
                    upload_task_id = task["id"]
                    break
        
        if not upload_url or not upload_parameters:
            orchestrator_connection.log_info("Error: Could not retrieve upload URL or parameters.")
            return None
        
        upload_data = {key: value for key, value in upload_parameters.items()}
        
        with open(file_path+f'.{DokumentType}', "rb") as file:
            upload_files = {"file": file}
            upload_response = requests.post(upload_url, data=upload_data, files=upload_files)
        
        if upload_response.status_code == 201:
            orchestrator_connection.log_info("File uploaded successfully!")
        else:
            orchestrator_connection.log_info(f"Upload failed: {upload_response.status_code} - {upload_response.text}")
            return None
        
        os.remove(file_path + f'.{DokumentType}')
        
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
        
        if "INVALID_CONVERSION_TYPE" in convert_export_response.text:
            orchestrator_connection.log_info("Error: Invalid conversion type.")
            return None
        
        export_task_id = convert_export_response_data["data"]["tasks"][1]["id"]
        
        while True:
            status_check_url = f"https://api.cloudconvert.com/v2/tasks/{export_task_id}"
            status_check_response = requests.get(status_check_url, headers=create_job_headers)
            status_check_data = status_check_response.json()
            
            task_status = status_check_data["data"]["status"]
            
            if task_status == "finished":
                download_url = status_check_data["data"]["result"]["files"][0]["url"]
                
                file_path = os.path.join("C:\\Users", os.getlogin(), "Downloads", f"{AktID:04} - {DokumentID} - {Titel}.pdf")
                with requests.get(download_url, stream=True) as r:
                    with open(file_path, "wb") as file:
                        for chunk in r.iter_content(chunk_size=8192):
                            file.write(chunk)
                
                orchestrator_connection.log_info(f"File downloaded successfully at: {file_path}")
                
                return file_path
            elif task_status not in ["waiting", "processing"]:
                orchestrator_connection.log_info(f"An error occurred:{status_check_response.text}")
                return None
            
            time.sleep(5)

    def process_documents(
        dt_AktIndex,
        AktID,
        Titel,
        Dokumentkategori,
        Dokumentdato,
        DokumentID,
        BilagTilDok,
        DokBilag,
        Omfattet,
        Aktstatus,
        Begrundelse,
        IsDocumentPDF
    ):
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
        
        dt_AktIndex = pd.concat([dt_AktIndex, pd.DataFrame([row_to_add])], ignore_index=True)

        # Sort and reset index
        dt_AktIndex = dt_AktIndex.sort_values(by="Akt ID", ascending=True).reset_index(drop=True)

        #  1. Prepare base path for deletion
        base_path = os.path.join("C:\\", "Users", os.getlogin(), "Downloads")

        #  2. Create a list of non-PDF documents (KEEPING extensions for return)
        ListOfNonPDFDocs = dt_AktIndex.loc[dt_AktIndex["IsDocumentPDF"] != True, "Filnavn"].tolist()

        #  3. Loop through all files and process them correctly
        for index, row in dt_AktIndex.iterrows():
            file_name_with_extension = row["Filnavn"]  # Get original filename
            
            file_name_for_deletion = file_name_with_extension  # Default: use full name for deletion
            
            # #  If it's NOT a PDF, remove the file extension *just for deletion*
            # if not row["IsDocumentPDF"]:
            #     parts = file_name_with_extension.rsplit(".", 1)  # Split at last dot

            #     if len(parts) == 2:  # If there's an extension
            #         file_name_for_deletion, extension = parts  # Use filename without extension

            #  4. Delete ALL files (PDF and non-PDF)
            file_path = os.path.join(base_path, file_name_for_deletion)  # Use modified name for deletion
            #orchestrator_connection.log_info(f"Filepath to be deleted: {file_path}")
            try:
                if os.path.exists(file_path):
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path, ignore_errors=True)
                        orchestrator_connection.log_info(f"Deleted directory: {file_path}")
            except Exception as e:
                raise Exception(f"Error deleting {file_path}: {e}")

        # Return dt_AktIndex and ListOfNonPDFDocs, both WITH extensions
        return dt_AktIndex, ListOfNonPDFDocs

    def fetch_document_info(DokumentID, session, AktID, Titel):
        url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/Data/{DokumentID}"
        response = session.get(url)
        DocumentData = response.text
        data = json.loads(DocumentData)
        item_properties = data.get("ItemProperties", "")
        file_type_match = re.search(r'ows_File_x0020_Type="([^"]+)"', item_properties)
        version_ui_match = re.search(r'ows__UIVersionString="([^"]+)"', item_properties)
        DokumentType = file_type_match.group(1) if file_type_match else "Not found"
        VersionUI = version_ui_match.group(1) if version_ui_match else "Not found"
        Feedback = " "
        file_path = os.path.join(
            "C:\\Users",
            os.getenv("USERNAME"),
            "Downloads",
            f"{AktID:04} - {DokumentID} - {Titel}"
        )
        return {"DokumentType": DokumentType, "VersionUI": VersionUI, "Feedback": Feedback, "file_path": file_path}

    def download_file(file_path, ByteResult, DokumentID, GoUsername, GoPassword):
        try:
            with open(file_path, "wb") as file:
                file.write(ByteResult)
            orchestrator_connection.log_info("File written successfully.")
        except Exception as initial_exception:
            orchestrator_connection.log_info(f"Failed, trying from URL: {DokumentID} Path: {file_path}")
            orchestrator_connection.log_info(initial_exception)

            ByteResult = bytes()

            try:
                max_retries = 2
                for attempt in range(max_retries):
                    try:
                        metadata_url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/MetadataWithSystemFields/{DokumentID}"
                        metadata_response = requests.get(
                            metadata_url,
                            auth=HttpNtlmAuth(GoUsername, GoPassword),
                            headers={"Content-Type": "application/json"},
                            timeout=60
                        )
                        
                        content = metadata_response.text
                        DocumentURL = content.split("ows_EncodedAbsUrl=")[1].split('"')[1]
                        DocumentURL = DocumentURL.split("\\")[0].replace("go.aarhus", "ad.go.aarhus")
                        orchestrator_connection.log_info(f"Document URL: {DocumentURL}")
                        
                        handler = requests.Session()
                        handler.auth = HttpNtlmAuth(GoUsername, GoPassword)
                        with handler.get(DocumentURL, stream=True) as download_response:
                            download_response.raise_for_status()
                            with open(file_path, "wb") as file:
                                for chunk in download_response.iter_content(chunk_size=8192):
                                    file.write(chunk)

                        orchestrator_connection.log_info("File downloaded successfully.")
                        break
                    except Exception as retry_exception:
                        orchestrator_connection.log_info(f"Retry {attempt + 1} failed: {retry_exception}")
                        if attempt == max_retries - 1:
                            raise RuntimeError(
                                f"Failed to download file after {max_retries} retries. "
                                f"DokumentID: {DokumentID}, Path: {file_path}"
                            )
                        time.sleep(5)

            except RuntimeError as nested_exception:
                orchestrator_connection.log_info(f"An unrecoverable error occurred: {nested_exception}")
                raise nested_exception

    def fetch_document_bytes(session, DokumentID, file_path=None, dokument_type=None, max_retries=30, retry_interval=5, delete_after_use=False):

        url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/DocumentBytes/{DokumentID}"
        ByteResult = None

        for attempt in range(max_retries):
            try:
                response = session.get(url, timeout=180)

                if response.status_code == 200:
                    ByteResult = response.content
                    orchestrator_connection.log_info(f"Success! ByteResult size: {len(ByteResult)} bytes")
                    break
                else:
                    orchestrator_connection.log_info(f"Attempt {attempt + 1}: Failed with status code {response.status_code}")
            except Exception as e:
                orchestrator_connection.log_info(f"Attempt {attempt + 1}: Exception occurred - {e}")

            time.sleep(retry_interval)
        else:
            orchestrator_connection.log_info("Max retries reached. File download failed.")
            return None

        # If a file path is given, save to file
        if file_path:
            file_path_with_extension = f"{file_path}.{dokument_type}" if dokument_type else file_path
            with open(file_path_with_extension, "wb") as file:
                file.write(ByteResult)

            # If delete_after_use is True, remove the file
            if delete_after_use:
                os.remove(file_path_with_extension)
                orchestrator_connection.log_info("File deleted after use.")

        return ByteResult


    # Create an empty DataFrame with the defined structure
    dt_AktIndex = pd.DataFrame(dt_AktIndex)
    dt_non_pdf_docs = []

    # ---- If-statement som tjekker om det er en GeoSag eller NovaSag ----
    if GeoSag == True:
        #Sagen er en geo sag 
        #dt_DocumentList['Dokumentdato'] = pd.to_datetime(dt_DocumentList['Dokumentdato'], errors='coerce')
        dt_DocumentList['Dokumentdato'] = pd.to_datetime(dt_DocumentList['Dokumentdato'], format="%d-%m-%Y", errors='coerce')

        with requests.Session() as session:
            session.auth = HttpNtlmAuth(GoUsername, GoPassword)
            session.headers.update({"Content-Type": "application/json"}) 
        
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
            mimetypes.add_type("application/x-msmetafile", ".emz")
            # Split title into name and extension
            parts = Titel.rsplit('.', 1)  # Splits at the last dot
            if len(parts) == 2:
                name, ext = parts
                # Check if it's a known file extension
                if mimetypes.guess_type(f"file.{ext}")[0]:  
                    Titel = name  # Remove extension
            else:
                orchestrator_connection.log_info("No file extension detected.")

            BilagTilDok = str(row["Bilag til Dok ID"])
            DokBilag = str(row["Bilag"])
            Dokumentkategori = str(row["Dokumentkategori"])
            Aktstatus = str(row["Gives der aktindsigt i dokumentet? (Ja/Nej/Delvis)"])
            Begrundelse = str(row["Begrundelse hvis nej eller delvis"])
            Dokumentdato =row['Dokumentdato']
            if isinstance(Dokumentdato, pd.Timestamp):
                Dokumentdato = Dokumentdato.strftime("%d-%m-%Y")
            else:
                Dokumentdato = datetime.strptime(Dokumentdato, "%Y-%m-%d").strftime("%d-%m-%Y")
            
            IsDocumentPDF = True
            orchestrator_connection.log_info(f"AktID til debug: {AktID}")

            # Declare the necessary variables
            base_path = "Teams/tea-teamsite10506/Delte dokumenter/Aktindsigter/"

            # Sanitize the title
            Titel = sanitize_title(Titel)

            Titel = calculate_available_title_length(base_path, Overmappe, Undermappe, AktID, DokumentID, Titel)


            if (("ja" in Aktstatus.lower() or "delvis" in Aktstatus.lower()) 
                and DokumentID != "" 
                and "ja" in Omfattet.lower()):

                Metadata = fetch_document_info(DokumentID, session, AktID, Titel)
                
                # Extracting variables for further use in the loop
                DokumentType = Metadata["DokumentType"]
                DokumenttypeBackup = DokumentType
                VersionUI = Metadata["VersionUI"]
                Feedback = Metadata["Feedback"]
                file_path = Metadata["file_path"]
                FilIsPDF = False 
                CanDocumentBeConverted = False
                conversionPossible = False
                # Tjekker om Goref-fil
                if "goref" in DokumentType:
                    orchestrator_connection.log_info("Dokumenter er .GORef")
                    ByteResult = fetch_document_bytes(session, DokumentID, file_path, delete_after_use=False)

                    if ByteResult:
                        with open(file_path, "r", encoding="utf-8") as file:
                            RefDokument = file.read()
                        
                        refdocument = RefDokument.split("?docid=")[1]
                        DokumentID = refdocument.split('"')[0]
                
                        os.remove(file_path)
                        orchestrator_connection.log_info("File deleted after use.")
                    
                    orchestrator_connection.log_info(f"GorefDokID: {DokumentID}")
                    #Henter dokument data
                    Metadata = fetch_document_info(DokumentID, session, AktID, Titel)
                
                    # Extracting variables for further use in the loop
                    DokumentType = Metadata["DokumentType"]
                    orchestrator_connection.log_info(f"Dokumenttype gotten from goref {DokumentType}")
                    VersionUI = Metadata["VersionUI"]
                    Feedback = Metadata["Feedback"]
                    file_path = Metadata["file_path"]


                if DokumentType.lower() == "pdf": # Hvis PDF downloader den byte-filen
                    #Downloader fil fra GO    
                    orchestrator_connection.log_info("Allerede PDF - downloader")
                    ByteResult = fetch_document_bytes(session, DokumentID, max_retries=5, retry_interval=30)

                    if ByteResult:
                        orchestrator_connection.log_info(f"File size: {len(ByteResult)} bytes")
                    else:
                        orchestrator_connection.log_info("No file was downloaded.")
                    file_path = (f"{file_path}.pdf")
                    download_file(file_path, ByteResult, DokumentID, GoUsername, GoPassword) 
                    #file_path = (f"{file_path}.pdf") 
                    FilIsPDF = True 
                                                                                            
                else: #Dokumentet er ikke en pdf - forsøger at konverterer
                                      
                    # Forsøger med GO-conversion
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
                            orchestrator_connection.log_info(f"Error Message: {response.text}")
                        
                        # Feedback and byte result
                        Feedback = response.text
                        ByteResult = response.content
                        # Check if ByteResult is empty
                        if len(ByteResult) == 0:
                            orchestrator_connection.log_info(f"Status Code: {response.status_code}")
                        else:
                            orchestrator_connection.log_info("ByteResult received successfully.")
                        
                    except Exception as e:
                        orchestrator_connection.log_info(f"Could not convert to pdf : {e}")
                        ByteResult = ""
                    
                    
                    # tjekker om go-conversion lykkedes eller ej
                    if "Document could not be converted" in Feedback or len(ByteResult) == 0:
                        orchestrator_connection.log_info("Go-convervision mislykkedes forsøger med Filarkiv")

                        #Downloader fil fra GO    
                        FilnavnFørPdf = f"Output.{DokumentType}"
                        ByteResult = fetch_document_bytes(session, DokumentID, file_path=FilnavnFørPdf)
                        if ByteResult:
                            orchestrator_connection.log_info(f"File size: {len(ByteResult)} bytes")
                        else:
                            orchestrator_connection.log_info("No file was downloaded.")
                        
                        download_path = f"{file_path}.{DokumentType}"
                        download_file(download_path, ByteResult, DokumentID, GoUsername, GoPassword)  
                        
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
                            orchestrator_connection.log_info("Filen konverteres med Filarkiv")

                        else:
                            conversionPossible = check_conversion_possible(DokumentType, CloudConvertAPI)
                            
                            if not conversionPossible:
                                orchestrator_connection.log_info(f"Skipping cause CloudConvert doesn't support: {DokumentType}->PDF")
                                ByteResult = bytes()                  
                                #Skal der sættes en bolean value?
                            else:
                                orchestrator_connection.log_info("Forsøger med CloudConvert")
                                file_path = convert_file_to_pdf(CloudConvertAPI, file_path, DokumentID, DokumentType,Titel, AktID)
                                if file_path:
                                    orchestrator_connection.log_info(f"PDF saved at: {file_path}")
                                    DokumentType = "pdf"
                                                    
                    else: # Go-conversion lykkedes downloader fil
                        orchestrator_connection.log_info("Go-conversion lykkedes")
                        if ByteResult:
                            orchestrator_connection.log_info(f"File size: {len(ByteResult)} bytes")
                        else:
                            orchestrator_connection.log_info("No file was downloaded.")
                        
                        file_path = (f"{file_path}.pdf")
                        FilIsPDF = True
                        download_file(file_path, ByteResult, DokumentID, GoUsername, GoPassword)
     
                if FilIsPDF or conversionPossible or CanDocumentBeConverted:
                    success = upload_to_filarkiv(
                        FilarkivURL, FilarkivCaseID, Filarkiv_access_token,
                        AktID, DokumentID, Titel, file_path, DokumentType= DokumentType
                    )

                    if success:
                        DokumentType = "pdf"
                    else:
                        IsDocumentPDF = False
                        # File path her burde allerede have filtype
                        #file_path = f"{file_path}.{DokumentType}"
                        upload_file_to_sharepoint(
                            site_url=SharePointURL,
                            Overmappe=Overmappe,
                            Undermappe=Undermappe,
                            file_path=file_path,
                            RobotUserName=RobotUserName,
                            RobotPassword=RobotPassword,
                            tenant = tenant, 
                            client_id = client_id, 
                            thumbprint = thumbprint, 
                            cert_path = cert_path
                        )

                else:
                    # This is when conversion is not possible and upload shouldn't even be attempted to Filarkiv
                    IsDocumentPDF = False
                    #File path her burde allerede have filtype
                    #file_path = f"{file_path}.{DokumentType}"
                    upload_file_to_sharepoint(
                        site_url=SharePointURL,
                        Overmappe=Overmappe,
                        Undermappe=Undermappe,
                        file_path=file_path,
                        RobotUserName=RobotUserName,
                        RobotPassword=RobotPassword,
                        tenant = tenant, 
                        client_id = client_id, 
                        thumbprint = thumbprint, 
                        cert_path = cert_path
            )
    
            else:
                Titel = f"{AktID:04} - {DokumentID} - {Titel}"
                DokumentType = "pdf"
                
            #Ændre dokumenttitlen:
            
            Titel = f"{AktID:04} - {DokumentID} - {Titel}.{DokumentType}"

            # Call function
            dt_AktIndex,non_pdf_docs= process_documents(
                dt_AktIndex,
                AktID,
                Titel,
                Dokumentkategori,
                Dokumentdato,
                DokumentID,
                BilagTilDok,
                DokBilag,
                Omfattet,
                Aktstatus,
                Begrundelse,
                IsDocumentPDF,
            )
            
            dt_non_pdf_docs.extend(non_pdf_docs) 

    #Det er en nova sag
    else:
        #Det er en Nova sag
        orchestrator_connection.log_info("Det er en Nova sag")
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

            Dokumentdato = row['Dokumentdato']
            if isinstance(Dokumentdato, pd.Timestamp):
                Dokumentdato = Dokumentdato.strftime("%d-%m-%Y")
            elif isinstance(Dokumentdato, str):
                Dokumentdato = datetime.strptime(Dokumentdato, "%d-%m-%Y").strftime("%d-%m-%Y")
            else:
                raise ValueError(f"Unexpected data type: {type(Dokumentdato)}")
            IsDocumentPDF = True
            orchestrator_connection.log_info(f"AktID til debug: {AktID}")

            # Declare the necessary variables
            base_path = "Teams/tea-teamsite10506/Delte dokumenter/Aktindsigter/"

            # Sanitize the title
            Titel = sanitize_title(Titel)

            Titel = calculate_available_title_length(base_path, Overmappe, Undermappe, AktID, DokumentID, Titel)


            if (("ja" in Aktstatus.lower() or "delvis" in Aktstatus.lower()) 
                and DokumentID != "" 
                and "ja" in Omfattet.lower()):
                
                orchestrator_connection.log_info("Henter dokument information")
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
                        orchestrator_connection.log_info(response.status_code)
                    else:
                        orchestrator_connection.log_info(f"Failed to fetch Sagstitel from NOVA. Status Code: {response.status_code}")
                except Exception as e:
                    raise Exception("Failed to fetch Sagstitel (Nova):", str(e))

                DokumentType = response.json()["documents"][0]["fileExtension"]
                DocumentUuid = response.json()["documents"][0]["documentUuid"]
                orchestrator_connection.log_info(DokumentType)
                
                #Downloader file
                
                TransactionID = str(uuid.uuid4())
                url = f"{KMDNovaURL}/Document/GetFile?api-version=2.0-Case"
                file_path = os.path.join("C:\\Users", os.getlogin(), "Downloads", f"{AktID:04} - {DokumentID} - {Titel}.{DokumentType}")

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
                        
                        orchestrator_connection.log_info(f"File successfully saved at: {file_path}")
                    else:
                        orchestrator_connection.log_info(f"Failed to fetch file from NOVA. Status Code: {response.status_code}")
                        orchestrator_connection.log_info(f"Response: {response.text}")  # Print error message from API

                except Exception as e:
                    raise Exception("Failed to fetch file from NOVA:", str(e))
                
                
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
                    orchestrator_connection.log_info("Filen skal ikke konverteres")

                else:
                    conversionPossible = check_conversion_possible(DokumentType, CloudConvertAPI)
                    
                    if not conversionPossible:
                        orchestrator_connection.log_info(f"Skipping cause CloudConvert doesn't support: {DokumentType}->PDF")
                        ByteResult = bytes()                  
                        #Skal der sættes en bolean value?
                    else:
                        file_path = convert_file_to_pdf(CloudConvertAPI, file_path, DokumentID, DokumentType,Titel, AktID)
                        if file_path:
                            orchestrator_connection.log_info(f"PDF saved at: {file_path}")
                            DokumentType = "pdf"
                                                    

                if conversionPossible or CanDocumentBeConverted:
                    
                    success = upload_to_filarkiv(
                        FilarkivURL, FilarkivCaseID, Filarkiv_access_token,
                        AktID, DokumentID, Titel, file_path, DokumentType= DokumentType
                    )

                    if success:
                        DokumentType = "pdf"

                    else:
                        IsDocumentPDF = False
                        #file path her burde allerede have filtype
                        #file_path = f"{file_path}.{DokumentType}"
                        upload_file_to_sharepoint(
                            site_url=SharePointURL,
                            Overmappe=Overmappe,
                            Undermappe=Undermappe,
                            file_path=file_path,
                            RobotUserName=RobotUserName,
                            RobotPassword=RobotPassword,
                            tenant = tenant, 
                            client_id = client_id, 
                            thumbprint = thumbprint, 
                            cert_path = cert_path
                        )
                
                else: # Uploader til Sharepoint
                    orchestrator_connection.log_info("Could not be converted or uploaded - uploading directly to SharePoint")
                    IsDocumentPDF = False 
                    upload_file_to_sharepoint(
                            site_url=SharePointURL,
                            Overmappe=Overmappe,
                            Undermappe=Undermappe,
                            file_path=file_path,
                            RobotUserName=RobotUserName,
                            RobotPassword=RobotPassword,
                            tenant = tenant, 
                            client_id = client_id, 
                            thumbprint = thumbprint, 
                            cert_path = cert_path
                        )
    
            else:
                Titel = f"{AktID:04} - {DokumentID} - {Titel}"
                DokumentType = "pdf"    
            
            #Ændre dokumenttitlen:
            Titel = f"{AktID:04} - {DokumentID} - {Titel}.{DokumentType}"

            # Call function
            dt_AktIndex,non_pdf_docs= process_documents(
                dt_AktIndex,
                AktID,
                Titel,
                Dokumentkategori,
                Dokumentdato,
                DokumentID,
                BilagTilDok,
                DokBilag,
                Omfattet,
                Aktstatus,
                Begrundelse,
                IsDocumentPDF,
            )
            
            dt_non_pdf_docs.extend(non_pdf_docs)

    
    #Send Email
    if dt_non_pdf_docs:
        # Send Email Notification
        FinalString = "<br><br>".join(set(dt_non_pdf_docs))  # Remove duplicates

        # SharePoint integration
        credentials = UserCredential(RobotUserName, RobotPassword)
        ctx = ClientContext(SharePointURL).with_credentials(credentials)

        cert_credentials = {
            "tenant": tenant,
            "client_id": client_id,
            "thumbprint": thumbprint,
            "cert_path": cert_path
        }
        ctx = ClientContext(SharePointURL).with_client_certificate(**cert_credentials)
        folder_or_file_url = f"/Teams/tea-teamsite10506/Delte Dokumenter/Aktindsigter/{Overmappe}/{Undermappe}"
        target_item = ctx.web.get_folder_by_server_relative_url(folder_or_file_url)


        try:
            result = target_item.share_link(2).execute_query()  # Organization view link
            link_url = result.value.sharingLinkInfo.Url

            # Prepare email
            sender = "aktbob@aarhus.dk"
            subject = f"{Sagsnummer} - Filer kan ikke konverteres til PDF"
            body = (
                f'Sag: <a href="https://mtmsager.aarhuskommune.dk/app#/t/ticket/{DeskProID}">{DeskProID} - {DeskProTitel}</a><br><br>'
                "Kære Sagsbehandler,<br><br>"
                "Følgende dokumenter kunne ikke konverteres til PDF:<br><br>"
                f"{FinalString}<br><br>"
                "Dokumenterne er blevet uploaded til SharePoint-mappen: "
                f'<a href="{link_url}">SharePoint</a><br><br>'
                "<b>Bemærk:</b> Du kan ikke bruge FilArkiv til at screene, gennemgå eller redigere denne fil.<br><br>"
                "<li>Hvis filen kan udleveres som den er:</b> Gå videre med aktindsigten som normalt.</li>"
                "<li>Hvis det er en mediefil (lyd/video):</b> Brug redigeringssoftware til at fjerne dele, som modtageren ikke må se/høre. "
                "Har du ikke værktøjer eller viden, kan du kontakte Aktbob-teamet for hjælp.</li>"
                "<li>Hvis filen ikke skal udleveres:</b> Vælg 'Nej' i dokumentlisten, angiv en gyldig begrundelse, og slet filen fra SharePoint.</li>"
                "<br>"
                "Øvrige dokumenter overføres til FilArkiv og gennemgås der. Når du overfører fra FilArkiv til udleveringsmappen, opdateres aktlisten automatisk.<br>"
            )

            smtp_server = "smtp.adm.aarhuskommune.dk"
            smtp_port = 25


            send_email(
                receiver=MailModtager,
                sender=sender,
                subject=subject,
                body=body,
                smtp_server=smtp_server,
                smtp_port=smtp_port,
                html_body=True
            )

        except Exception as e:
            raise Exception(f"Error sending email: {e}")     
    
    dt_AktIndex = dt_AktIndex.drop('IsDocumentPDF', axis=1)
    



    return {
    "out_dt_AktIndex": dt_AktIndex,
    }




