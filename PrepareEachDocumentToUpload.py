from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
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
from GetFilarkivAcessToken import GetFilarkivToken
from email import policy
from email.parser import BytesParser
from pathlib import Path
import base64

def invoke_PrepareEachDocumentToUpload(Arguments_PrepareEachDocumentToUpload, orchestrator_connection: OrchestratorConnection):

    # henter in_argumenter:
    dt_DocumentList = Arguments_PrepareEachDocumentToUpload.get("in_dt_Documentlist")
    CloudConvertAPI = Arguments_PrepareEachDocumentToUpload.get("in_CloudConvertAPI")
    MailModtager = Arguments_PrepareEachDocumentToUpload.get("in_MailModtager")
    RobotUserName = Arguments_PrepareEachDocumentToUpload.get("in_RobotUserName")
    RobotPassword = Arguments_PrepareEachDocumentToUpload.get("in_RobotPassword")
    FilarkivCaseID = Arguments_PrepareEachDocumentToUpload.get("in_FilarkivCaseID")
    SharePointURL = Arguments_PrepareEachDocumentToUpload.get("in_SharePointUrl")
    Overmappe = Arguments_PrepareEachDocumentToUpload.get("in_Overmappe")
    Undermappe = Arguments_PrepareEachDocumentToUpload.get("in_Undermappe")
    Sagsnummer = Arguments_PrepareEachDocumentToUpload.get("in_Sagsnummer")
    GeoSag = Arguments_PrepareEachDocumentToUpload.get("in_GeoSag")
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

    # Create an empty DataFrame with the defined structure
    dt_AktIndex = pd.DataFrame(dt_AktIndex)
    dt_non_pdf_docs = []
    mimetypes.add_type("application/x-msmetafile", ".emz")
    
    timestamp = time.time()  
    
    document_number = 1

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
            elapsed = time.time() - timestamp
            if elapsed >= 31 * 60:  # 45 minutes in seconds
                print("30 minutes passed, fetching new filarkiv tokens and resetting timestamp.")
                Filarkiv_access_token = GetFilarkivToken(orchestrator_connection)
                timestamp = time.time()
            Omfattet = str(row["Omfattet af ansøgningen? (Ja/Nej)"])
            DokumentID = str(row["Dok ID"])
            
            # Handle AktID conversion
            AktID = row['Akt ID']
            if isinstance(AktID, str):  
                AktID = int(AktID.replace('.', ''))
            elif isinstance(AktID, int):  
                AktID = AktID

            Titel = str(row["Dokumenttitel"])
            
            # Split title into name and extension
            Titel, ext = os.path.splitext(Titel)

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
            
            orchestrator_connection.log_info(f"AktID til debug: {AktID}")

            # Declare the necessary variables
            base_path = "Teams/tea-teamsite10506/Delte dokumenter/Aktindsigter/"

            # Sanitize the title
            Titel = sanitize_title(Titel)

            Titel = calculate_available_title_length(base_path, Overmappe, Undermappe, AktID, DokumentID, Titel)

            if (("ja" in Aktstatus.lower() or "delvis" in Aktstatus.lower()) 
                and DokumentID != ""):

                Metadata = fetch_document_info_go(DokumentID, session, AktID, Titel)

                # Extracting variables for further use in the loop
                DokumentType = Metadata["DokumentType"]
                VersionUI = Metadata["VersionUI"]
                file_title = Metadata["file_title"]
                CanDocumentBeConverted = False
                conversionPossible = False
                file_path = file_title+"."+DokumentType
                
                # Tjekker om Goref-fil
                if "goref" in DokumentType:
                    orchestrator_connection.log_info("Dokumenter er .GORef")
                    ByteResult = fetch_document_bytes(session, DokumentID, file_path, delete_after_use=False, orchestrator_connection = orchestrator_connection)

                    if ByteResult:
                        with open(file_path, "r", encoding="utf-8") as file:
                            RefDokument = file.read()
                        
                        refdocument = RefDokument.split("?docid=")[1]
                        DokumentID = refdocument.split('"')[0]
                
                        os.remove(file_path)
                        orchestrator_connection.log_info("File deleted after use.")
                    
                    orchestrator_connection.log_info(f"GorefDokID: {DokumentID}")
                    #Henter dokument data
                    Metadata = fetch_document_info_go(DokumentID, session, AktID, Titel)
                
                    # Extracting variables for further use in the loop
                    DokumentType = Metadata["DokumentType"]
                    orchestrator_connection.log_info(f"Dokumenttype gotten from goref {DokumentType}")
                    VersionUI = Metadata["VersionUI"]
                    file_title = Metadata["file_title"]
                    file_path = file_title+"."+DokumentType

                    
                if DokumentType.lower() == "pdf": # Hvis PDF downloader den byte-filen
                    #Downloader fil fra GO    
                    orchestrator_connection.log_info("Allerede PDF - downloader")
                    ByteResult = fetch_document_bytes(session, DokumentID, max_retries=5, retry_interval=30)

                    if ByteResult:
                        orchestrator_connection.log_info(f"File size: {len(ByteResult)} bytes")
                    else:
                        orchestrator_connection.log_info("No file was downloaded.")
                    download_file(file_path, ByteResult, DokumentID, GoUsername, GoPassword, orchestrator_connection) 
                    
                                                                                            
                else: #Dokumentet er ikke en pdf - forsøger at konverterer               
                    ByteResult = GOPDFConvert(DokumentID, VersionUI, GoUsername, GoPassword)
                    
                    if len(ByteResult) == 0:
                        orchestrator_connection.log_info("Go-convervision mislykkedes forsøger med Filarkiv")

                        #Downloader fil fra GO    
                        ByteResult = fetch_document_bytes(session, DokumentID, file_path=file_path)
                        if ByteResult:
                            orchestrator_connection.log_info(f"File size: {len(ByteResult)} bytes")
                        else:
                            orchestrator_connection.log_info("No file was downloaded.")
                        
                        download_file(file_path, ByteResult, DokumentID, GoUsername, GoPassword, orchestrator_connection)  

                        if DokumentType.lower() in ["mht", "mhtml"]:
                            orchestrator_connection.log_info("CDW MHTML detected – converting to HTML")

                            try:
                                new_html_path = cdw_mhtml_to_html(file_path)

                                os.remove(file_path)
                                file_path = new_html_path
                                DokumentType = "html"
                                CanDocumentBeConverted = True

                                orchestrator_connection.log_info(
                                    f"MHTML converted to HTML: {file_path}"
                                )

                            except Exception as e:
                                orchestrator_connection.log_error(
                                    f"Failed MHTML→HTML conversion: {e}"
                                )
                                CanDocumentBeConverted = False


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
                            orchestrator_connection.log_info("Filen konverteres med Filarkiv")
                            CanDocumentBeConverted = True
                        else:
                            CanDocumentBeConverted = False
                            conversionPossible = check_conversion_possible(DokumentType, CloudConvertAPI)
                            
                            if not conversionPossible:
                                orchestrator_connection.log_info(f"Skipping cause CloudConvert doesn't support: {DokumentType}->PDF")
                                ByteResult = bytes()                  
                                #Skal der sættes en bolean value?
                            else:
                                orchestrator_connection.log_info("Forsøger med CloudConvert")
                                conversion = convert_file_to_pdf(CloudConvertAPI, file_path, DokumentID, DokumentType)
                                if conversion:
                                    file_path = conversion
                                    orchestrator_connection.log_info(f"PDF saved at: {file_path}")
                                    DokumentType = "pdf"

                    else: # Go-conversion lykkedes downloader fil
                        orchestrator_connection.log_info("Go-conversion lykkedes")                  
                        file_path = (f"{file_title}.pdf")
                        DokumentType = "pdf"
                        download_file(file_path, ByteResult, DokumentID, GoUsername, GoPassword, orchestrator_connection)

                if file_path.lower().endswith(".pdf") or CanDocumentBeConverted:
                    success, document_number = upload_to_filarkiv(
                    FilarkivURL, FilarkivCaseID, Filarkiv_access_token,
                    AktID, DokumentID, Titel, file_path,
                    DokumentType=DokumentType,
                    orchestrator_connection=orchestrator_connection,
                    document_number=document_number
                )


                    if success:
                        DokumentType = "pdf"
                        os.remove(file_path)
                        file_path = file_title+DokumentType
                        IsDocumentPDF = True
                        
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
                        os.remove(file_path)

                else:
                    # This is when conversion is not possible and upload shouldn't even be attempted to Filarkiv
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
                    os.remove(file_path)
            else:
                DokumentType = ".pdf"
                IsDocumentPDF = True
                
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
            
            elapsed = time.time() - timestamp
            if elapsed >= 31 * 60:  # 45 minutes in seconds
                print("30 minutes passed, fetching new filarkiv tokens and resetting timestamp.")
                Filarkiv_access_token = GetFilarkivToken(orchestrator_connection)
                timestamp = time.time()
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

                response = requests.put(url, headers=headers, json=payload)
                response.raise_for_status

                DokumentType = response.json()["documents"][0]["fileExtension"]
                DocumentUuid = response.json()["documents"][0]["documentUuid"]
                orchestrator_connection.log_info(DokumentType)
                
                #Downloader file
                
                TransactionID = str(uuid.uuid4())
                url = f"{KMDNovaURL}/Document/GetFile?api-version=2.0-Case"
                file_path = f"{AktID:04} - {DokumentID} - {Titel}.{DokumentType}"

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


                # Send request to API (Use GET if API expects it; otherwise, use POST)
                response = requests.put(url, headers=headers, json=payload)
                response.raise_for_status()

                with open(file_path, "wb") as file:
                    file.write(response.content)
                
                orchestrator_connection.log_info(f"File successfully saved at: {file_path}")
                
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
                        conversion = convert_file_to_pdf(CloudConvertAPI, file_path, DokumentID, DokumentType)
                        if conversion:
                            file_path = conversion
                            orchestrator_connection.log_info(f"PDF saved at: {file_path}")
                            DokumentType = "pdf"
                                                    

                if conversionPossible or CanDocumentBeConverted:
                    
                    success, document_number = upload_to_filarkiv(
                        FilarkivURL, FilarkivCaseID, Filarkiv_access_token,
                        AktID, DokumentID, Titel, file_path,
                        DokumentType=DokumentType,
                        orchestrator_connection=orchestrator_connection,
                        document_number=document_number
                    )


                    if success:
                        os.remove(file_path)
                        DokumentType = "pdf"
                        IsDocumentPDF = True

                    else:
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
                        os.remove(file_path)
                
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
                    os.remove(file_path)
    
            else:
                DokumentType = "pdf"   
                IsDocumentPDF = True 
        
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


    
    dt_AktIndex = dt_AktIndex.drop('IsDocumentPDF', axis=1)

    return {
    "out_dt_AktIndex": dt_AktIndex,
    }


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

def upload_to_filarkiv(FilarkivURL, FilarkivCaseID, Filarkiv_access_token, AktID, DokumentID, Titel, file_path, DokumentType, orchestrator_connection, document_number):
    Filarkiv_DocumentID = None  # Ensure it is initialized
    FileName = f"{AktID:04} - {DokumentID} - {Titel}"

    DocumentDate = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    data = {
        "caseId": FilarkivCaseID,
        "securityClassificationLevel": 0,
        "title": FileName,
        "documentNumber": document_number,
        "documentDate": DocumentDate,
        "direction": 0,
        "documentReference": DokumentID
    }
    response = requests.post(f"{FilarkivURL}/Documents", headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, data=json.dumps(data))
    if response.status_code in [200, 201]:
        Filarkiv_DocumentID = response.json().get("id")
        orchestrator_connection.log_info(f"Anvender følgende Filarkiv_DocumentID: {Filarkiv_DocumentID}")
    else:
        orchestrator_connection.log_info(f"Failed to create document. Response: {response.text}")

    if Filarkiv_DocumentID is None:
        orchestrator_connection.log_info("Fejl: Filarkiv_DocumentID blev ikke genereret. Afbryder processen.")
        return False, document_number+1
    
    extension = f".{DokumentType}"
    mime_type = {
        ".txt": "text/plain", ".pdf": "application/pdf", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
        ".gif": "image/gif", ".doc": "application/msword", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xls": "application/vnd.ms-excel", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".csv": "text/csv",
        ".json": "application/json", ".xml": "application/xml"
    }.get(extension, "application/octet-stream")
    FileName += extension
    orchestrator_connection.log_info(f"Anvender følgende dokumentID: {Filarkiv_DocumentID}")
    response = requests.post(f"{FilarkivURL}/Files", headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, json={"documentId": Filarkiv_DocumentID, "fileName": FileName, "sequenceNumber": 0, "mimeType": mime_type})
    if response.status_code in [200, 201]:
        FileID = response.json().get('id')
        orchestrator_connection.log_info(f"FileID: {FileID}")
    else:
        orchestrator_connection.log_info(f"Failed to create file metadata. {response.text}")
        return False, document_number+1
    
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
                orchestrator_connection.log_info(f"Failed to upload file. Status Code: {response.status_code} - deleting file + document")
                url = f"https://core.filarkiv.dk/api/v1/Files"
                data = {"id": FileID}
                response = requests.delete(url, headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, data=json.dumps(data))
                orchestrator_connection.log_info(f"File deletion status code: {response.status_code}")

                url = f"https://core.filarkiv.dk/api/v1/Documents"
                data = {"id": Filarkiv_DocumentID}
                response = requests.delete(url, headers={"Authorization": f"Bearer {Filarkiv_access_token}", "Content-Type": "application/json"}, data=json.dumps(data))
                orchestrator_connection.log_info(f"Document deletion status code: {response.status_code}")
                return False, document_number+1

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
    return True, document_number+1


def check_conversion_possible(dokument_type, cloudconvert_api):
    
    url = f"https://api.cloudconvert.com/v2/convert/formats?filter[input_format]={dokument_type}&filter[output_format]=pdf&filter[operation]=convert"
    
    headers = {
        "Authorization": cloudconvert_api
    }
    
    conversion_possible = False

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
    
    return conversion_possible

def convert_file_to_pdf(CloudConvertAPI, file_path, DokumentID, DokumentType):
    print("Conversion is supported!")
    
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
        print("Error: No tasks found in job creation response.")
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
    
    
    upload_data = {key: value for key, value in upload_parameters.items()}
    
    with open(file_path, "rb") as file:
        upload_files = {"file": file}
        upload_response = requests.post(upload_url, data=upload_data, files=upload_files)
    
    if upload_response.status_code == 201:
        print("File uploaded successfully!")
    else:
        print(f"Upload failed: {upload_response.status_code} - {upload_response.text}")
        return None
    
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
        print("Error: Invalid conversion type.")
        return None
    
    export_task_id = convert_export_response_data["data"]["tasks"][1]["id"]
    
    while True:
        status_check_url = f"https://api.cloudconvert.com/v2/tasks/{export_task_id}"
        status_check_response = requests.get(status_check_url, headers=create_job_headers)
        status_check_data = status_check_response.json()
        
        task_status = status_check_data["data"]["status"]
        
        if task_status == "finished":
            os.remove(file_path)
            download_url = status_check_data["data"]["result"]["files"][0]["url"]
            
            with requests.get(download_url, stream=True) as r:
                with open(file_path, "wb") as file:
                    for chunk in r.iter_content(chunk_size=8192):
                        file.write(chunk)
            
            print(f"File downloaded successfully at: {file_path}")
            
            return file_path
        elif task_status not in ["waiting", "processing"]:
            print(f"An error occurred:{status_check_response.text}")
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
                    print(f"Deleted directory: {file_path}")
        except Exception as e:
            raise Exception(f"Error deleting {file_path}: {e}")

    # Return dt_AktIndex and ListOfNonPDFDocs, both WITH extensions
    return dt_AktIndex, ListOfNonPDFDocs

def fetch_document_info_go(DokumentID, session, AktID, Titel):
    url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/Data/{DokumentID}"
    response = session.get(url)
    DocumentData = response.text
    data = json.loads(DocumentData)
    item_properties = data.get("ItemProperties", "")
    file_type_match = re.search(r'ows_File_x0020_Type="([^"]+)"', item_properties)
    version_ui_match = re.search(r'ows__UIVersionString="([^"]+)"', item_properties)
    DokumentType = file_type_match.group(1) if file_type_match else "unknown"
    VersionUI = version_ui_match.group(1) if version_ui_match else "Not found"
    Feedback = " "
    file_title = f"{AktID:04} - {DokumentID} - {Titel}"
    return {"DokumentType": DokumentType, "VersionUI": VersionUI, "Feedback": Feedback, "file_title": file_title}

def download_file(file_path, ByteResult, DokumentID, GoUsername, GoPassword, orchestrator_connection):
    try:
        with open(file_path, "wb") as file:
            file.write(ByteResult)
        orchestrator_connection.log_info("File written successfully.")
        return
    except Exception as initial_exception:
        orchestrator_connection.log_info(f"Failed, trying from URL: {DokumentID} Path: {file_path}")
        orchestrator_connection.log_info(initial_exception)

        ByteResult = bytes()

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
            return
        except Exception as retry_exception:
            orchestrator_connection.log_info(f"Retry {attempt + 1} failed: {retry_exception}")
            if attempt == max_retries - 1:
                raise RuntimeError(
                    f"Failed to download file after {max_retries} retries. "
                    f"DokumentID: {DokumentID}, Path: {file_path}"
                )
            time.sleep(5)


def fetch_document_bytes(session: requests.Session, DokumentID, file_path=None, max_retries=30, retry_interval=5, delete_after_use=False):

    url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/DocumentBytes/{DokumentID}"
    ByteResult = None
    
    response = None
    for attempt in range(max_retries):
        try:
            response = session.get(url, timeout=180)
            response.raise_for_status()
            if response.status_code == 200:
                ByteResult = response.content
                if b"HTTP Error 503. The service is unavailable." in ByteResult:
                    ByteResult = None
                    print(f"Attempt {attempt + 1}: Failed due to HTTP Error 503. The service is unavailable")
                    continue
                else:
                    break
            else:
                print(f"Attempt {attempt + 1}: Failed with status code {response.status_code}")
        except Exception as e:
            print(f"Attempt {attempt + 1}: Exception occurred - {e}")

        time.sleep(retry_interval)
        
    # If a file path is given, save to file
    if file_path and ByteResult:
        with open(file_path, "wb") as file:
            file.write(ByteResult)

        # If delete_after_use is True, remove the file
        if delete_after_use:
            os.remove(file_path)

    return ByteResult

def GOPDFConvert (DokumentID, VersionUI, GoUsername, GoPassword):
    try:
        url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/ConvertToPDF/{DokumentID}/{VersionUI}"
        response = requests.get(
            url,
            auth=HttpNtlmAuth(GoUsername, GoPassword),
            headers={"Content-Type": "application/json"},
            timeout=None  # Equivalent to client.Timeout = -1
        )
        
        response.raise_for_status
        
        # Feedback and byte result
        Feedback = response.text
        if "Document could not be converted" in Feedback:
            return ""
        else:
            return response.content
    
    except Exception as e:
        return ""

def _decode_html_part(part):
    """
    Decode HTML payload using declared charset with safe fallbacks.
    Critical for Danish characters (æ ø å).
    """
    payload = part.get_payload(decode=True)

    # Charset declared in the MIME part
    charset = part.get_content_charset()

    # Outlook / CDW fallback order
    for enc in [charset, "windows-1252", "iso-8859-1", "utf-8"]:
        if not enc:
            continue
        try:
            return payload.decode(enc)
        except UnicodeDecodeError:
            continue

    # Last-resort fallback (never crashes)
    return payload.decode("utf-8", errors="replace")
def cdw_mhtml_to_html(mhtml_path):
    """
    Converts CDW/Outlook MHTML mail to a single self-contained HTML file.
    Preserves Danish characters correctly.
    Returns path to new HTML file.
    """
    with open(mhtml_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    html_body = None
    attachments = []

    for part in msg.walk():
        ctype = part.get_content_type()
        disp = part.get_content_disposition()

        if ctype == "text/html" and html_body is None:
            html_body = _decode_html_part(part)

        elif disp == "attachment":
            attachments.append({
                "filename": part.get_filename(),
                "ctype": ctype,
                "data": part.get_payload(decode=True),
            })

    def attachment_html(att):
        fn = att["filename"]
        ct = att["ctype"]
        data = base64.b64encode(att["data"]).decode("ascii")

        if ct and ct.startswith("image/"):
            return f"""
            <div class="attachment">
                <h4>{fn}</h4>
                <img src="data:{ct};base64,{data}" style="max-width:100%;">
            </div>
            """

        if ct == "application/pdf":
            return f"""
            <div class="attachment">
                <h4>{fn}</h4>
                <iframe src="data:application/pdf;base64,{data}"
                        width="100%" height="800"></iframe>
            </div>
            """

        return f"""
        <div class="attachment">
            <h4>{fn}</h4>
            <a download="{fn}" href="data:{ct};base64,{data}">
                Download attachment
            </a>
        </div>
        """

    attachments_html = "\n".join(attachment_html(a) for a in attachments)

    final_html = f"""<!DOCTYPE html>
<html lang="da">
<head>
<meta charset="utf-8">
<title>Mailarkiv</title>
<style>
body {{
    font-family: Arial, Helvetica, sans-serif;
    font-size: 11pt;
}}
.attachments {{
    page-break-before: always;
}}
.attachment {{
    margin-bottom: 30px;
}}
</style>
</head>
<body>

{html_body or ""}

<div class="attachments">
<h2>Bilag</h2>
{attachments_html}
</div>

</body>
</html>
"""

    html_path = Path(mhtml_path).with_suffix(".html")
    html_path.write_text(final_html, encoding="utf-8")

    return str(html_path)