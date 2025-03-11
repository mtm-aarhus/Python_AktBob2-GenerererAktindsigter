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
                
                print("Henter dokument information")
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
                    raise Exception("Failed to fetch Sagstitel (Nova):", str(e))

                DokumentType = response.json()["documents"][0]["fileExtension"]
                DocumentUuid = response.json()["documents"][0]["documentUuid"]
                print(DokumentType)
                
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
                        print("Failed to fetch file from NOVA. Status Code:", response.status_code)
                        print("Response:", response.text)  # Print error message from API

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
                    print("Filen skal ikke konverteres")

                else:
                    conversionPossible = check_conversion_possible(DokumentType, CloudConvertAPI)
                    
                    if not conversionPossible:
                        print(f"Skipping cause CloudConvert doesn't support: {DokumentType}->PDF")
                        ByteResult = bytes()                  
                        #Skal der sættes en bolean value?
                    else:
                        file_path = convert_file_to_pdf(CloudConvertAPI, file_path, DokumentID, DokumentType,Titel, AktID)
                        if file_path:
                            print(f"PDF saved at: {file_path}")
                            DokumentType = "pdf"
                                                    

                if conversionPossible or CanDocumentBeConverted:
                    
                    upload_to_filarkiv(FilarkivURL,FilarkivCaseID, Filarkiv_access_token, AktID, DokumentID,Titel, file_path)
                    if conversionPossible:
                        DokumentType = "pdf"

                else: # Uploader til Sharepoint
                    orchestrator_connection.log_info("Could not be converted or uploaded - uploading directly to SharePoint")
                    IsDocumentPDF = False 
                    upload_file_to_sharepoint(
                            site_url=SharePointURL,
                            Overmappe=Overmappe,
                            Undermappe=Undermappe,
                            file_path=file_path,
                            RobotUserName=RobotUserName,
                            RobotPassword=RobotPassword
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































            Begrundelse = str(row["Begrundelse hvis nej eller delvis"])
            Dokumentdato =row['Dokumentdato']
            print(str(Dokumentdato))
            if isinstance(Dokumentdato, pd.Timestamp):
                Dokumentdato = Dokumentdato.strftime("%d-%m-%Y")
                print(f"følgende dokument: {Titel} - har følgende dato:({type(Dokumentdato)})")
            else:
                print(f"følgende dokument: {Titel} - har følgende dato:({type(Dokumentdato)})")
                print(Dokumentdato)
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
                
                orchestrator_connection.log_info("Dokumentet er omfattet i ansøgningen")
                Metadata = fetch_document_info(DokumentID, session, AktID, Titel)
                
                # Extracting variables for further use in the loop
                DokumentType = Metadata["DokumentType"]
                VersionUI = Metadata["VersionUI"]
                Feedback = Metadata["Feedback"]
                file_path = Metadata["file_path"]


                # Tjekker om Goref-fil
                if ".goref" in file_path:
                    
                    ByteResult = fetch_document_bytes(session, DokumentID, file_path, delete_after_use=True)

                    if ByteResult:
                        with open("temp_document", "r", encoding="utf-8") as file:
                            RefDokument = file.read()
                        
                        refdocument = RefDokument.split("?docid=")[1]
                        DokumentID = refdocument.split('"')[0]
                    
                    #Henter dokument data
                    Metadata = fetch_document_info(DokumentID, session, AktID, Titel)
                
                    # Extracting variables for further use in the loop
                    DokumentType = Metadata["DokumentType"]
                    VersionUI = Metadata["VersionUI"]
                    Feedback = Metadata["Feedback"]
                    file_path = Metadata["file_path"]

                if DokumentType.lower() == "pdf": # Hvis PDF downloader den byte-filen
                    print("Allerede PDf - downloader byte")
                    
                    ByteResult = fetch_document_bytes(session, DokumentID, max_retries=5, retry_interval=30)

                    if ByteResult:
                        print(f"File size: {len(ByteResult)} bytes")
                    else:
                        print("No file was downloaded.")

                    
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
                        raise Exception(f"An exception occurred: {e}")
                    
                    
                    # tjekker om go-conversion lykkedes eller ej
                    if "Document could not be converted" in Feedback or len(ByteResult) == 0:

                        conversionPossible= check_conversion_possible(DokumentType, CloudConvertAPI)
                        
                        if not conversionPossible:
                            print(f"Skipping cause CloudConvert doesn't support: {DokumentType}->PDF")
                            ByteResult = bytes()                  
                        else:

                            FilnavnFørPdf = f"Output.{DokumentType}"
                            ByteResult = fetch_document_bytes(session, DokumentID, file_path=FilnavnFørPdf)

                            if ByteResult:
                                print(f"File saved as: {FilnavnFørPdf}")

                            file_path = convert_file_to_pdf(CloudConvertAPI, FilnavnFørPdf, DokumentID, DokumentType,Titel, AktID)
                            if file_path:
                                print(f"PDF saved at: {file_path}")
                                DokumentType = "pdf"


                if "Document could not be converted" in Feedback or len(ByteResult) == 0:
                    print(f"Could not be converted, uploading as {file_path}.{DokumentType}")

                    file_path = f"{file_path}.{DokumentType}"
                    ByteResult = fetch_document_bytes(session, DokumentID, file_path, max_retries=5, retry_interval=60)

                    if ByteResult:
                        print("File downloaded successfully.")
                    else:
                        print("ByteResult is empty.")

                else: 
                    file_path = (f"{file_path}.pdf")   

                
                #Downloader filen og konventerer fra byte til fil
                download_file(file_path, ByteResult, DokumentID, GoUsername, GoPassword)
                

                if ".pdf" in file_path: # SKAL MÅSKE ÆNDRES, SÅ DEN TAGER HØJDE FOR ANDET END PDF. 
                    upload_to_filarkiv(FilarkivURL,FilarkivCaseID, Filarkiv_access_token, AktID, DokumentID,Titel, file_path)
                    DokumentType = "pdf"
                
                else: # Filtypen er ikke understøttet, uploader til Sharepoint
                    orchestrator_connection.log_info("Could not be converted or uploaded - uploading directly to SharePoint")
                    IsDocumentPDF = False 
                    upload_file_to_sharepoint(
                        site_url=SharePointURL,
                        Overmappe=Overmappe,
                        Undermappe=Undermappe,
                        file_path=file_path,
                        RobotUserName=RobotUserName,
                        RobotPassword=RobotPassword
                    )
                    

            else:
                orchestrator_connection.log_info("Dokumentet skal ikke med i ansøgningen")
                Titel = f"{AktID:04} - {DokumentID} - {Titel}"
                DokumentType = "pdf"
                
            #Ændre dokumenttitlen:
            Titel = f"{AktID:04} - {DokumentID} - {Titel}.{DokumentType}"