from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

def invoke_GenerateCasefolder(Arguments_GenerateCaseFolder, orchestrator_connection: OrchestratorConnection): 
    
    from datetime import datetime
    import requests
    import os 
    import json
    from msal import PublicClientApplication
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.webs.web import Web

    RobotUserName = Arguments_GenerateCaseFolder.get("in_RobotUserName")
    RobotPassword = Arguments_GenerateCaseFolder.get("in_RobotPassword")
    Sagsnummer = Arguments_GenerateCaseFolder.get("in_Sagsnummer")
    SharePointAppID = Arguments_GenerateCaseFolder.get("in_SharePointAppID")
    SharePointTenant = Arguments_GenerateCaseFolder.get("in_SharePointTenant")
    SharePointUrl = Arguments_GenerateCaseFolder.get("in_SharePointUrl")
    Overmappe = Arguments_GenerateCaseFolder.get("in_Overmappe")
    Undermappe = Arguments_GenerateCaseFolder.get("in_Undermappe")
    Sagstitel = Arguments_GenerateCaseFolder.get("in_Sagstitel")
    Filarkiv_access_token = Arguments_GenerateCaseFolder.get("in_Filarkiv_access_token") 
    DeskProTitel = Arguments_GenerateCaseFolder.get("in_DeskProTitel")
    DeskProID = Arguments_GenerateCaseFolder.get("in_DeskProID")
    FilarkivURL = Arguments_GenerateCaseFolder.get("in_FilarkivURL")
    

    # ---- Opretter sagen i Filarkiv ---- #
    CaseDate = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    CaseTypeId = "02d43e06-edb9-4497-15cc-08d9dbf0c7dc"
    CaseStatusId = "b7c08e6d-b2f1-4830-8e33-08d9dbf0c9e5"
    archiveId = "b54fe478-2fa9-4a2b-c220-08dbbdac5df5"



# ---- Tjekker om sagen findes i forvejen ----# 

    url = f"{FilarkivURL}/cases?caseNumber={DeskProID} / {Sagsnummer}"

    headers = {
        "Authorization": f"Bearer {Filarkiv_access_token}",
        "Content-Type": "application/xml"
    }

    try:
        # Send GET request
        response = requests.get(url, headers=headers)
        print("FilArkiv respons:", response.status_code)

        Respons = response.json()  # Parse the response as JSON
        
        # Check if the response contains data
        if Respons and isinstance(Respons, list) and len(Respons) > 0:
            orchestrator_connection.log_info("Sagen findes i Filarkiv")
            
            for current_item in Respons:
                # Check if "title" exists and contains the specific string
                title = current_item.get("title", "")
                
                if "UDDATERET/SLETTET" in title:
                    continue  # Skip this item and move to the next
                    

                else:
                    #Sagen er oprettet i forvejen - omdøber
                    Old_FilarkivCaseID = current_item.get("id","")
                    old_Title = (f"{title} - UDDATERET/SLETTET")

                # Define the variables
                NewJson = {
                    "id": Old_FilarkivCaseID,
                    "caseNumber": f"{DeskProID} / {Sagsnummer}",
                    "title": old_Title,
                    "alternateTitle": Sagstitel,
                    "caseDate": CaseDate,
                    "caseTypeId": CaseTypeId,
                    "caseStatusId": CaseStatusId,
                    "archiveId": archiveId,
                    "caseReference": DeskProID,
                    "securityClassificationLevel": 0  
                }

                Newjson_payload = json.dumps(NewJson)
                url = f"{FilarkivURL}/Cases"

                headers = {
                "Authorization": f"Bearer {Filarkiv_access_token}",
                "Content-Type": "application/Json"
                }

                try:
                    # Send POST request to create a new case
                    response = requests.put(url, headers=headers, data=Newjson_payload)
                    print("FilArkiv respons:", response.status_code)

                except Exception as e:
                    raise Exception("Kunne ikke omdøbe sagen:", str(e))

             
            # Prepare the payload for creating a new case
            payload = {
                "caseNumber": f"{DeskProID} / {Sagsnummer}",
                "title": f"[{DeskProID}]: {DeskProTitel} ({Sagsnummer})",
                "alternateTitle": Sagstitel,
                "caseDate": CaseDate,
                "caseTypeId": CaseTypeId,
                "caseStatusId": CaseStatusId,
                "archiveId": archiveId,
                "caseReference": DeskProID,
                "securityClassificationLevel": 0
            }

            # Convert dictionary to JSON string
            json_payload = json.dumps(payload)
            
            url = f"{FilarkivURL}/Cases"

            headers = {
                "Authorization": f"Bearer {Filarkiv_access_token}",
                "Content-Type": "application/Json"
            }

            try:
                # Send POST request to create a new case
                response = requests.post(url, headers=headers, data=json_payload)
                response_json = response.json()
                Out_FilarkivCaseID = response_json["id"]

            except Exception as e:
                raise Exception("Kunne ikke oprette sagen på ny.......", str(e))


        else:
            
            # Prepare the payload for creating a new case
            payload = {
                "caseNumber": f"{DeskProID} / {Sagsnummer}",
                "title": f"[{DeskProID}]: {DeskProTitel} ({Sagsnummer})",
                "alternateTitle": Sagstitel,
                "caseDate": CaseDate,
                "caseTypeId": CaseTypeId,
                "caseStatusId": CaseStatusId,
                "archiveId": archiveId,
                "caseReference": DeskProID,
                "securityClassificationLevel": 0
            }

            # Convert dictionary to JSON string
            json_payload = json.dumps(payload)
            
            url = f"{FilarkivURL}/Cases"

            headers = {
                "Authorization": f"Bearer {Filarkiv_access_token}",
                "Content-Type": "application/Json"
            }
            print(json_payload)
            print(url)
            print(headers)
            try:
                # Send POST request to create a new case
                response = requests.post(url, headers=headers, data=json_payload)

                response_json = response.json()
                Out_FilarkivCaseID = response_json["id"]

            except Exception as e:
                raise Exception("Kunne ikke oprette sagen på ny:", str(e))

    except Exception as e:
            raise Exception("Error occurred while processing the request:", str(e))

    # ---- Opretter Aktindsigtsmapperne i SharePoint ---- #


    # Authenticate
    credentials = UserCredential(RobotUserName, RobotPassword)
    ctx = ClientContext(SharePointUrl).with_credentials(credentials)
    

    def folder_exists(ctx, folder_url):
        """Check if a folder exists in SharePoint."""
        try:
            folder = ctx.web.get_folder_by_server_relative_url(folder_url)
            ctx.load(folder)
            ctx.execute_query()
            return True
        except:
            return False

    try:
        # Ensure that the SharePoint site is accessible
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print(f"Connected to SharePoint site: {web.properties['Title']}")

        # Define the document library name (update if different)
        library_name = "/Teams/tea-teamsite10506/Delte Dokumenter/Aktindsigter"  # Update if your document library has a different name



        # Construct folder paths (relative to the SharePoint site)
        overmappe_path = f"{library_name}/{Overmappe}"
        undermappe_path = f"{library_name}/{Overmappe}/{Undermappe}"

        # Step 1: Check and create Folder1 (Overmappe)
        if not folder_exists(ctx, overmappe_path):
            ctx.web.folders.add(overmappe_path)
            ctx.execute_query()
            print(f"Folder '{Overmappe}' created successfully.")
        else:
            print(f"Folder '{Overmappe}' already exists.")

        # Step 2: Check and create Folder2 (Undermappe inside Overmappe)
        if not folder_exists(ctx, undermappe_path):
            ctx.web.folders.add(undermappe_path)
            ctx.execute_query()
            print(f"Folder '{Undermappe}' created successfully inside '{Overmappe}'.")
        else:
            print(f"Folder '{Undermappe}' already exists inside '{Overmappe}'.")

    except Exception as ex:
        raise Exception(f"Error: {ex}")

    finally:
        return {
        "out_FilarkivCaseID": Out_FilarkivCaseID,
        }
