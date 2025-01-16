def invoke_GenerateCasefolder(Arguments_GenerateCaseFolder): 
    
    from datetime import datetime
    import requests
    import os 
    import json
    from msal import PublicClientApplication

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
            print("Sagen findes i Filarkiv")
            
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
                    print("Kunne ikke omdøbe sagen:", str(e))

             
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
                print("Kunne ikke oprette sagen på ny:", str(e))


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

            try:
                # Send POST request to create a new case
                response = requests.post(url, headers=headers, data=json_payload)

                response_json = response.json()
                Out_FilarkivCaseID = response_json["id"]

            except Exception as e:
                print("Kunne ikke oprette sagen på ny:", str(e))

    except Exception as e:
            print("Error occurred while processing the request:", str(e))

    # ---- Opretter Aktindsigtsmapperne i SharePoint ---- #
    try:
        if SharePointUrl.startswith("https://"):
            SharePointUrl = SharePointUrl[8:]

        SharePointUrl = SharePointUrl.replace(".sharepoint.com", ".sharepoint.com:")

        # MSAL configuration for getting the access token
        scopes = ["https://graph.microsoft.com/.default"]
        app = PublicClientApplication(client_id=SharePointAppID, authority=f"https://login.microsoftonline.com/{SharePointTenant}")

        token_response = app.acquire_token_by_username_password(username=RobotUserName, password=RobotPassword, scopes=scopes)
        access_token = token_response.get("access_token")

        if not access_token:
            raise Exception("Failed to acquire access token. Check your credentials.")

        headers = {"Authorization": f"Bearer {access_token}"}

        # Step 1: Get the SharePoint site information
        site_request_url = f"https://graph.microsoft.com/v1.0/sites/{SharePointUrl}"
    
        site_response = requests.get(site_request_url, headers=headers)
        site_response.raise_for_status()
        site_json = site_response.json()

        if "id" not in site_json:
            raise Exception("Key 'id' not found in site response")
        site_id = site_json["id"]
 
        # Step 2: Get the document library (drive) information
        drive_request_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        drive_response = requests.get(drive_request_url, headers=headers)
        drive_response.raise_for_status()
        drive_json = drive_response.json()

        if "id" not in drive_json:
            raise Exception("Key 'id' not found in drive response")
        drive_id = drive_json["id"]
        print(f"Drive ID: {drive_id}")

        # Step 3: Create Folder1 (Mappe1)
        drive_item_request_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/Aktindsigter:/children"
        folder_request_body1 = {
            "name": Overmappe,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail"  # Or "replace", "fail"
        }
        
        folder_response1 = requests.post(drive_item_request_url, headers=headers, json=folder_request_body1)

        if folder_response1.status_code == 201:
            print(f"Folder '{Overmappe}' created successfully.")
        else:
            print(f"Error: Failed to create folder '{Overmappe}'. Status Code: {folder_response1.status_code}, Details: {folder_response1.text}")

        # Step 4: Create Folder2 (Mappe2 inside Mappe1)
        drive_item_request_url_for_mappe2 = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/Aktindsigter/{Overmappe}:/children"
        folder_request_body2 = {
            "name": Undermappe,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail"  # Or "replace", "fail"
        }

        folder_response2 = requests.post(drive_item_request_url_for_mappe2, headers=headers, json=folder_request_body2)

        if folder_response2.status_code == 201:
            print(f"Folder '{Undermappe}' created successfully inside '{Overmappe}'.")
        else:
            print(f"Error: Failed to create folder '{Undermappe}' inside '{Overmappe}'. Status Code: {folder_response2.status_code}, Details: {folder_response2.text}")

    except Exception as ex:
        print(f"Error: {ex}")
        raise

    finally:
        return {
        "out_FilarkivCaseID": Out_FilarkivCaseID,
        }
