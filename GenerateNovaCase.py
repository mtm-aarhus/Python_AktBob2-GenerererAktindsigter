from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
def invoke_GenerateNovaCase(Arguments_GenerateNovaCase,orchestrator_connection: OrchestratorConnection):
    import uuid
    import requests
    import json
    import os
    from GetKmdAcessToken import GetKMDToken
    from datetime import datetime
    import base64
    from docx import Document
    import io
    import re
    
     # henter in_argumenter:
    Sagsnummer = Arguments_GenerateNovaCase.get("in_Sagsnummer")
    KMDNovaURL = Arguments_GenerateNovaCase.get("in_KMDNovaURL")
    KMD_access_token = Arguments_GenerateNovaCase.get("in_NovaToken")
    AktSagsURL = Arguments_GenerateNovaCase.get("in_AktSagsURL")
    IndsenderNavn = Arguments_GenerateNovaCase.get("in_IndsenderNavn")
    IndsenderMail = Arguments_GenerateNovaCase.get("in_IndsenderMail")  
    AktindsigtsDato = Arguments_GenerateNovaCase.get("in_AktindsigtsDato")
    DeskProID = Arguments_GenerateNovaCase.get("in_DeskProID")
    DeskProAPI = orchestrator_connection.get_credential("DeskProAPI")
    DeskProAPIKey = DeskProAPI.password
    AktindsigtsDato = AktindsigtsDato.rstrip('Z')

     ### --- Henter caseinfo --- ###
    TransactionID = str(uuid.uuid4())

    # Define API URL
    Caseurl = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"

    # Define headers
    headers = {
        "Authorization": f"Bearer {KMD_access_token}",
        "Content-Type": "application/json"
    }

    data = {
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
            "state": {
                "progressState": True,
                "activeCode": True
            },
            "sensitivity": {
                "sensitivityCtrBy": True
            },
            "securityUnit": {
                "departmentCtrlBy": True
            },
            "responsibleDepartment": {
                "fkOrgIdentity": {
                    "fkUuid": True,
                    "type": True,
                    "fullName": True
                },
                "departmentCtrlBy": True
            },  
            "availability": {
                "availabilityCtrBy": True
            },   
            "caseAttributes": {
                "title": True,
                "userFriendlyCaseNumber": True,
                "caseDate": True
            },
            "caseParty": {
                "index": True,
                "identificationType": True,
                "identification": True,
                "partyRole": True,
                "partyRoleName": True,
                "name": True,
                "ParticipantRole": True
            },
            "caseworker": {
                "kspIdentity": {
                    "novaUserId": True,
                    "racfId": True,
                    "fullName": True
                }
            },
            "buildingCase": {
            "propertyInformation":{
                "caseAddress":True,
                "esrPropertyNumber": True,
                "bfeNumber": True,
                "cadastralId": True
         }
        }
    }
    }
    try:
        response = requests.put(Caseurl, headers=headers, json=data)
        
        # Handle response
        if response.status_code == 200:
            response_data = response.json()
            case = response_data["cases"][0]  # Assuming there's at least one case

            # Extract required case attributes
            title = case["caseAttributes"]["title"]
            caseDate = case["caseAttributes"]["caseDate"]
            progressState = case["state"]["progressState"]
            sensitivityCtrBy = case["sensitivity"]["sensitivityCtrBy"]
            SecurityUnitCtrlBy = case["securityUnit"]["departmentCtrlBy"]
            ResponsibleDepartmentCtrlBy = case["responsibleDepartment"]["departmentCtrlBy"]
            availabilityCtrBy = case["availability"]["availabilityCtrBy"]
            
            # Extract bfeNumber from buildingCase -> propertyInformation
            bfeNumber = case["buildingCase"]["propertyInformation"]["bfeNumber"]
            CadastralId = case["buildingCase"]["propertyInformation"]["cadastralId"]
            print(bfeNumber)
            print(CadastralId)

            primary_case_parties = [
                {
                    "index": party["index"],
                    "identificationType": party["identificationType"],
                    "identification": party["identification"],
                    "partyRole": party["partyRole"],
                    "partyRoleName": party["partyRoleName"],
                    "participantRole": party["participantRole"],
                    "name": party["name"]
                }
                for party in case.get("caseParties", []) if party["partyRole"] == "PRI"
            ]

            # If at least one primary case party exists, assign values to variables
            if primary_case_parties:
                index = primary_case_parties[0]["index"]
                identificationType = primary_case_parties[0]["identificationType"]
                identification = primary_case_parties[0]["identification"]
                partyRole = primary_case_parties[0]["partyRole"]
                partyRoleName = primary_case_parties[0]["partyRoleName"]
                participantRole = primary_case_parties[0]["participantRole"]
                name = primary_case_parties[0]["name"]

                # Print to verify
                print("Index:", index)
                print("Identification Type:", identificationType)
                print("Identification:", identification)
                print("Party Role:", partyRole)
                print("Party Role Name:", partyRoleName)
                print("Participant Role:", participantRole)
                print("Name:", name)
            else:
                print("No primary case parties found.")

            # Print extracted case attributes
            print("title:", title)
            print("Case Date:", caseDate)
            print("Progress State:", progressState)
            print("Sensitivity Controlled By:", sensitivityCtrBy)
            print("Security Unit Controlled By:", SecurityUnitCtrlBy)
            print("Responsible Department Controlled By:", ResponsibleDepartmentCtrlBy)
            print("Availability Controlled By:", availabilityCtrBy)
        

        else:
            print("Failed to send request. Status Code:", response.status_code)
            print("Response Data:", response.text)  # Print error response
    except Exception as e:
        raise Exception("Failed to fetch case data:", str(e))



    ### ---- Henter deskpro info: --- ####

    Deskprourl = f"https://mtmsager.aarhuskommune.dk/api/v2/tickets/{DeskProID}"

    headers = {
    'Authorization': DeskProAPIKey,
    'Cookie': 'dp_last_lang=da'
    }
    # Regex pattern for old case numbers
    case_number_pattern = re.compile(r"^[A-Za-z]\d{4}-\d{1,10}$")
    old_case_numbers = []
    BFEMatch = False

    try:
        response = requests.get(Deskprourl, headers=headers)
        
        if response.status_code != 200:
            raise Exception(f"Request failed with status code {response.status_code}: {response.text}")
        
        data = response.json()
        fields = data.get("data", {}).get("fields", {})

        for field_data in fields.values():
            value = field_data.get("value")

            # Check if value is a string
            if isinstance(value, str):
                if case_number_pattern.match(value):
                    old_case_numbers.append(value)

            # Check if value is a list (could contain strings or numbers)
            elif isinstance(value, list):
                for item in value:
                    if isinstance(item, str) and case_number_pattern.match(item):
                        old_case_numbers.append(item)

        # Now loop through the found case numbers
        for case_number in old_case_numbers:
            print("Found old case number:", case_number)
             ### --- Henter caseinfo --- ###
            TransactionID = str(uuid.uuid4())

            # Define API URL
            Caseurl = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"

            # Define headers
            headers = {
                "Authorization": f"Bearer {KMD_access_token}",
                "Content-Type": "application/json"
            }

            data = {
                "common": {
                    "transactionId": TransactionID
                },
                "paging": {
                    "startRow": 1,
                    "numberOfRows": 100
                },
                "caseAttributes": {
                    "userFriendlyCaseNumber": case_number
                },
                "caseGetOutput": { 
                    "buildingCase": {
                        "propertyInformation":{
                            "bfeNumber": True
                            }
                        }
                    }
                }
            try:
                response = requests.put(Caseurl, headers=headers, json=data)

                if response.status_code == 200:
                    response_data = response.json()
                    if response_data.get("cases"):
                        case = response_data["cases"][0]
                        OldbfeNumber = case["buildingCase"]["propertyInformation"]["bfeNumber"]

                        if str(OldbfeNumber) == str(bfeNumber):
                            print(f"Match found: Old BFE ({OldbfeNumber}) == Current BFE ({bfeNumber})")
                            old_case_number = case_number
                            BFEMatch = True
                            break  # Exit loop after first match
                        else:
                            print(f"No match: Old BFE ({OldbfeNumber}) != Current BFE ({bfeNumber})")
                    else:
                        print(f"No cases found for {case_number}")
                else:
                    print(f"KMD API call failed for {case_number}, status: {response.status_code}, message: {response.text}")

            except Exception as e:
                print(f"An error occurred while calling KMD API for {case_number}: {e}")

        if not BFEMatch:
            print("No matching BFE number found in any case.")
        else:
            print("BFE match confirmed!")
             ## finder først den existerende sag i Nova ved at søge på titlen og checke at datoen er aktindsigtsdatoen:"
            TransactionID = str(uuid.uuid4())
            # Parse the string into a datetime object

            # Add one day
            new_date_obj = AktindsigtsDato + timedelta(days=1)

            print("Original:", AktindsigtsDato)
            print("Plus one day:", new_date_obj)
            

            # Define API URL
            Caseurl = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"

            # Define headers
            headers = {
                "Authorization": f"Bearer {KMD_access_token}",
                "Content-Type": "application/json"
            }


            data = {
            "common": {
                "transactionId": TransactionID
            },
            "paging": {
                "startRow": 1,
                "numberOfRows": 100
            },
            "caseAttributes": {
                "title": f"Test gustav - Anmodning om aktindsigt i {old_case_number}",
                "fromCaseDate": AktindsigtsDato,
                "toCaseDate": new_date_obj

            },
            "states":{
                "states":[{
                    "progressState":"Afgjort"
            }]
            },
            "caseGetOutput": { 
                "caseAttributes":{
                "userFriendlyCaseNumber": True # giver det sagsnummer som skal opdateres
                }
            }
            }
            # Make the request
            response = requests.post(Caseurl, headers=headers, json=data)

            # Check status and handle response
            if response.status_code == 200:
                case = response_data["cases"][0]
                OldAktindsigtscase = case["caseAttributes"]["userFriendlyCaseNumber"]
                print(OldAktindsigtscase)
            else:
                raise Exception(f"API request failed with status {response.status_code}: {response.text}")
                        


    except Exception as e:
        print(f"An error occurred during ticket processing: {e}")


    if BFEMatch == True:
        print("BFE matcher opdaterer sagen")
       
    else:
        print("No matching BFE number found opretter sagen på ny.")
        # ### ---  Opretter sagen --- ####   
        JournalDate = datetime.now().strftime("%Y-%m-%dT00:00:00")
        TransactionID = str(uuid.uuid4())
        Uuid = str(uuid.uuid4())
        JournalUuid = str(uuid.uuid4())
        Index_Uuid = str(uuid.uuid4())
        link_text = "GO Aktindsigtssag"
        BuildingClassUuid = str(uuid.uuid4())
        # Step 1: Create a new Word document
        doc = Document()
        doc.add_paragraph("Aktindsigtssag Link: " + AktSagsURL)  # Add content to the document


        # Step 2: Save document to a BytesIO stream
        doc_stream = io.BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)  # Reset stream position

        # Step 3: Convert document to base64
        base64_JournalNote = base64.b64encode(doc_stream.read()).decode("utf-8")

        url = f"{KMDNovaURL}/Case/Import?api-version=2.0-Case"

        # Define headers
        headers = {
            "Authorization": f"Bearer {KMD_access_token}",
            "Content-Type": "application/json"
        }

        # Define JSON payload
        payload = {
            "common": {
                "transactionId": TransactionID,
                "uuid": Uuid  
            },
            "caseAttributes": {
                "title": f"Test gustav - Anmodning om aktindsigt i {Sagsnummer}",
                "caseDate": AktindsigtsDate,
                "caseCategory": "BomByg"
            },
            "caseClassification": {
                "kleNumber": {"code": "02.00.00"}, 
                "proceedingFacet": {"code": "A53"}
            },
            "state": "Afgjort", 
            "sensitivity": "Følsomme",
            "caseworker": { 
                "kspIdentity": {
                    "novaUserId": "78897bfc-2a36-496d-bc76-07e7a6b0850e",
                    "racfId": "AZX0075",
                    "fullName": "Aktindsigter Novabyg"
                }
            },
            "SensitivityCtrlBy": sensitivityCtrBy,
            "AvailabilityCtrlBy": availabilityCtrBy,
            "SecurityUnitCtrlBy": SecurityUnitCtrlBy,
            "ResponsibleDepartmentCtrlBy": ResponsibleDepartmentCtrlBy,
            "responsibleDepartment": {
                "fkOrgIdentity": {
                    "fkUuid": "15deb66c-1685-49ac-8344-cfbf84fe6d84",
                    "type": "Afdeling",
                    "fullName": "Digitalisering"
                }
            },
            "caseParties": [
                {
                    "index": index,
                    "identificationType": identificationType,
                    "identification": identification, 
                    "partyRole": partyRole,
                    "partyRoleName": partyRoleName, 
                    "participantRole": participantRole, 
                    "name": name 
                },
                {
                    "index": Index_Uuid,
                    "identificationType": "Frit",
                    "identification": IndsenderNavn,
                    "partyRole": "IND",
                    "partyRoleName": "Indsender",
                    "participantRole": "Sekundær",
                    "name": IndsenderNavn,
                    "participantContactInformation": IndsenderMail
                }
            ],
            "journalNotes": [
                {
                    "uuid": JournalUuid,
                    "approved": True,
                    "journalNoteAttributes":
                    {
                        "journalNoteDate": JournalDate, 
                        "title": link_text,
                        "editReasonApprovedJournalnote": "Oprettelse",
                        "journalNoteAuthor": "AKTBOB",
                        "author": {
                            "fkOrgIdentity": {
                                "fkUuid": "15deb66c-1685-49ac-8344-cfbf84fe6d84",
                                "type": "Afdeling",
                                "fullName": "Digitalisering"
                                }
                        },
                        "journalNoteType": "Bruger",
                        "format": "Ooxml",
                        "note":base64_JournalNote

                    }
                }
            ],
            "buildingCase": {
                "buildingCaseAttributes": {
                    "buildingCaseClassId": "2a33734b-c596-4edf-93eb-23daae4bfc3e",
                    "buildingCaseClassName": "Aktindsigt"
                },
                "propertyInformation":{
                    "cadastralId": CadastralId,
                    "bfeNumber": bfeNumber

                },
                "UserdefinedFields": [
                        {
                            "type": "1. Politisk kategori",
                            "value": "Aktindsigt"
                        }
                    ]
            }  
        }
        # Make the API request
        try:
            response = requests.post(url, headers=headers, json=payload)
            
            # Handle response
            if response.status_code == 200:
                print(response.text)
                orchestrator_connection.log_info(f"Request Successful. Status Code:{response.status_code}")
            else:
                print("Failed to send request. Status Code:", response.status_code)
                print("Response Data:", response.text)  # Print error response
        except Exception as e:
            raise Exception("Failed to fetch Sagstitel (Nova):", str(e))

    return {
    "out_Text": "Aktindsigtssagen er oprettet i Nova"
    }