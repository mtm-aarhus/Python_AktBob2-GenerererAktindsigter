from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
def invoke_GenerateNovaCase(Arguments_GenerateNovaCase,orchestrator_connection: OrchestratorConnection):
    import uuid
    import requests
    import json
    import os
    from GetKmdAcessToken import GetKMDToken
    from datetime import datetime
    
     # henter in_argumenter:
    Sagsnummer = Arguments_GenerateNovaCase.get("in_Sagsnummer")
    KMDNovaURL = Arguments_GenerateNovaCase.get("in_KMDNovaURL")
    KMD_access_token = Arguments_GenerateNovaCase.get("in_NovaToken")
    AktSagsURL = Arguments_GenerateNovaCase.get("in_AktSagsURL")
 
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



    # ### ---  Opretter sagen --- ####   
    CurrentDate = datetime.now().strftime("%Y-%m-%dT00:00:00")
    TransactionID = str(uuid.uuid4())
    Uuid = str(uuid.uuid4())

    # Define API URL
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
            "title": f"Test gustav - Anmodning om aktindsigt i {Sagsnummer} - Se beskrivelse",#spørg Byggeri om det rigtige anvendes
            "caseDate": CurrentDate,
            "description": AktSagsURL
        },
        "caseClassification": {
            "kleNumber": {"code": "02.00.00"}, #Fast - men spørg Byggeri om det rigtige anvendes
            "proceedingFacet": {"code": "A53"}#Fast - men spørg Byggeri om det rigtige anvendes
        },
        "state": progressState, 
        "sensitivity": "Følsomme",# Fast
        "caseworker": { #Fast - men spørg Byggeri om det rigtige anvendes
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
        "responsibleDepartment": {#Fast - men spørg Byggeri om det rigtige anvendes
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
            }
        ]
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