def invoke_GenerateNovaCase(Arguments_GenerateNovaCase):
    print("Opretter Nova Aktindsigtssagen")
    import uuid
    import requests
    import json
    
    # henter in_argumenter:
    KMDNovaURL = Arguments_GenerateNovaCase.get("in_KMDNovaURL")
    KMD_access_token = Arguments_GenerateNovaCase.get("in_NovaToken")
    Sagsnummer = Arguments_GenerateNovaCase.get("in_Sagsnummer")

    ### --- Henter caseinfo --- ###
    


    ### ---  Opretter sagen --- ####   

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
            "title": "Test gustav - Anmodning om aktindsigt",
            "caseDate": "2025-01-28T00:00:00+00:00",
            "description": "GO: https://google.com"
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
        "SensitivityCtrlBy": "Bruger",
        "AvailabilityCtrlBy": "Regler",
        "SecurityUnitCtrlBy": "Regler",
        "ResponsibleDepartmentCtrlBy": "Regler",
        "responsibleDepartment": {
            "fkOrgIdentity": {
                "fkUuid": "15deb66c-1685-49ac-8344-cfbf84fe6d84",
                "type": "Afdeling",
                "fullName": "Digitalisering"
            }
        },
        "caseParties": [
            {
                "index": "9c60ce1c-5f57-44ab-b805-44800017000c",
                "identificationType": "BfeNummer",
                "identification": "4248928",
                "partyRole": "PRI",
                "partyRoleName": "Primær sagspart",
                "participantRole": "Primær",
                "name": "Grøndalsvej 1A, 8260 Viby J"
            }
        ]
    }

    # Make the API request
    try:
        response = requests.put(url, headers=headers, json=payload)
        
        # Handle response
        if response.status_code == 200:
            print("Request Successful. Status Code:", response.status_code)
        else:
            print("Failed to send request. Status Code:", response.status_code)
            print("Response Data:", response.text)  # Print error response
    except Exception as e:
        print("Failed to fetch Sagstitel (Nova):", str(e))

    return {
    "Test": "Test"
    }