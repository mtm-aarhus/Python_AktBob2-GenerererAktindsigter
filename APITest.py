import requests
import os
from requests_ntlm import HttpNtlmAuth
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import json


orchestrator_connection = OrchestratorConnection("GOTest", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
API_url = orchestrator_connection.get_constant("GOApiTESTURL").value
go_credentials = orchestrator_connection.get_credential("GOTestApiUser")
session = requests.Session()
session.auth = HttpNtlmAuth(go_credentials.username, go_credentials.password)


payload = json.dumps({
  "CaseTypePrefix": "GEO",
  "MetadataXml": "<z:row xmlns:z=\"#RowsetSchema\" ows_Title=\"Case From Api Gustav\" ows_CaseStatus=\"Åben\" />",
  "ReturnWhenCaseFullyCreated": True
})

headers = {
  'Content-Type': 'application/json',
}

URL = API_url+"/geosager/_goapi/Cases"
response = session.post(URL, data=payload, headers=headers, timeout=500)

# Handle the response
if response.status_code == 200:
    # Check if the response is JSON
    if 'application/json' in response.headers.get('Content-Type', ''):
        try:
            print("Request successful:", response.json())
        except json.JSONDecodeError:
            print("Request successful, but response is not in JSON format:", response.text)
    else:
        print("Request successful, but response is not JSON:", response.text)
else:
    print("Request failed:", response.status_code, response.text)
