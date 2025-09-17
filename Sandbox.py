import os
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from GetKmdAcessToken import GetKMDToken
from GetFilarkivAcessToken import GetFilarkivToken
import GetDocumentList
import requests
from requests_ntlm import HttpNtlmAuth
import GenerateCaseFolder
import PrepareEachDocumentToUpload
import GenerateAndUploadAktlistePDF
import GenerererSagsoversigt
import GenerateNovaCase
import json
from SendSMTPMail import send_email, EmailAttachment  # Import the function and dataclass

#   ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("AktbobGenererAktindsigter", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)


#Certification stuff added after sharepoint problems
certification = orchestrator_connection.get_credential("SharePointCert")
api = orchestrator_connection.get_credential("SharePointAPI")
tenant =  api.username
client_id =  api.password
thumbprint = certification.username
cert_path = certification.password

orchestrator_connection.log_trace("Running process.")
GraphAppIDAndTenant = orchestrator_connection.get_credential("GraphAppIDAndTenant")
SharePointAppID = GraphAppIDAndTenant.username
SharePointTenant = GraphAppIDAndTenant.password
SharePointURL = orchestrator_connection.get_constant("AktbobSharePointURL").value
CloudConvert = orchestrator_connection.get_credential("CloudConvertAPI")
CloudConvertAPI = CloudConvert.password
UdviklerMailAktbob =  orchestrator_connection.get_constant("UdviklerMailAktbob").value
RobotCredentials = orchestrator_connection.get_credential("RobotCredentials")
RobotUserName = RobotCredentials.username
RobotPassword = RobotCredentials.password
GOAPILIVECRED = orchestrator_connection.get_credential("GOAktApiUser")
GoUsername = GOAPILIVECRED.username
GoPassword = GOAPILIVECRED.password
KMDNovaURL = orchestrator_connection.get_constant("KMDNovaURL").value
FilarkivURL = orchestrator_connection.get_constant("FilarkivURL").value
AktbobAPI = orchestrator_connection.get_credential("AktbobAPIKey")
AktbobAPIKey = AktbobAPI.password


# ---- Henter access tokens ----
KMD_access_token = GetKMDToken(orchestrator_connection)
Filarkiv_access_token = GetFilarkivToken(orchestrator_connection)


# ---- Deffinerer Go-session ----
def GO_Session(GoUsername, GoPassword):
    session = requests.Session()
    session.auth = HttpNtlmAuth(GoUsername, GoPassword)
    session.headers.update({
        "Content-Type": "application/json"
    })
    return session

# ---- Initialize Go-Session ----
go_session = GO_Session(GoUsername, GoPassword)



    #---- Henter kø-elementer ----
queue = json.loads("""""")

Sagsnummer = queue.get("Sagsnummer")
MailModtager = queue.get("MailModtager")
DeskProID = queue.get("DeskProID")
DeskProTitel = queue.get("DeskProTitel")
PodioID = queue.get("PodioID")
Overmappe = queue.get("Overmappe")
Undermappe = queue.get("Undermappe")
GeoSag = queue.get("GeoSag")
NovaSag = queue.get("NovaSag")
AktSagsURL = queue.get("AktSagsURL")
IndsenderNavn =  queue.get("IndsenderNavn")
IndsenderMail = queue.get("IndsenderMail")
AktindsigtsDato = queue.get("AktindsigtsDato")

sender = "aktbob@aarhus.dk" 
smtp_server = "smtp.adm.aarhuskommune.dk"   
smtp_port = 25               

orchestrator_connection.log_info(f"Sagsnummer: {Sagsnummer}")
orchestrator_connection.log_info(f"DeskProID: {DeskProID}")
orchestrator_connection.log_info(f"DeskProTitel: {DeskProTitel}")
orchestrator_connection.log_info(f"PodioID: {PodioID}")
orchestrator_connection.log_info(f"Overmappe: {Overmappe}")
orchestrator_connection.log_info(f"Undermappe: {Undermappe}")
orchestrator_connection.log_info(f"GeoSag: {GeoSag}")
orchestrator_connection.log_info(f"NovaSag: {NovaSag}")
orchestrator_connection.log_info(f"AktSagsURL: {AktSagsURL}")
orchestrator_connection.log_info(f"IndsenderNavn: {IndsenderNavn}")
orchestrator_connection.log_info(f"IndsenderMail : {IndsenderMail }")
orchestrator_connection.log_info(f"AktindsigtsDato: {AktindsigtsDato}")

# ---- Run "GetDokumentlist" ----
Arguments = {
    "in_RobotUserName": RobotUserName,
    "in_RobotPassword": RobotPassword,
    "in_Sagsnummer": Sagsnummer,
    "in_SharePointUrl": SharePointURL,
    "in_Overmappe": Overmappe,
    "in_Undermappe": Undermappe,
    "in_GeoSag": GeoSag, 
    "in_NovaSag": NovaSag,
    "GoUsername": GoUsername,
    "GoPassword":  GoPassword,
    "KMD_access_token": KMD_access_token,
    "KMDNovaURL": KMDNovaURL,
    "in_MailModtager": MailModtager, 
    "tenant":  tenant,
    "client_id":  client_id,
    "thumbprint": thumbprint,
    "cert_path": cert_path,
}

# ---- Run "GetDocumentList" ----
GetDocumentList_Output_arguments = GetDocumentList.invoke(Arguments, go_session,orchestrator_connection)
Sagstitel = GetDocumentList_Output_arguments.get("sagstitel")
print(f"Sagstitel: {Sagstitel}")
dt_DocumentList = GetDocumentList_Output_arguments.get("dt_DocumentList")
DokumentlisteDatoString = GetDocumentList_Output_arguments.get("out_DokumentlisteDatoString")

if dt_DocumentList.empty:
    print("Number of rows:",len(dt_DocumentList))
        ###---- Send mail til sagsansvarlig ----####

    # Define email details
    
    subject = f"{Sagsnummer} er en tom sag"
    body = f"""Sagen: {Sagsnummer} er en tom sag. Vær opmærksom på, at processen ikke kan behandle tomme sager.<br><br>
    Det anbefales at følge <a href="https://aarhuskommune.atlassian.net/wiki/spaces/AB/pages/64979049/AKTBOB+--+Vejledning">vejledningen</a>, 
    hvor du også finder svar på de fleste spørgsmål og fejltyper.
    """

    # Call the send_email function
    send_email(
        receiver=MailModtager,
        sender=sender,
        subject=subject,
        body=body,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        html_body=True
    )
    raise Exception("return")
else:
    print("Number of rows:",len(dt_DocumentList))
    
# --- Validate: Omfattet = "Ja" but Aktstatus is blank ---

col_omf = "Omfattet af ansøgningen? (Ja/Nej)"
col_akt = "Gives der aktindsigt i dokumentet? (Ja/Nej/Delvis)"

# Normalize Omfattet
dt_DocumentList[col_omf] = (
    dt_DocumentList[col_omf]
    .astype(str)
    .str.strip()
    .str.lower()
    .replace({"nan": ""})
)

# Aktstatus: keep NaN/blank detection separate
dt_DocumentList[col_akt] = dt_DocumentList[col_akt].astype(str).str.strip()
dt_DocumentList[col_akt] = dt_DocumentList[col_akt].replace({"nan": ""})

# Condition: Omfattet == "ja" and Aktstatus is blank
mask_blank = (dt_DocumentList[col_omf] == "ja") & (dt_DocumentList[col_akt] == "")
conflicts = dt_DocumentList.loc[mask_blank].copy()

if not conflicts.empty:
    print(f"{len(conflicts)} dokument(er) er omfattet, men Aktstatus er blank.")
    cols_to_show = [c for c in ["Dok ID", "Dokumenttitel", col_omf, col_akt] if c in conflicts.columns]
    print(conflicts[cols_to_show].to_string(index=False))
    subject = f"{Sagsnummer}: Dokumentliste mangler udfyldning"
    body = f"""Sag: <a href="https://mtmsager.aarhuskommune.dk/app#/t/ticket/{DeskProID}">{DeskProID} - {DeskProTitel}</a><br><br>
    Dokumentlisten har {len(conflicts)} række(r) hvor dokumenter har 'Ja' i 'Omfattet af ansøgningen? (Ja/Nej)', men der er ikke valgt noget i 
    'Gives der aktindsigt i dokumentet? (Ja/Nej/Delvis)'. Sørg for at alle rækker der er omfattet af ansøgningen har et svar om hvorvidt der gives aktindsigt, og genkør herefter processen i Podio.<br><br>
    Det anbefales at følge <a href="https://aarhuskommune.atlassian.net/wiki/spaces/AB/pages/64979049/AKTBOB+--+Vejledning">vejledningen</a>, 
    hvor du også finder svar på de fleste spørgsmål og fejltyper.
    """
    send_email(
        receiver=MailModtager,
        sender=sender,
        subject=subject,
        body=body,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        html_body=True
    )
    raise Exception("return")

else:
    print("Ingen dokumenter med Omfattet='Ja' og tom Aktstatus.")


# ---- Run "GenerateCaseFolder" ----
Arguments_GenerateCaseFolder = {
    "in_Sagsnummer": Sagsnummer,
    "in_RobotUserName": RobotUserName,
    "in_RobotPassword": RobotPassword,
    "in_SharePointAppID": SharePointAppID,
    "in_SharePointTenant":SharePointTenant,
    "in_SharePointUrl": SharePointURL,
    "in_Overmappe": Overmappe,
    "in_Undermappe": Undermappe,
    "in_DeskProTitel": DeskProTitel,
    "in_DeskProID": DeskProID,
    "in_Filarkiv_access_token": Filarkiv_access_token,
    "in_Sagstitel": Sagstitel,
    "in_FilarkivURL": FilarkivURL,
    "tenant":  tenant,
    "client_id":  client_id,
    "thumbprint": thumbprint,
    "cert_path": cert_path,
}

GenerateCaseFolder_Output_arguments = GenerateCaseFolder.invoke_GenerateCasefolder(Arguments_GenerateCaseFolder,orchestrator_connection)
FilarkivCaseID = GenerateCaseFolder_Output_arguments.get("out_FilarkivCaseID")
orchestrator_connection.log_info(f"FilarkivCaseID: {FilarkivCaseID}")

###---- Send mail til sagsansvarlig ----####

# Define email details
subject = f"{Sagsnummer}: Screening igangsat"
body = f"""Sag: <a href="https://mtmsager.aarhuskommune.dk/app#/t/ticket/{DeskProID}">{DeskProID} - {DeskProTitel}</a><br><br>
Robotten er nu gået i gang med screening af dokumenterne.<br><br>
Procestiden varierer afhængigt af antallet af dokumenter. Du vil modtage en mail når dokumenterne er klar.<br><br>
Det anbefales at følge <a href="https://aarhuskommune.atlassian.net/wiki/spaces/AB/pages/64979049/AKTBOB+--+Vejledning">vejledningen</a>, 
hvor du også finder svar på de fleste spørgsmål og fejltyper.
"""
        
# Call the send_email function
send_email(
    receiver=MailModtager,
    sender=sender,
    subject=subject,
    body=body,
    smtp_server=smtp_server,
    smtp_port=smtp_port,
    html_body=True
)


# ---- Run "PrepareEachDocumentToUpload" ----
Arguments_PrepareEachDocumentToUpload = {
    "in_dt_Documentlist": dt_DocumentList,
    "in_CloudConvertAPI": CloudConvertAPI,
    "in_MailModtager": MailModtager,
    "in_UdviklerMail": UdviklerMailAktbob,
    "in_RobotUserName": RobotUserName,
    "in_RobotPassword": RobotPassword,
    "in_Filarkiv_access_token": Filarkiv_access_token,
    "in_FilarkivCaseID": FilarkivCaseID,
    "in_SharePointAppID": SharePointAppID,
    "in_SharePointTenant":SharePointTenant,
    "in_SharePointUrl": SharePointURL,
    "in_Overmappe": Overmappe,
    "in_Undermappe": Undermappe,
    "in_Sagsnummer": Sagsnummer,
    "in_GeoSag": GeoSag,
    "in_NovaSag": NovaSag,
    "in_FilarkivURL": FilarkivURL,
    "in_NovaToken": KMD_access_token,
    "in_KMDNovaURL": KMDNovaURL,
    "in_GoUsername": GoUsername,
    "in_GoPassword":  GoPassword,
    "in_DeskProTitel": DeskProTitel,
    "in_DeskProID": DeskProID,
    "tenant":  tenant,
    "client_id":  client_id,
    "thumbprint": thumbprint,
    "cert_path": cert_path,
}

PrepareEachDocumentToUpload_Output_arguments = PrepareEachDocumentToUpload.invoke_PrepareEachDocumentToUpload(Arguments_PrepareEachDocumentToUpload,orchestrator_connection)
dt_AktIndex = PrepareEachDocumentToUpload_Output_arguments.get("out_dt_AktIndex")

# ---- Run "Generate&UploadAktlistPDF" ----

Arguments_GenerateAndUploadAktlistePDF = {
"in_dt_AktIndex": dt_AktIndex,
"in_Sagsnummer": Sagsnummer,
"in_DokumentlisteDatoString":DokumentlisteDatoString, 
"in_RobotUserName": RobotUserName,
"in_RobotPassword": RobotPassword,
"in_SagsTitel": Sagstitel,
"in_SharePointAppID": SharePointAppID,
"in_SharePointTenant": SharePointTenant,
"in_Overmappe": Overmappe,
"in_Undermappe": Undermappe,
"in_SharePointURL": SharePointURL,
"in_GoUsername":GoUsername,
"in_GoPassword": GoPassword,
"tenant":  tenant,
"client_id":  client_id,
"thumbprint": thumbprint,
"cert_path": cert_path,
}

GenerateAndUploadAktlistePDF_Output_arguments = GenerateAndUploadAktlistePDF.invoke_GenerateAndUploadAktlistePDF(Arguments_GenerateAndUploadAktlistePDF,orchestrator_connection)
Test = GenerateAndUploadAktlistePDF_Output_arguments.get("out_Text")
print(Test)


# ---- Run "GenererSagsoversigt" ----
Arguments_GenererSagsoversigt = {
"in_RobotUserName": RobotUserName,
"in_RobotPassword": RobotPassword,
"in_MailModtager": MailModtager,
"in_SharePointAppID": SharePointAppID,
"in_SharePointTenant": SharePointTenant,
"in_SharePointURL": SharePointURL,
"in_Sagsnummer": Sagsnummer,
"in_SagsTitel": Sagstitel,
"in_Overmappe": Overmappe,
"in_Undermappe": Undermappe,
"in_GoUsername":GoUsername,
"in_GoPassword": GoPassword,
"in_NovaToken": KMD_access_token,
"in_KMDNovaURL": KMDNovaURL,
"tenant":  tenant,
"client_id":  client_id,
"thumbprint": thumbprint,
"cert_path": cert_path,
}

GenererSagsoversigt_Output_arguments = GenerererSagsoversigt.invoke_GenererSagsoversigt(Arguments_GenererSagsoversigt,orchestrator_connection)
Test = GenererSagsoversigt_Output_arguments.get("out_Text")
print(Test)



# if NovaSag == True: 
#     # ---- Run "GenerateNovaCase" ----
#     Arguments_GenerateNovaCase = {
#     "in_Sagsnummer": Sagsnummer,
#     "in_NovaToken": KMD_access_token,
#     "in_KMDNovaURL": KMDNovaURL,
#     "in_AktSagsURL": AktSagsURL,
#     "in_IndsenderNavn": IndsenderNavn,
#     "in_IndsenderMail" : IndsenderMail,
#     "in_AktindsigtsDato": AktindsigtsDato,
#     "in_DeskProID": DeskProID
#     }

#     GenerateNovaCase_Output_arguments = GenerateNovaCase.invoke_GenerateNovaCase(Arguments_GenerateNovaCase,orchestrator_connection)
#     Test = GenerateNovaCase_Output_arguments.get("out_Text")
#     print(Test)


# ---- Run "SendFilarkivCaseId&PodioIDToPodio"
# Define the API endpoint
try:
    url = "https://aktbob-external-api.grayglacier-2d22de15.northeurope.azurecontainerapps.io/Api/Jobs/CheckOCRScreeningStatus"

    # Define headers
    headers = {
        "ApiKey": AktbobAPIKey,  # Ensure ApiKey is defined
        "Content-Type": "application/json"
    }

    # Define JSON body
    json_body = {
        "filArkivCaseId": FilarkivCaseID,  # Ensure FilarkivCaseID is defined
        "podioItemId": PodioID  # Ensure PodioID is defined
    }

    # Make the POST request
    response = requests.post(url, headers=headers, json=json_body)

    # Handle response

    print("Response Status:", response.status_code)
    print("Response:", response.text)
except requests.exceptions.RequestException as e:
    print(f"Error making request: {e}")
    raise Exception(f"Request to API failed: {e}")
