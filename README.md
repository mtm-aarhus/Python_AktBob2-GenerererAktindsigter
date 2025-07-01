
# 📄 README

## Aktindsigt Automation Robot

**Aktindsigt Automation** is a comprehensive robot for **Teknik og Miljø, Aarhus Kommune**. It processes access-to-records requests (aktindsigt) by retrieving case documents, generating overviews and indexes, storing files in SharePoint and FilArkiv, and notifying stakeholders automatically.

---

## 🚀 Features

✅ **Case Data Retrieval**  
- Integrates with KMD Nova and Geo to fetch metadata and document lists  
- Retrieves deskpro ticket information to enrich context

📤 **Document Preparation and Conversion**  
- Downloads files from SharePoint and KMD  
- Converts documents to PDF using CloudConvert if needed  
- Generates Excel indexes (Aktliste) and case overviews (Sagsoversigt)

🗂️ **Structured Folder Management**  
- Creates case folders in FilArkiv and SharePoint  
- Ensures consistent naming conventions and folder hierarchy  

📡 **Automated Uploads**  
- Uploads all processed documents to SharePoint libraries  
- Registers documents in FilArkiv with metadata  

📧 **Notifications and Alerts**  
- Emails caseworkers if document lists are empty or processing fails  
- Sends confirmations after successful processing

🔐 **Credential and Token Management**  
- Securely fetches and refreshes API tokens (KMD, FilArkiv) via OpenOrchestrator  
- All credentials are stored encrypted  

---

## 🧭 Process Flow

1. **Token Management**
   - Fetches or refreshes KMD and FilArkiv access tokens (`GetKmdAcessToken.py`, `GetFilarkivAcessToken.py`)
2. **Document List Retrieval**
   - Retrieves a list of all documents for the requested case (`GetDocumentList.py`)
3. **Validation**
   - If no documents are found, notifies the requester by email and exits
4. **Case Folder Creation**
   - Creates or updates the case in FilArkiv (`GenerateCaseFolder.py`)
   - Generates a Nova case record (`GenerateNovaCase.py`)
5. **Document Preparation**
   - Downloads documents
   - Converts to PDF where needed (`PrepareEachDocumentToUpload.py`)
   - Prepares Excel index (`GenerateAndUploadAktlistePDF.py`)
   - Creates a case overview (`GenerererSagsoversigt.py`)
6. **Uploads**
   - Uploads all final documents to SharePoint (`SharePointUploader.py`)
   - Registers files in FilArkiv
7. **Cleanup and Confirmation**
   - Logs operations and sends notifications

---

## 🔐 Privacy & Security

- All APIs use HTTPS
- Credentials are managed in OpenOrchestrator
- No personal data is stored locally after processing
- Temporary files are deleted after upload

---

## ⚙️ Dependencies

- Python 3.10+
- `requests`
- `requests-ntlm`
- `pandas`
- `pyodbc`
- `python-docx`
- `openpyxl`
- `reportlab`
- `office365-rest-python-client`
- `CloudConvert`

---

## 👷 Maintainer

Gustav Chatterton  
*Digital udvikling, Teknik og Miljø, Aarhus Kommune*
