def invoke_GenererSagsoversigt(Arguments_GenererSagsoversigt):
    import os
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.user_credential import UserCredential
    import pandas as pd
    import requests
    import time
    from datetime import datetime
    from requests_ntlm import HttpNtlmAuth
    import re
    import json
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.platypus import SimpleDocTemplate, Table as ReportTable, TableStyle as ReportTableStyle, Paragraph, Frame, PageTemplate
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors as reportlab_colors
    from SharePointUploader import upload_file_to_sharepoint
    import uuid
    # henter in_argumenter:

    RobotUserName = Arguments_GenererSagsoversigt.get("in_RobotUserName")
    RobotPassword = Arguments_GenererSagsoversigt.get("in_RobotPassword")
    MailModtager = Arguments_GenererSagsoversigt.get("in_MailModtager")
    SharePointAppID = Arguments_GenererSagsoversigt.get("in_SharePointAppID")
    SharePointTenant = Arguments_GenererSagsoversigt.get("in_SharePointTenant")
    SharePointURL = Arguments_GenererSagsoversigt.get("in_SharePointURL")
    Sagsnummer = Arguments_GenererSagsoversigt.get("in_Sagsnummer")
    Sagstitel = Arguments_GenererSagsoversigt.get("in_SagsTitel")
    Overmappe = Arguments_GenererSagsoversigt.get("in_Overmappe")
    Undermappe = Arguments_GenererSagsoversigt.get("in_Undermappe")
    GoUsername = Arguments_GenererSagsoversigt.get("in_GoUsername")
    GoPassword = Arguments_GenererSagsoversigt.get("in_GoPassword")
    NovaToken = Arguments_GenererSagsoversigt.get("in_NovaToken")
    KMDNovaURL = Arguments_GenererSagsoversigt.get("in_KMDNovaURL")
    max_retries = 2  # Number of retry attempts
    print(RobotPassword)
    print(RobotPassword)
    print(SharePointURL)
    def sharepoint_client(RobotUserName, RobotPassword, SharePointURL) -> ClientContext:

        try:
            credentials = UserCredential(RobotUserName, RobotPassword)
            ctx = ClientContext(SharePointURL).with_credentials(credentials)
            
            # Load the SharePoint web to test the connection
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            
            return ctx
        except Exception as e:
            print(f"Authentication failed: {e}")
            raise
    
    #Følgende to hører sammen:
    def extract_case_info(metadata_json):

        try:
            # Parse the JSON response
            metadata = json.loads(metadata_json).get("Metadata", "")

            # Extract Sagstitel (Case Title)
            title_match = re.search(r'ows_Title="([^"]+)"', metadata)
            Sagstitel = title_match.group(1) if title_match else "Unknown"

            # Extract Startdato (Case Date)
            date_match = re.search(r'ows_Modtaget="([^"]+)"', metadata)
            if date_match:
                raw_date = date_match.group(1).split(" ")[0]  # Extract only YYYY-MM-DD
                Startdato = datetime.strptime(raw_date, "%Y-%m-%d")  # Convert to DateTime
            else:
                Startdato = None  # Default to None if no date found

            return Sagstitel, Startdato

        except Exception as e:
            print(f"Error extracting metadata: {e}")
            return "Unknown", None
    def fetch_metadata(Sagsnummer):

        url = f"https://ad.go.aarhuskommune.dk/_goapi/Cases/Metadata/{Sagsnummer}"
        auth = HttpNtlmAuth(GoUsername, GoPassword)
        headers = {"Content-Type": "application/json"}

        for attempt in range(max_retries):
            try:
                response = requests.get(url, headers=headers, auth=auth, timeout=60)
                response.raise_for_status()  # Raise exception for HTTP errors

                metadata_json = response.text

                # Extract case information
                Sagstitel, Startdato = extract_case_info(metadata_json)

                return Sagstitel, Startdato  # Return extracted values

            except Exception as retry_exception:
                print(f"Retry {attempt + 1} for {Sagsnummer} failed: {retry_exception}")
                if attempt == max_retries - 1:
                    print(f"Failed to fetch metadata after {max_retries} retries for {Sagsnummer}.")
                    return "Unknown", None  # Return default values if all retries fail
                time.sleep(5)  # Wait before the next retry
    
    def GetNovaCaseinfo(Sagsnummer):
        TransactionID = str(uuid.uuid4())
        url = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"

        headers = {
            "Authorization": f"Bearer {NovaToken}",
            "Content-Type": "application/json"
        }

        payload = {
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
                "caseAttributes": {
                    "title": True,
                    "userFriendlyCaseNumber": True,
                    "caseDate": True
                }
            }
        }

        try:
            response = requests.put(url, headers=headers, json=payload)
            print("Nova API Response:", response.status_code)

            if response.status_code == 200:
                case_data = response.json().get('cases', [{}])[0].get('caseAttributes', {})
                sagstitel = case_data.get('title', 'N/A')
                Startdato = case_data.get('caseDate', None)

                # Convert Startdato from string to datetime object
                if Startdato:
                    try:
                        Startdato = datetime.fromisoformat(Startdato)  # Convert from 'YYYY-MM-DDTHH:MM:SS' to datetime
                    except ValueError:
                        print(f"Invalid date format for Startdato: {Startdato}")
                        Startdato = None  # If format is wrong, set to None

            else:
                print("Failed to fetch Sagstitel from NOVA. Status Code:", response.status_code)
                sagstitel, Startdato = None, None
        except Exception as e:
            print("Failed to fetch Sagstitel (Nova):", str(e))
            sagstitel, Startdato = None, None

        return sagstitel, Startdato

    def get_folders_from_sharepoint(client: ClientContext, overmappe_url: str):

        try:

            # Get Overmappe folder
            overmappe_folder = client.web.get_folder_by_server_relative_url(overmappe_url)
            client.load(overmappe_folder)
            client.execute_query()

            # Get all subfolders
            subfolders = overmappe_folder.folders
            client.load(subfolders)
            client.execute_query()

            data_table = []

            for folder in subfolders:
                client.load(folder)  # Ensure folder data is loaded
                client.execute_query()

                Sagsnummer = folder.properties.get("Name", "Unknown")

                if " - " in Sagsnummer:
                    Sagsnummer = Sagsnummer.split(" - ")[0]

                geosag_pattern = r"^[A-Z]{3}-\d{4}-\d{6}$"
                novasag_pattern = r"^[A-Za-z]\d{4}-\d{1,10}$"

                # Fetch metadata using the folder name
                if re.fullmatch(geosag_pattern, Sagsnummer):
                    Sagstitel, Startdato = fetch_metadata(Sagsnummer)
                elif re.fullmatch(novasag_pattern, Sagsnummer):
                    Sagstitel, Startdato = GetNovaCaseinfo(Sagsnummer)
                else:
                    print("Det er hverken et geosagsnummer eller Novasagsnummer")

                data_table.append({
                    "Sagsnummer": Sagsnummer,  
                    "Sagstitel": Sagstitel,
                    "Startdato": Startdato  
                })

                # Convert to DataFrame
                dt_Sagsliste = pd.DataFrame(data_table)

                # Sort by Startdato (most recent first), keeping missing values (`None`) at the bottom
                dt_Sagsliste = dt_Sagsliste.sort_values(by="Startdato", ascending=False, na_position='last')

                # Convert date back to `dd-MM-yyyy` format for final output
                dt_Sagsliste["Startdato"] = dt_Sagsliste["Startdato"].apply(lambda x: x.strftime("%d-%m-%Y") if pd.notna(x) else "Unknown")

            return dt_Sagsliste
        except Exception as e:
            print(f"Error retrieving folders: {e}")
            return None
    
    #Følgende to høre sammen:
    def wrap_text(text, max_chars):
    
        if pd.isna(text): 
            return ""
        if not isinstance(text, str):
            text = str(text)
        words = text.split()
        wrapped_lines = []
        line = ""
        for word in words:
            if len(line) + len(word) + 1 <= max_chars:
                line += " " + word if line else word
            else:
                wrapped_lines.append(line)
                line = word
        if line:
            wrapped_lines.append(line)
        return "<br/>".join(wrapped_lines)
    def dataframe_to_pdf(df, image_path, output_pdf_path, sags_id, my_date_string):

        # PDF Setup
        page_width, page_height = landscape(A4)
        margin = 40

        # Define styles
        styles = getSampleStyleSheet()

        header_style = ParagraphStyle(
            'header_style',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=10,
            textColor=reportlab_colors.white,
            alignment=1,  # CENTER
            leading=12,
            spaceAfter=5,
        )

        cell_style = ParagraphStyle(
            'cell_style',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=8,
            textColor=reportlab_colors.black,
            alignment=1,  # CENTER
            leading=10,
            spaceAfter=2,
        )

        # Column configuration
        column_widths = [200, 365, 200]  # Adjust width for better layout
        char_limits = [15, 50, 15]  # Define character limits per column

        headers = ["Sagsnummer", "Sagstitel", "Startdato"]

        # Create header row
        table_data = [[Paragraph(header, header_style) for header in headers]]

        # Add data rows
        for _, row in df.iterrows():
            table_row = [
                Paragraph(wrap_text(row.get("Sagsnummer", ""), char_limits[0]), cell_style),
                Paragraph(wrap_text(row.get("Sagstitel", ""), char_limits[1]), cell_style),
                Paragraph(wrap_text(row.get("Startdato", ""), char_limits[2]), cell_style),
            ]
            table_data.append(table_row)

        # Create table
        report_table = ReportTable(table_data, colWidths=column_widths)
        report_table.setStyle(ReportTableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), reportlab_colors.HexColor("#3661D8")),
            ('GRID', (0, 0), (-1, -1), 1, reportlab_colors.black),
            ('BOX', (0, 0), (-1, -1), 1, reportlab_colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))

        # Define first page layout
        def first_page(canvas, doc):
            canvas.saveState()
            
            # Draw Image
            image_width = 100
            image_height = 45
            image_x = margin
            image_y = page_height - margin - image_height
            canvas.drawImage(image_path, image_x, image_y, width=image_width, height=image_height)

            # Add Title
            title = f"SAGSOVERSIGT"
            canvas.setFont("Helvetica-Bold", 14)
            title_y = image_y - 20  # Title below the image
            canvas.drawString(margin, title_y, title)
            
            # Measure the width of the title
            title_width = canvas.stringWidth(title, "Helvetica-Bold", 14)

            # Add Black Line **RIGHT BELOW** the Title
            line_y = title_y - 5  # 5 units below the title
            canvas.setStrokeColor(reportlab_colors.black)
            canvas.setLineWidth(1)
            canvas.line(margin, line_y, margin + title_width, line_y) 
            
            # Add Date
            date_string = f"Dato for sagsoversigt: {my_date_string}"
            canvas.setFont("Helvetica", 10)
            text_width = canvas.stringWidth(date_string, "Helvetica", 10)
            canvas.drawString(page_width - margin - text_width, image_y, date_string)

            canvas.restoreState()

        # Define subsequent pages
        def later_pages(canvas, doc):
            canvas.saveState()
            canvas.restoreState()

        # PDF document setup
        doc = SimpleDocTemplate(output_pdf_path, pagesize=landscape(A4),
                                leftMargin=margin, rightMargin=margin,
                                topMargin=margin, bottomMargin=margin)

        # Reserve space at the top for image, title, and date
        table_start_y = page_height - margin - 100  # Adjusted Y position to avoid overlap

        # Define frames (where table goes)
        frame_first_page = Frame(margin, margin, page_width - 2 * margin, table_start_y - margin, id='first_page_table_frame')
        frame_later_pages = Frame(margin, margin, page_width - 2 * margin, page_height - 2 * margin, id='later_page_table_frame')

        # Define page templates
        first_page_template = PageTemplate(id='FirstPage', frames=frame_first_page, onPage=first_page)
        later_page_template = PageTemplate(id='LaterPages', frames=frame_later_pages, onPage=later_pages)
        
        doc.addPageTemplates([first_page_template, later_page_template])

        # Build the PDF with the table content
        doc.build([report_table])

        print(f"PDF saved to {output_pdf_path}")



    # Main logic
    try:
        # Inputs
        site_relative_path = "/Teams/tea-teamsite10506/Delte Dokumenter"

        # Authenticate to SharePoint
        client = sharepoint_client(RobotUserName, RobotPassword, SharePointURL)

        # Construct Overmappe path
        overmappe_url = f"{site_relative_path}/Aktindsigter/{Overmappe}"
        

        dt_Sagsliste = get_folders_from_sharepoint(client, overmappe_url)

        if dt_Sagsliste is not None and not dt_Sagsliste.empty:
            downloads_path = os.path.join("C:\\Users", os.getlogin(), "Downloads")
            image_path = os.path.join(os.getcwd(), "aak.jpg")
            PDFAktlisteFilnavn = f"Sagsoversigt.pdf"
            output_pdf_path = os.path.join(downloads_path, PDFAktlisteFilnavn)

            dataframe_to_pdf(dt_Sagsliste, image_path, output_pdf_path, Overmappe, datetime.today().strftime("%d-%m-%Y"))


        else:
            print("No data found. PDF generation skipped.")

    except Exception as e:
        print(f"An error occurred: {e}")


    # Upload Excel to Sharepoint
    upload_file_to_sharepoint(
        site_url=SharePointURL,
        overmappe=Overmappe,
        undermappe="",
        file_path=output_pdf_path,
        sharepoint_app_id=SharePointAppID,
        sharepoint_tenant=SharePointTenant,
        robot_username=RobotUserName,
        robot_password=RobotPassword
    )

    #Deleting local files: 
    os.remove(output_pdf_path)




    return {
        "out_Text": f"Det virker",
    }
