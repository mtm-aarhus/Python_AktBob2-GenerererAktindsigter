def invoke_GenererSagsoversigt(Arguments_GenererSagsoversigt):
    import os
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.user_credential import UserCredential

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



    def sharepoint_client(RobotUserName, RobotPassword, SharePointURL) -> ClientContext:
        """
        Authenticate to SharePoint and return the client context.
        """
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

    def get_folders_from_sharepoint(client: ClientContext, overmappe_url: str):
        """
        Retrieve and print folder names within the specified Overmappe folder.
        """
        try:
            overmappe_folder = client.web.get_folder_by_server_relative_url(overmappe_url)
            client.load(overmappe_folder)
            client.execute_query()
            
            # Get the subfolders within the Overmappe folder
            subfolders = overmappe_folder.folders
            client.load(subfolders)
            client.execute_query()
            
            folder_names = [folder.properties["Name"] for folder in subfolders]
            
            print("Folders found:")
            for folder_name in folder_names:
                if " - " in folder_name:
                    folder_name = folder_name.split(" - ")[0]
                print(f" - {folder_name}")
            
            return folder_names
        except Exception as e:
            print(f"Error retrieving folders: {e}")
            raise

    # Main logic
    try:
        # Inputs

        site_relative_path = "/Teams/tea-teamsite10506/Delte Dokumenter"
    
        
        # Authenticate to SharePoint
        client = sharepoint_client(RobotUserName, RobotPassword, SharePointURL)
        
        # Construct Overmappe path
        overmappe_url = f"{site_relative_path}/Dokumentlister/{Overmappe}"
        print(f"Overmappe URL: {overmappe_url}")
        
        # Retrieve folder names
        folder_names = get_folders_from_sharepoint(client, overmappe_url)
        
    except Exception as e:
        print(f"An error occurred: {e}")




    
    return {
    "out_Text": folder_names,
    }