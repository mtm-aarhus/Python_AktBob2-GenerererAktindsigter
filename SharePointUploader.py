import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

def upload_file_to_sharepoint(site_url,Overmappe, Undermappe, file_path, RobotUserName, RobotPassword):
    try:
        # Authenticate
        credentials = UserCredential(RobotUserName, RobotPassword)
        ctx = ClientContext(site_url).with_credentials(credentials)

        # Construct folder path
        # Construct folder path
        if Undermappe:
            folder_url = f"/Teams/tea-teamsite10506/Delte Dokumenter/Aktindsigter/{Overmappe}/{Undermappe}"
        else:
            folder_url = f"/Teams/tea-teamsite10506/Delte Dokumenter/Aktindsigter/{Overmappe}"

        folder = ctx.web.get_folder_by_server_relative_path(folder_url)
        ctx.load(folder)
        ctx.execute_query()

        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)

        # Attempt normal upload for files ≤ 4MB
        if file_size <= 50 * 1024*1024:  
            try:
                with open(file_path, "rb") as f:
                    file = folder.files.add(file_name, f.read(), overwrite=True).execute_query()
                print(f"File uploaded successfully using normal upload.")
                return  # 
            except Exception as e:
                print(f"Normal upload failed for '{file_name}', switching to chunked upload. Error: {e}")

        size_chunk = 1000000  # 1MB chunks
        def print_upload_progress(offset):
            print(f"Uploaded {offset} bytes of {file_size} ({round(offset / file_size * 100, 2)}%)")

        with open(file_path, "rb") as f:
            uploaded_file = folder.files.create_upload_session(f, size_chunk, print_upload_progress).execute_query()

        print(f"File uploaded successfully using chunked upload.")

    except Exception as e:
        raise Exception(f"Failed to upload file '{file_path}': {e}")
