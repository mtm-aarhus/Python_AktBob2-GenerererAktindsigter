import os
import requests
from msal import PublicClientApplication

def upload_file_to_sharepoint(
    site_url,
    overmappe,
    undermappe,
    file_path,
    sharepoint_app_id,
    sharepoint_tenant,
    robot_username,
    robot_password
):
    try:
        # Build SharePoint folder path and extract file name
                # Determine the target SharePoint folder path
        if undermappe:  # If undermappe is provided, include it in the path
            sharepoint_folder_path = f"/Aktindsigter/{overmappe}/{undermappe}"
        else:  # Otherwise, upload only to overmappe
            sharepoint_folder_path = f"/Aktindsigter/{overmappe}"

        file_name = os.path.basename(file_path)

        # Normalize the site URL
        if site_url.startswith("https://"):
            site_url = site_url[8:]
        site_url = site_url.replace(".sharepoint.com", ".sharepoint.com:")

        scopes = ["https://graph.microsoft.com/.default"]

        # Authenticate and acquire access token
        msal_app = PublicClientApplication(
            client_id=sharepoint_app_id,
            authority=f"https://login.microsoftonline.com/{sharepoint_tenant}"
        )

        token_response = msal_app.acquire_token_by_username_password(
            username=robot_username,
            password=robot_password,
            scopes=scopes
        )
        if "access_token" not in token_response:
            raise Exception(f"Failed to acquire token: {token_response}")
        access_token = token_response["access_token"]
        headers = {"Authorization": f"Bearer {access_token}"}

        # Get the site ID
        site_request_url = f"https://graph.microsoft.com/v1.0/sites/{site_url}"
        site_response = requests.get(site_request_url, headers=headers)
        site_response.raise_for_status()
        site_id = site_response.json()["id"]

        # Get the drive ID
        drive_request_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        drive_response = requests.get(drive_request_url, headers=headers)
        drive_response.raise_for_status()
        drive_id = drive_response.json()["id"]

        # Direct file upload
        drive_item_request_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_folder_path}/{file_name}:/content"
        with open(file_path, "rb") as file_stream:
            upload_response = requests.put(
                drive_item_request_url,
                headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/octet-stream"},
                data=file_stream
            )

        if not upload_response.ok:
            raise Exception(f"File upload failed: {upload_response.text}")

        print("Upload complete")

    except Exception as ex:
        print(f"Error: {ex}")
        print("Filen kunne ikke overføres, prøver chunk upload")

        try:
            # Create upload session
            upload_session_request_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_folder_path}/{file_name}:/createUploadSession"
            upload_session_body = {
                "@microsoft.graph.conflictBehavior": "replace",
                "name": file_name
            }
            upload_session_response = requests.post(
                upload_session_request_url,
                headers=headers,
                json=upload_session_body
            )
            if not upload_session_response.ok:
                raise Exception(f"Failed to create upload session: {upload_session_response.text}")
            upload_url = upload_session_response.json()["uploadUrl"]

            # Upload file in chunks
            with open(file_path, "rb") as file_stream:
                total_length = os.path.getsize(file_path)
                max_slice_size = 320 * 16384  # Max chunk size
                bytes_remaining = total_length
                slice_start = 0

                while bytes_remaining > 0:
                    # Determine chunk size
                    slice_size = min(max_slice_size, bytes_remaining)
                    file_stream.seek(slice_start)
                    slice_data = file_stream.read(slice_size)

                    # Prepare headers for chunk upload
                    chunk_headers = {
                        "Content-Range": f"bytes {slice_start}-{slice_start + slice_size - 1}/{total_length}",
                        "Content-Type": "application/octet-stream"
                    }

                    # Upload the chunk
                    slice_response = requests.put(upload_url, headers=chunk_headers, data=slice_data)
                    if not slice_response.ok:
                        raise Exception(f"Slice upload failed: {slice_response.text}")

                    # Update progress
                    bytes_remaining -= slice_size
                    slice_start += slice_size
                    print(f"Uploaded {slice_start} bytes of {total_length} bytes")

            print("Chunk upload complete")

        except Exception as chunk_ex:
            print(f"Chunk upload failed: {chunk_ex}")