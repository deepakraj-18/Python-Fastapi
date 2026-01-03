import msal
import requests
import io
from typing import Dict, Any, Optional
import os
from datetime import datetime
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config.config import settings

class SharePointUtils:
    def __init__(self):
        self.tenant_id = settings.TENANT_ID
        self.client_id = settings.CLIENT_ID
        self.client_secret = settings.CLIENT_SECRET
        self.drive_id = settings.DRIVE_ID
        self.template_path = settings.TEMPLATE_PATH_ON_SP
        self.output_path = settings.OUTPUT_PATH_ON_SP
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scopes = ["https://graph.microsoft.com/.default"]
        self.site_url = settings.SITE_URL
        self.site_id = settings.SITE_ID

    def get_access_token(self):
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret
        )
        result = app.acquire_token_for_client(scopes=self.scopes)
        if "access_token" not in result:
            raise Exception(f"Token Error: {result.get('error_description')}")
        return result["access_token"]

    def get_file_metadata(self, file_id: str) -> Dict[str, Any]:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{file_id}"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Failed to get file metadata: {response.text}")
        
        return response.json()

    def download_file_by_path(self, sharepoint_path: str) -> io.BytesIO:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:{sharepoint_path}:/content"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            return self._download_file_alternative(sharepoint_path, token)
        
        return io.BytesIO(response.content)

    def download_file_by_path_with_drive(self, sharepoint_path: str, drive_id: str) -> io.BytesIO:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_path}:/content"
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            return self._download_file_alternative_with_drive(sharepoint_path, token, drive_id)
        return io.BytesIO(response.content)

    def _download_file_alternative(self, sharepoint_path: str, token: str) -> io.BytesIO:
        headers = {"Authorization": f"Bearer {token}"}
        encoded_path = sharepoint_path.replace('/', '%2F')
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive/root:{sharepoint_path}:/content"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Failed to download file from both methods. Status: {response.status_code}, Error: {response.text}")
        
        return io.BytesIO(response.content)

    def _download_file_alternative_with_drive(self, sharepoint_path: str, token: str, drive_id: str) -> io.BytesIO:
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sharepoint_path}:/content"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Failed to download file with custom drive. Status: {response.status_code}, Error: {response.text}")
        
        return io.BytesIO(response.content)

    def download_file_by_id(self, file_id: str) -> io.BytesIO:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{file_id}/content"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Failed to download file by ID: {response.text}")
        
        return io.BytesIO(response.content)

    def get_default_template(self) -> io.BytesIO:
        return self.download_file_by_path(self.template_path)

    def upload_new_file(self, file_stream: io.BytesIO, file_name: str, folder_path: str = "/Output") -> Dict[str, Any]:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        full_path = f"{folder_path}/{file_name}"
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:{full_path}:/content"
        
        file_stream.seek(0)
        response = requests.put(upload_url, headers=headers, data=file_stream.read())
        if response.status_code not in [200, 201]:
            raise Exception(f"Upload failed: {response.text}")
        
        result = response.json()
        return result

    def update_existing_file(self, file_id: str, file_stream: io.BytesIO) -> Dict[str, Any]:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{file_id}/content"
        
        file_stream.seek(0)
        response = requests.put(upload_url, headers=headers, data=file_stream.read())
        if response.status_code not in [200, 201]:
            raise Exception(f"Update failed: {response.text}")
        
        result = response.json()
        return result

    def get_template_by_id(self, template_id: str) -> io.BytesIO:
        if template_id == "DEFAULT" or template_id == "PROPOSAL_TEMPLATE":
            return self.get_default_template()
        else:
            try:
                return self.download_file_by_id(template_id)
            except Exception as e:
                return self.get_default_template()

    def search_files(self, search_query: str) -> Dict[str, Any]:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root/search(q='{search_query}')"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Search failed: {response.text}")
        
        return response.json()

    def list_folder_contents(self, folder_path: str = "/Templates") -> Dict[str, Any]:
        token = self.get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:{folder_path}:/children"
        
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Failed to list folder contents: {response.text}")
        
        return response.json()

    def generate_file_name(self, template_name: str = "Proposal") -> str:
        from datetime import datetime
        import time
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S") + f"_{int(time.time() * 1000) % 1000:03d}"
        return f"Generated_{template_name}_{timestamp}.docx"

    def get_document_by_name(self, document_name: str, is_old_document: bool, drive_id: Optional[str] = None) -> io.BytesIO:
        if is_old_document:
            current_drive_id = self.drive_id
            file_path = f"{self.output_path}/{document_name}"
            folder_type = "Output"
        else:
            current_drive_id = drive_id
            file_path = document_name
            folder_type = "Custom Drive"
            if not current_drive_id:
                raise ValueError(f"drive_id is required for new documents but was None for '{document_name}'. Please provide a valid driveId in the request payload.")

        try:
            document = self.download_file_by_path_with_drive(file_path, current_drive_id)
            return document
        except Exception as error:
            raise Exception(f"Document '{document_name}' not found in {folder_type}: {error}")

    def find_document_in_output(self, document_id: str) -> io.BytesIO:
        try:
            return self.download_file_by_id(document_id)
        except Exception as direct_error:
            try:
                folder_contents = self.list_folder_contents("/Output")
                
                for item in folder_contents.get("value", []):
                    if item.get("id") == document_id:
                        return self.download_file_by_id(document_id)
                
                raise Exception(f"Document ID '{document_id}' not found in Output directory")
            except Exception as search_error:
                raise Exception(f"Failed to find document '{document_id}' in Output directory: {search_error}")

    def update_existing_document_with_version(self, document_id: str, file_stream: io.BytesIO, increment_version: bool = True) -> Dict[str, Any]:
        import time
        import json
        
        max_retries = 3
        retry_delay = 2
        
        for attempt in range(max_retries):
            try:
                current_metadata = self.get_file_metadata(document_id)
                current_version = current_metadata.get("versionInfo", {}).get("majorVersion", 1)
                
                token = self.get_access_token()
                headers = {"Authorization": f"Bearer {token}"}
                upload_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{document_id}/content"
                
                file_stream.seek(0)
                response = requests.put(upload_url, headers=headers, data=file_stream.read())
                
                if response.status_code not in [200, 201]:
                    try:
                        error_data = response.json()
                        error_code = error_data.get("error", {}).get("code", "")
                        error_message = error_data.get("error", {}).get("message", "")
                        inner_error_code = error_data.get("error", {}).get("innerError", {}).get("code", "")
                        
                        if error_code == "notAllowed" and inner_error_code == "resourceLocked":
                            if attempt < max_retries - 1:
                                time.sleep(retry_delay)
                                retry_delay *= 2
                                continue
                            else:
                                raise Exception(f"Document is locked and cannot be updated after {max_retries} attempts. Please close the document in Word Online and try again. Error: {error_message}")
                        else:
                            raise Exception(f"Update failed: {response.text}")
                    except json.JSONDecodeError:
                        raise Exception(f"Update failed with status {response.status_code}: {response.text}")
                
                result = response.json()
                
                if increment_version:
                    new_version = current_version + 1
                    if "versionInfo" not in result:
                        result["versionInfo"] = {}
                    result["versionInfo"]["majorVersion"] = new_version
                
                return result
            except Exception as e:
                if "resourceLocked" in str(e) and attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    retry_delay *= 2
                    continue
                else:
                    raise

    def extract_metadata(self, sharepoint_response: Dict[str, Any]) -> Dict[str, Any]:
        version_info = sharepoint_response.get("versionInfo", {})
        
        version = 1
        if version_info and "majorVersion" in version_info:
            version = version_info["majorVersion"]
        elif "version" in sharepoint_response:
            version = sharepoint_response["version"]
        
        return {
            "fileId": sharepoint_response.get("id"),
            "fileName": sharepoint_response.get("name"),
            "webUrl": sharepoint_response.get("webUrl"),
            "version": version,
            "size": sharepoint_response.get("size", 0),
            "lastModified": sharepoint_response.get("lastModifiedDateTime"),
            "downloadUrl": sharepoint_response.get("@microsoft.graph.downloadUrl"),
            "driveId": sharepoint_response.get("parentReference", {}).get("driveId")
        }
