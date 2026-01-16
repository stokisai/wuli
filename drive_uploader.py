import google.auth
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os
import pickle
import logging
import logging
import configparser
import time
import random
logger = logging.getLogger(__name__)

# 使用完整的drive权限，以便访问现有文件夹
SCOPES = ['https://www.googleapis.com/auth/drive']

class DriveUploader:
    def __init__(self, config_path="config.ini"):
        self.config = configparser.ConfigParser()
        self.config.read(config_path, encoding='utf-8')
        self.creds_path = self.config["Paths"]["CredentialsPath"]
        self.parent_id = self.config["Drive"].get("ParentFolderId", None)
        self.service = None

    def authenticate(self):
        """Authenticates with Google Drive API."""
        creds = None
        # The file token.pickle stores the user's access and refresh tokens.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                try:
                    creds = pickle.load(token)
                except Exception:
                    logger.warning("Token file seems corrupted, re-authenticating.")
        
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    logger.error(f"Error refreshing token: {e}")
                    creds = None
            
            if not creds:
                if not os.path.exists(self.creds_path):
                     logger.error(f"Credentials file not found at {self.creds_path}")
                     # raise Exception("Credentials not found") 
                     # Allow instantiation without auth for testing parts, but methods will fail
                     return False

                flow = InstalledAppFlow.from_client_secrets_file(
                    self.creds_path, SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        try:
            self.service = build('drive', 'v3', credentials=creds)
            return True
        except Exception as e:
            logger.error(f"Failed to build drive service: {e}")
            return False

    def create_folder(self, folder_name, parent_id=None):
        """Creates a folder or returns existing one."""
        if not self.service:
            logger.error("Drive service not initialized.")
            return None
        
        # 清理文件夹名称 - 替换反斜杠和斜杠为下划线，避免API查询错误
        folder_name = folder_name.replace("\\", "_").replace("/", "_")
        
        if parent_id is None:
            parent_id = self.parent_id

        # Check if folder exists
        query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
        if parent_id:
            query += f" and '{parent_id}' in parents"
        
        try:
            results = self.service.files().list(q=query, fields="files(id, name)").execute()
            items = results.get('files', [])
            if items:
                logger.info(f"Folder '{folder_name}' already exists: {items[0]['id']}")
                return items[0]['id']
            
            # Create folder
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            if parent_id:
                file_metadata['parents'] = [parent_id]
                
            file = self.service.files().create(body=file_metadata, fields='id').execute()
            logger.info(f"Created folder '{folder_name}': {file.get('id')}")
            return file.get('id')
        except Exception as e:
            logger.error(f"Error creating folder {folder_name}: {e}")
            return None

    def make_public(self, file_id):
        """Makes the file publicly readable."""
        if not self.service: 
            return False
        try:
            self.service.permissions().create(
                fileId=file_id,
                body={'role': 'reader', 'type': 'anyone'}
            ).execute()
            return True
        except Exception as e:
            logger.error(f"Error setting permissions for {file_id}: {e}")
            return False

    def upload_file(self, file_path, folder_id):
        """Uploads a file to the specified folder and makes it public."""
        if not self.service:
            logger.error("Drive service not initialized.")
            return None
            
        file_name = os.path.basename(file_path)
        
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, resumable=True)
        
        # RETRY LOGIC (3 Attempts)
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # If we are retrying, media might need to be reset or recreated if stream was consumed?
                # MediaFileUpload with filename usually handles reopen, but safer to recreate if failed.
                if attempt > 0:
                    logger.info(f"Retrying upload attempt {attempt+1}/{max_retries}...")
                    media = MediaFileUpload(file_path, resumable=True)
                    time.sleep(2 * attempt) # Exponential backoff: 0s, 2s, 4s

                file = self.service.files().create(body=file_metadata,
                                                   media_body=media,
                                                   fields='id, webViewLink').execute()
                file_id = file.get('id')
                logger.info(f"Uploaded file '{file_name}': {file_id}")
                
                # Set permissions (Separate retry potentially needed, but usually fast)
                if self.make_public(file_id):
                    logger.info(f"Made '{file_name}' public.")
                else:
                    logger.warning(f"Could not make '{file_name}' public.")
                    
                return file # Success
            
            except Exception as e:
                logger.warning(f"Upload attempt {attempt+1} failed for {file_name}: {e}")
                if attempt == max_retries - 1:
                    logger.error(f"Final upload failure for {file_name} after {max_retries} attempts.")
                    return None
        return None

    def get_direct_link(self, file_id):
        """Converts file ID to direct link format."""
        return f"https://drive.google.com/uc?export=view&id={file_id}"
