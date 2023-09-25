import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class GoogleSheetsHandler:
    def __init__(self, scopes, spreadsheet_id):
        self.scopes = scopes
        self.spreadsheet_id = spreadsheet_id
        self.credentials = self.get_credentials()

    def get_credentials(self):
        credentials = None
        if os.path.exists("token.json"):
            credentials = Credentials.from_authorized_user_file("token.json", self.scopes)

        if not credentials or not credentials.valid:
            if credentials and credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("Lib\CredentialsGsheet.json", self.scopes)
                credentials = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(credentials.to_json())
        return credentials

    def fetch_data_from_spreadsheet(self, range_name):
        try:
            service = build("sheets", "v4", credentials=self.credentials)
            sheets = service.spreadsheets()
            result = sheets.values().get(spreadsheetId=self.spreadsheet_id, range=range_name).execute()
            values = result.get("values", [])
            return values
        except HttpError as error:
            print(error)

