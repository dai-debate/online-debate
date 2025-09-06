import os
import sys
from pathlib import Path

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request as OAuthRequest
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError


DEFAULET_SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
]


def get_credentials(
        path_client_secret: Path, path_token: Path = Path('google-auth-token.json'),
        scopes: list[str] = DEFAULET_SCOPES):

    creds = None
    try:
        if path_token and path_token.exists():
            creds = Credentials.from_authorized_user_file(path_token, scopes)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(OAuthRequest())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(path_client_secret, scopes)
                creds = flow.run_local_server(port=0)

            with path_token.open('w', encoding='utf-8', ) as ofp:
                ofp.write(creds.to_json())
            os.chmod(path_token, 0o600)

    except HttpError as err:
        print(f'An error occurred: {err}', file=sys.stderr)

    return creds
