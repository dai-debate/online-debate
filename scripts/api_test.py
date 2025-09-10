import logging
import argparse
from pathlib import Path

from googleapiclient.discovery import build as build_google_service
from googleapiclient.errors import HttpError

from google_api import get_credentials, SpreadSheet


logger = logging.getLogger()
logger.setLevel(logging.DEBUG)


class AppNamespace(argparse.Namespace):
    path_client_secrets: Path
    path_token: Path


def main():
    parser = argparse.ArgumentParser(description='Test Google Drive and Sheets API access.')
    parser.add_argument(
        '--client', type=Path, dest='path_client_secrets',
        default='dkoshien-app.json', help='Path to client secrets file')
    parser.add_argument(
        '--token', type=Path, dest='path_token',
        default='google-auth-token.json', help='Path to OAuth token file')
    args = parser.parse_args(namespace=AppNamespace())

    creds = get_credentials(args.path_client_secrets, args.path_token, scopes=[
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets',
    ])

    try:
        sheet_id = '1bQvY_UblM8LK7JppLRDmVZGzTs9p68F9ASI-Lp94N0w'
        spreadsheet = SpreadSheet(creds, sheet_id)

        drive_service = build_google_service('drive', 'v3', credentials=creds)
        sheets_service = build_google_service('sheets', 'v4', credentials=creds)

        # Test Drive API
        drive_results = drive_service.files().list(pageSize=10).execute()
        items = drive_results.get('files', [])
        print('Drive Files:')
        for item in items:
            print(f"{item['name']} ({item['id']})")

        # Test Sheets API
        spreadsheet = {
            'properties': {
                'title': 'API Test Spreadsheet'
            }
        }
        spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
        print(f'Spreadsheet ID: {spreadsheet.get("spreadsheetId")}')

    except HttpError as error:
        print(f'An error occurred: {error}')


if __name__ == '__main__':
    main()
