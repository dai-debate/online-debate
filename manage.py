import argparse
import yaml
import string
import secrets
from datetime import datetime
import math
import gspread
import gspread.utils as gsutils
from oauth2client.client import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from pyzoom import ZoomClient
from typing import List, Dict, Any
from worksheet import WorksheetEx


def generate_room(credentials: Credentials, file_id: str, sheet_index: int,
                  prefix: str, judge_num: int, api_key: Dict[str, str], settings: Dict[str, Any]):
    """試合会場を生成する

    :param credentials: Google の認証情報
    :type credentials: Credentials
    :param file_id: 管理用スプレッドシートのID
    :type file_id: str
    :param sheet_index: 対戦表シートのインデックス
    :type sheet_index: int
    :param prefix: 会場名のプレフィックス
    :type prefix: str
    :param judge_num: ジャッジの人数
    :type judge_num: int
    :param api_key: Zoom の APIキー/APIシークレット
    :type api_key: Dict[str, str]
    :param settings: Zoom ミーティングの設定情報
    :type settings: Dict[str, Any]
    """
    def get_users(client: ZoomClient) -> List[Dict[str, Any]]:
        response = client.raw.get_all_pages('/users', query={'status': 'active'})
        return response['users']

    def find_user(users: List[Dict[str, Any]], key: str, value: Any) -> List[Dict[str, Any]]:
        users = list(filter(lambda x: x[key] == value, users))
        return users

    def generate_password(length: int = 6):
        chars = string.digits
        return ''.join(secrets.choice(chars) for x in range(length))

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet = WorksheetEx.cast(book.get_worksheet(sheet_index))

    values = sheet.get_all_values()

    now = datetime.now()
    year = now.year
    month, day = values.pop(0)[0].split('/')

    client = ZoomClient(api_key['api-key'], api_key['api-secret'])

    users = get_users(client)

    meetings = []

    for value in values:
        matchName = value[0]
        hour_s, min_s = value[1].split(':')
        hour_e, min_e = value[2].split(':')
        start_time = datetime(year, int(month), int(day), int(hour_s), int(min_s))
        end_time = datetime(year, int(month), int(day), int(hour_e), int(min_e))
        duration = math.ceil((end_time - start_time).total_seconds()/60.0)
        userId = value[4+judge_num+1] if len(find_user(users, 'email', value[4+judge_num+1])) > 0 else None
        url = value[4+judge_num+3]
        meeting_id = value[4+judge_num+4]
        password = value[4+judge_num+5]

        if url and meeting_id and password:
            meetings.append([url, meeting_id, password])
        else:
            request = {
                'topic': prefix + matchName,
                'type': 2,
                'start_time': start_time.strftime('%Y-%m-%dT%H:%M:%S+09:00'),
                'duration': duration,
                'timezone': 'Asia/Tokyo',
                'password': generate_password(),
                'agenda': prefix + matchName,
                'settings': settings
            }
            response = client.raw.post(f'/users/{userId}/meetings', body=request)
            if response.ok:
                data = response.json()
                meetings.append([data['join_url'], data['id'], data['password']])
            else:
                meetings.append([None, None, None])

    start = gsutils.rowcol_to_a1(2, 4+judge_num+3)
    end = gsutils.rowcol_to_a1(1+len(values), 4+judge_num+5)
    sheet.batch_update([
        {'range': f'{start}:{end}', 'values': meetings}
    ])


def main():
    """メイン関数
    """
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('-c', '--config', type=str, default='config.yaml', help='Config file')
    parser.add_argument('-k', '--key', type=str, default='zoom-key.yaml', help='Zoom Config file')
    parser.add_argument('-s', '--settings', type=str, default='zoom-setting.yaml', help='Zoom Config file')
    parser.add_argument('command', type=str, choices=['generate-room'], help='Command')
    args = parser.parse_args()

    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
    ]

    with open(args.config, encoding='utf-8') as ifp1, open(args.key, encoding='utf-8') as ifp2, open(args.settings, encoding='utf-8') as ifp3:
        cfg = yaml.load(ifp1, Loader=yaml.SafeLoader)
        key = yaml.load(ifp2, Loader=yaml.SafeLoader)
        settings = yaml.load(ifp3, Loader=yaml.SafeLoader)
        credentials = ServiceAccountCredentials.from_json_keyfile_name(cfg['auth']['key_file'], scope)

        if args.command == 'generate-room':
            generate_room(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['prefix'], cfg['judge_num'], key, settings)


if __name__ == "__main__":
    main()
