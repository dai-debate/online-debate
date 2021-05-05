import argparse
import yaml
import string
import secrets
from datetime import datetime
import math
import time
import gspread
import gspread.utils as gsutils
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
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
        return f"'{''.join(secrets.choice(chars) for x in range(length))}"

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet = WorksheetEx.cast(book.get_worksheet(sheet_index))

    values = sheet.get_all_values()

    year, month, day = values.pop(0)[1].split('/')
    values.pop(0)

    client = ZoomClient(api_key['api-key'], api_key['api-secret'])

    users = get_users(client)

    meetings = []

    for value in values:
        matchName = value[0]
        hour_s, min_s = value[2].split(':')
        hour_e, min_e = value[3].split(':')
        start_time = datetime(year, int(month), int(day), int(hour_s), int(min_s))
        end_time = datetime(year, int(month), int(day), int(hour_e), int(min_e))
        duration = math.ceil((end_time - start_time).total_seconds()/60.0)
        userId = value[5+judge_num+1] if len(find_user(users, 'email', value[5+judge_num+1])) > 0 else None
        url = value[5+judge_num+3]
        meeting_id = value[5+judge_num+4]
        password = value[5+judge_num+5]

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

    start = gsutils.rowcol_to_a1(3, 6+judge_num+3)
    end = gsutils.rowcol_to_a1(2+len(values), 6+judge_num+5)
    sheet.batch_update([
        {'range': f'{start}:{end}', 'values': meetings}
    ], value_input_option='USER_ENTERED')


def generate_ballot(credentials: Credentials, file_id: str, sheet_index_matches: int, sheet_index_vote: int,
                    judge_num: int, ballot_config: Dict[str, Any]):
    """対戦表に基づき、勝敗・ポイント記入シートを生成する

    :param credentials: Google の認証情報
    :type credentials: Credentials
    :param file_id: 管理用スプレッドシートのID
    :type file_id: str
    :param sheet_index_matches: 対戦表シートのインデックス
    :type sheet_index_matches: int
    :param sheet_index_vote: 投票シートのインデックス
    :type sheet_index_vote: int
    :param judge_num: ジャッジの人数
    :type judge_num: int
    :param ballot_config: 勝敗・ポイント記入シートの参照関係設定
    :type ballot_config: Dict[str, Any]
    """
    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet_matches = WorksheetEx.cast(book.get_worksheet(sheet_index_matches))

    values = sheet_matches.get_all_values()

    values.pop(0)
    values.pop(0)

    sheet_vote = WorksheetEx.cast(book.get_worksheet(sheet_index_vote))
    row_count = len(sheet_vote.col_values(1))
    if row_count > 1:
        sheet_vote.delete_rows(2, row_count)
    sheet_vote = WorksheetEx.cast(book.get_worksheet(sheet_index_vote))

    votes = []
    new_ballots = []

    for i, value in enumerate(values):
        ballots = []
        for j in range(judge_num):

            if not value[6*j]:
                continue

            new_book = gc.copy(ballot_config['template'])
            ballots.append(new_book.url)
            new_sheet = WorksheetEx.cast(new_book.get_worksheet(0))

            new_book.batch_update({
                'requests': [
                    {
                        'updateSpreadsheetProperties': {
                            'properties': {
                                'title': f"{ballot_config['title']} {value[0]} #{j}"
                            },
                            'fields': 'title'
                        }
                    }
                ]
            })

            gauth = GoogleAuth()
            gauth.credentials = credentials

            gdrive = GoogleDrive(gauth)
            gfile = gdrive.CreateFile({'id': new_book.id})
            gfile.FetchMetadata(fetch_all=True)

            gfile['parents'] = [{'id': ballot_config['folder']}]
            gfile.Upload()

            row = 2 + j + judge_num * i
            vote = [''] * 12
            vote[1] = j
            vote[2] = value[6+j]
            vote[4] = f'=IF({gsutils.rowcol_to_a1(row,10)}={gsutils.rowcol_to_a1(row,4)},1,0)'
            vote[7] = f'=IF({gsutils.rowcol_to_a1(row,10)}={gsutils.rowcol_to_a1(row,7)},1,0)'
            for link in ballot_config['to_vote']:
                vote[link[1]] = f'=IMPORTRANGE("{new_book.id}","{new_sheet.title}!{link[0]}")'
                pass

            votes.append(vote)

            for link in ballot_config['to_ballot']:
                if type(link[0]) == int:
                    if len(link) >= 3 and link[2]:
                        new_sheet.update_acell(link[1], value[link[0]+j])
                    else:
                        new_sheet.update_acell(link[1], value[link[0]])
                elif type(link[0]) == str:
                    cell = sheet_matches.acell(link[0], value_render_option='FORMATTED_VALUE')
                    new_sheet.update_acell(link[1], cell.value)
                elif type(link[0]) == list:
                    options = [value[x] for x in link[0]]
                    new_sheet.set_data_validation(link[1], WorksheetEx.conditiontype.ONE_OF_LIST, options, strict=True, custom_ui=True)
                time.sleep(1)

            print(f"{ballot_config['title']} {value[0]} #{j}")

        new_ballots.append(ballots)

    sheet_vote.append_rows(votes, value_input_option='USER_ENTERED', insert_data_option='INSERT_ROWS')

    start = gsutils.rowcol_to_a1(3, 7)
    end = gsutils.rowcol_to_a1(2+len(new_ballots), 6+judge_num)
    target_range = f'{start}:{end}' if judge_num > 1 or len(new_ballots) > 2 else f'{start}'
    judges = sheet_matches.get(target_range)

    if(len(judges)) > 0:
        new_values = [[f'=HYPERLINK("{col}","{judges[i][j]}")' for j, col in enumerate(row)] for i, row in enumerate(new_ballots)]
        sheet_matches.batch_update([
            {'range': f'{start}:{end}', 'values': new_values}
        ], value_input_option='USER_ENTERED')

    pass


def main():
    """メイン関数
    """
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('-c', '--config', type=str, default='config.yaml', help='Config file')
    parser.add_argument('-k', '--key', type=str, default='zoom-key.yaml', help='Zoom Config file')
    parser.add_argument('-s', '--settings', type=str, default='zoom-setting.yaml', help='Zoom Config file')
    parser.add_argument('command', type=str, choices=['generate-room', 'generate-ballot'], help='Command')
    args = parser.parse_args()

    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    with open(args.config, encoding='utf-8') as ifp1, open(args.key, encoding='utf-8') as ifp2, open(args.settings, encoding='utf-8') as ifp3:
        cfg = yaml.load(ifp1, Loader=yaml.SafeLoader)
        key = yaml.load(ifp2, Loader=yaml.SafeLoader)
        settings = yaml.load(ifp3, Loader=yaml.SafeLoader)
        credentials = ServiceAccountCredentials.from_json_keyfile_name(cfg['auth']['key_file'], scope)

        if args.command == 'generate-room':
            generate_room(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['prefix'], cfg['judge_num'], key, settings)
        elif args.command == 'generate-ballot':
            generate_ballot(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['sheets']['vote'], cfg['judge_num'], cfg['ballot'])

        print('Complete.')


if __name__ == "__main__":
    main()
