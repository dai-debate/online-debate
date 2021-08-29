import argparse
import yaml
import string
import secrets
from datetime import datetime
import math
import time
import re
import sys
import gspread
import gspread.utils as gsutils
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.client import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from pyzoom import ZoomClient
from pyzoom import err as zoom_error
from typing import List, Dict, Any
from worksheet import WorksheetEx


def generate_room(credentials: Credentials, file_id: str, sheet_index: int,
                  prefix: str, judge_num: int, staff_num: int, api_key: Dict[str, str], settings: Dict[str, Any], **kwargs):
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
    :param staff_num: スタッフの人数
    :type staff_num: int
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

    offset = kwargs['offset'] if 'offset' in kwargs else 0
    limit = kwargs['limit'] if 'limit' in kwargs else sys.maxsize

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet = WorksheetEx.cast(book.get_worksheet(sheet_index))

    values = sheet.get_all_values()

    year, month, day = values.pop(0)[1].split('/')
    values.pop(0)

    client = ZoomClient(api_key['api-key'], api_key['api-secret'])

    users = get_users(client)

    meetings = []

    for i, value in enumerate(values):

        if i < offset:
            continue

        if i >= limit:
            break

        matchName = value[0]
        hour_s, min_s = value[2].split(':')
        hour_e, min_e = value[3].split(':')
        start_time = datetime(int(year), int(month), int(day), int(hour_s), int(min_s))
        end_time = datetime(int(year), int(month), int(day), int(hour_e), int(min_e))
        duration = math.ceil((end_time - start_time).total_seconds()/60.0)
        userId = value[5+judge_num+staff_num+1] if len(find_user(users, 'email', value[5+judge_num+staff_num+1])) > 0 else None
        url = value[5+judge_num+staff_num+3]
        meeting_id = value[5+judge_num+staff_num+4]
        password = value[5+judge_num+staff_num+5]

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
                meetings.append([data['join_url'], f"'{data['id']}", f"'{data['password']}"])
            else:
                meetings.append([None, None, None])

        print(prefix + matchName)

    start = gsutils.rowcol_to_a1(3+offset, 6+judge_num+staff_num+3)
    end = gsutils.rowcol_to_a1(2+len(meetings)+offset, 6+judge_num+staff_num+5)
    sheet.batch_update([
        {'range': f'{start}:{end}', 'values': meetings}
    ], value_input_option='USER_ENTERED')


def clear_room(credentials: Credentials, file_id: str, sheet_index: int,
               judge_num: int, staff_num: int, api_key: Dict[str, str], **kwargs):
    """試合会場を生成する

    :param credentials: Google の認証情報
    :type credentials: Credentials
    :param file_id: 管理用スプレッドシートのID
    :type file_id: str
    :param sheet_index: 対戦表シートのインデックス
    :type sheet_index: int
    :param judge_num: ジャッジの人数
    :type judge_num: int
    :param staff_num: スタッフの人数
    :type staff_num: int
    :param api_key: Zoom の APIキー/APIシークレット
    :type api_key: Dict[str, str]
    """
    def delete_meetings(client: ZoomClient, ids: List[str], offset: int, limit: int):
        count = 0
        for i, id in enumerate(ids):

            if i < offset:
                continue

            if i >= limit:
                break

            if id:
                try:
                    response = client.raw.get_all_pages(f'/meetings/{id}')
                except zoom_error.NotFound:
                    continue
                title = response['topic']
                response = client.raw.delete(f'/meetings/{id}')
                print(title)

            count += 1

        return count

    offset = kwargs['offset'] if 'offset' in kwargs else 0
    limit = kwargs['limit'] if 'limit' in kwargs else sys.maxsize

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet = WorksheetEx.cast(book.get_worksheet(sheet_index))

    values = sheet.get_all_values()
    values = values[2:]
    ids = [v[6+judge_num+staff_num+2+1] for v in values]

    client = ZoomClient(api_key['api-key'], api_key['api-secret'])

    count = delete_meetings(client, ids, offset, limit)

    update_values = [['']*3 for i in range(count)]
    start = gsutils.rowcol_to_a1(3+offset, 6+judge_num+staff_num+3)
    end = gsutils.rowcol_to_a1(2+len(update_values)+offset, 6+judge_num+staff_num+5)
    sheet.batch_update([
        {'range': f'{start}:{end}', 'values': update_values}
    ], value_input_option='USER_ENTERED')

    pass


def generate_ballot(credentials: Credentials, file_id: str, sheet_index_matches: int, sheet_index_vote: int,
                    judge_num: int, ballot_config: Dict[str, Any], **kwargs):
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

    offset = kwargs['offset'] if 'offset' in kwargs else 0
    limit = kwargs['limit'] if 'limit' in kwargs else sys.maxsize

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet_matches = WorksheetEx.cast(book.get_worksheet(sheet_index_matches))

    values = sheet_matches.get_all_values()
    values = values[2:]

    sheet_vote = WorksheetEx.cast(book.get_worksheet(sheet_index_vote))
    row_count = len(sheet_vote.col_values(1))
    if offset <= 0:
        if row_count > 1:
            sheet_vote.delete_rows(2, row_count)
        sheet_vote = WorksheetEx.cast(book.get_worksheet(sheet_index_vote))

    votes = []
    new_ballots = []
    actual_judge_num = 0

    for i, value in enumerate(values):

        if i < offset:
            continue

        if i >= limit:
            break

        if not value[4] or not value[5]:
            new_ballots.append([None]*judge_num)
            continue

        ballots = []
        for j in range(judge_num):

            if not value[6+j]:
                continue

            actual_judge_num += 1

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

            row = row_count + actual_judge_num
            vote = [''] * 11
            vote[0] = f"'{value[0]}"
            vote[1] = j
            vote[2] = value[6+j]
            vote[4] = f'=IF({gsutils.rowcol_to_a1(row,10)}="肯定",1,0)'
            vote[7] = f'=IF({gsutils.rowcol_to_a1(row,10)}="否定",1,0)'
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

    start = gsutils.rowcol_to_a1(3+offset, 7)
    end = gsutils.rowcol_to_a1(2+len(new_ballots)+offset, 6+judge_num)
    target_range = f'{start}:{end}' if judge_num > 1 or len(new_ballots) > 2 else f'{start}'
    judges = sheet_matches.get(target_range)

    if(len(judges)) > 0:
        new_values = [[f'=HYPERLINK("{col}","{judges[i][j]}")' if col else f'{judges[i][j]}' for j, col in enumerate(row)] for i, row in enumerate(new_ballots)]
        sheet_matches.batch_update([
            {'range': f'{start}:{end}', 'values': new_values}
        ], value_input_option='USER_ENTERED')

    pass


def generate_member_list(credentials: Credentials, file_id: str, sheet_index_matches: int, member_list_config: Dict[str, Any], **kwargs):
    """対戦表に基づき、出場メンバー届を生成する

    :param credentials: Google の認証情報
    :type credentials: Credentials
    :param file_id: 管理用スプレッドシートのID
    :type file_id: str
    :param sheet_index_matches: 対戦表シートのインデックス
    :type sheet_index_matches: int
    :param member_list_config: 勝敗・ポイント記入シートの参照関係設定
    :type member_list_config: Dict[str, Any]
    """

    offset = kwargs['offset'] if 'offset' in kwargs else 0
    limit = kwargs['limit'] if 'limit' in kwargs else sys.maxsize

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet_matches = WorksheetEx.cast(book.get_worksheet(sheet_index_matches))

    values = sheet_matches.get_all_values()
    values = values[2:]

    new_member_lists = []

    for i, value in enumerate(values):

        if i < offset:
            continue

        if i >= limit:
            break

        member_lists = []
        for j in range(2):

            if not value[4+j]:
                member_lists.append(None)
                continue

            side = '肯定' if j == 0 else '否定'

            new_book = gc.copy(member_list_config['template'])
            member_lists.append(new_book.url)
            new_sheet = WorksheetEx.cast(new_book.get_worksheet(0))

            new_book.batch_update({
                'requests': [
                    {
                        'updateSpreadsheetProperties': {
                            'properties': {
                                'title': f"{member_list_config['title']} {value[0]} {side}"
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

            gfile['parents'] = [{'id': member_list_config['folder']}]
            gfile.Upload()

            for link in member_list_config['to_list']:
                if type(link) == list:
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
                elif type(link) == dict:
                    if 'side' in link:
                        new_sheet.update_acell(link['side'], side)
                time.sleep(1)

            print(f"{member_list_config['title']} {value[0]} {side}")

        new_member_lists.append(member_lists)

    start = gsutils.rowcol_to_a1(3+offset, 5)
    end = gsutils.rowcol_to_a1(2+len(new_member_lists)+offset, 6)
    target_range = f'{start}:{end}'
    lists = sheet_matches.get(target_range)

    new_values = [[f'=HYPERLINK("{col}","{lists[i][j]}")' if col else f'{lists[i][j]}' for j, col in enumerate(row)] for i, row in enumerate(new_member_lists)]
    sheet_matches.batch_update([
        {'range': f'{start}:{end}', 'values': new_values}
    ], value_input_option='USER_ENTERED')

    pass


def generate_aggregate(credentials: Credentials, file_id: str, sheet_index_matches: int,
                       judge_num: int, staff_num: int, aggregate_config: Dict[str, Any], **kwargs):
    """対戦表に基づき、集計用紙を生成する

    :param credentials: Google の認証情報
    :type credentials: Credentials
    :param file_id: 管理用スプレッドシートのID
    :type file_id: str
    :param sheet_index_matches: 対戦表シートのインデックス
    :type sheet_index_matches: int
    :param judge_num: ジャッジの人数
    :type judge_num: int
    :param staff_num: スタッフの人数
    :type staff_num: int
    :param aggregate_config: 勝敗・ポイント記入シートの参照関係設定
    :type aggregate_config: Dict[str, Any]
    """

    offset = kwargs['offset'] if 'offset' in kwargs else 0
    limit = kwargs['limit'] if 'limit' in kwargs else sys.maxsize

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet_matches = WorksheetEx.cast(book.get_worksheet(sheet_index_matches))

    values = sheet_matches.get_all_values(value_render_option='FORMULA')
    values = values[2:]

    new_aggregates = []
    pattern = r'=HYPERLINK\("https://docs\.google\.com/spreadsheets/d/(.*?)","(.*?)"\)'

    for i, value in enumerate(values):

        if i < offset:
            continue

        if i >= limit:
            break

        if not values[4] or not values[5]:
            new_aggregates.append(None)
            continue

        new_book = gc.copy(aggregate_config['template'])
        new_aggregates.append(new_book.url)
        new_sheet = WorksheetEx.cast(new_book.get_worksheet(0))

        new_book.batch_update({
            'requests': [
                {
                    'updateSpreadsheetProperties': {
                        'properties': {
                            'title': f"{aggregate_config['title']} {value[0]}"
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

        gfile['parents'] = [{'id': aggregate_config['folder']}]
        gfile.Upload()

        for link in aggregate_config['to_aggregate']:
            if type(link) == list:
                if type(link[0]) == int:
                    new_sheet.update_acell(link[1], value[link[0]])
                elif type(link[0]) == str:
                    cell = sheet_matches.acell(link[0], value_render_option='FORMATTED_VALUE')
                    new_sheet.update_acell(link[1], cell.value)
                elif type(link[0]) == list:
                    options = [value[x] for x in link[0]]
                    new_sheet.set_data_validation(link[1], WorksheetEx.conditiontype.ONE_OF_LIST, options, strict=True, custom_ui=True)
            time.sleep(1)

        links = [''] * len(aggregate_config['link'])
        for j in range(judge_num):

            match = re.match(pattern, value[6+j])
            if not match:
                continue

            ballot = match.group(1)

            for k, link in enumerate(aggregate_config['link']):
                if link[2] == 'POINT':
                    links[k] = links[k] + f'{"=" if j==0 else "+"}IMPORTRANGE("{ballot}","{link[0]}")'
                elif link[2] == 'VOTE_AFF':
                    links[k] = links[k] + f'{"=" if j==0 else "+"}IF(IMPORTRANGE("{ballot}","{link[0]}")="肯定",1,0)'
                elif link[2] == 'VOTE_NEG':
                    links[k] = links[k] + f'{"=" if j==0 else "+"}IF(IMPORTRANGE("{ballot}","{link[0]}")="否定",1,0)'

        for k, link in enumerate(aggregate_config['link']):
            new_sheet.update_acell(link[1], links[k])
            time.sleep(1)

        print(f"{aggregate_config['title']} {value[0]}")

    start = gsutils.rowcol_to_a1(3+offset, 6+judge_num+staff_num+9)
    end = gsutils.rowcol_to_a1(2+len(new_aggregates)+offset, 6+judge_num+staff_num+9)
    sheet_matches.batch_update([
        {'range': f'{start}:{end}', 'values': [[f'=HYPERLINK("{v}","Link")' if v else ''] for v in new_aggregates]}
    ], value_input_option='USER_ENTERED')

    pass


def generate_advice(credentials: Credentials, file_id: str, sheet_index_matches: int,
                    judge_num: int, staff_num: int, advice_config: Dict[str, Any], **kwargs):
    """対戦表に基づき、アドバイスシートを生成する

    :param credentials: Google の認証情報
    :type credentials: Credentials
    :param file_id: 管理用スプレッドシートのID
    :type file_id: str
    :param sheet_index_matches: 対戦表シートのインデックス
    :type sheet_index_matches: int
    :param judge_num: ジャッジの人数
    :type judge_num: int
    :param staff_num: スタッフの人数
    :type staff_num: int
    :param advice_config: 勝敗・ポイント記入シートの参照関係設定
    :type advice_config: Dict[str, Any]
    """

    offset = kwargs['offset'] if 'offset' in kwargs else 0
    limit = kwargs['limit'] if 'limit' in kwargs else sys.maxsize

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet_matches = WorksheetEx.cast(book.get_worksheet(sheet_index_matches))

    values = sheet_matches.get_all_values()
    values = values[2:]

    new_advice = []

    for i, value in enumerate(values):

        if i < offset:
            continue

        if i >= limit:
            break

        advice_list = []

        for j in range(2):

            if not value[4+j]:
                advice_list.append(None)
                continue

            side = '肯定' if j == 0 else '否定'

            new_book = gc.copy(advice_config['template'])
            advice_list.append(new_book.url)
            new_sheet = WorksheetEx.cast(new_book.get_worksheet(0))

            new_book.batch_update({
                'requests': [
                    {
                        'updateSpreadsheetProperties': {
                            'properties': {
                                'title': f"{advice_config['title']} {value[0]} {side}"
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

            gfile['parents'] = [{'id': advice_config['folder']}]
            gfile.Upload()

            for link in advice_config['to_advice']:
                if type(link) == list:
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
                elif type(link) == dict:
                    if 'aff' in link and side == '肯定':
                        for x in link['aff']:
                            if type(x) == int:
                                new_sheet.update_acell(link[1], value[link[0]])
                            elif type(x) == str:
                                new_sheet.update_acell(x, side)
                            elif type(x) == list:
                                new_sheet.update_acell(x[1], value[x[0]])
                    if 'neg' in link and side == '否定':
                        for x in link['neg']:
                            if type(x) == int:
                                new_sheet.update_acell(link[1], value[link[0]])
                            elif type(x) == str:
                                new_sheet.update_acell(x, side)
                            elif type(x) == list:
                                new_sheet.update_acell(x[1], value[x[0]])
                time.sleep(1)

            print(f"{advice_config['title']} {value[0]} {side}")

        new_advice.append(advice_list)

    start = gsutils.rowcol_to_a1(3+offset, 6+judge_num+staff_num+10)
    end = gsutils.rowcol_to_a1(2+len(new_advice)+offset, 6+judge_num+staff_num+11)

    new_values = [[f'=HYPERLINK("{col}","Link")' if col else '' for col in row] for row in new_advice]
    sheet_matches.batch_update([
        {'range': f'{start}:{end}', 'values': new_values}
    ], value_input_option='USER_ENTERED')

    pass


def update_live(credentials: Credentials, file_id: str, sheet_index_matches: int,
                judge_num: int, staff_num: int, api_key: Dict[str, str], **kwargs):
    """Zoomミーティングとライブストリーミングの関連付けを行う

    :param credentials: Google の認証情報
    :type credentials: Credentials
    :param file_id: 管理用スプレッドシートのID
    :type file_id: str
    :param sheet_index_matches: 対戦表シートのインデックス
    :type sheet_index_matches: int
    :param judge_num: ジャッジの人数
    :type judge_num: int
    :param staff_num: スタッフの人数
    :type staff_num: int
    :param api_key: Zoom の APIキー/APIシークレット
    :type api_key: Dict[str, str]
    :raises RuntimeError: 関連付けに失敗した場合に例外を送出
    """

    offset = kwargs['offset'] if 'offset' in kwargs else 0
    limit = kwargs['limit'] if 'limit' in kwargs else sys.maxsize

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet_matches = WorksheetEx.cast(book.get_worksheet(sheet_index_matches))

    values = sheet_matches.get_all_values()
    values = values[2:]

    client = ZoomClient(api_key['api-key'], api_key['api-secret'])

    for i, value in enumerate(values):

        if i < offset:
            continue

        if i >= limit:
            break

        meeting_id = value[5+judge_num+staff_num+4]
        stream_url = value[5+judge_num+staff_num+6]
        stream_key = value[5+judge_num+staff_num+7]
        page_url = value[5+judge_num+staff_num+8]

        if meeting_id and stream_url and stream_key and page_url:
            response = client.raw.patch(f'/meetings/{meeting_id}/livestream', body={
                'stream_url': stream_url,
                'stream_key': stream_key,
                'page_url': page_url,
            })
            if not response.ok:
                raise RuntimeError(f'Update live failed: {meeting_id}')
        print(value[0])
    pass


def update_ballot(credentials: Credentials, file_id: str, sheet_index_matches: int, sheet_index_vote: int,
                  judge_num: int, ballot_config: Dict[str, Any], **kwargs):
    """対戦表の変更点を勝敗・ポイント記入シートに反映する

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

    offset = kwargs['offset'] if 'offset' in kwargs else 0
    limit = kwargs['limit'] if 'limit' in kwargs else sys.maxsize

    gc = gspread.authorize(credentials)

    book = gc.open_by_key(file_id)
    sheet_matches = WorksheetEx.cast(book.get_worksheet(sheet_index_matches))
    sheet_vote = WorksheetEx.cast(book.get_worksheet(sheet_index_vote))

    values = sheet_matches.get_all_values(value_render_option='FORMULA')
    values = values[2:]

    pattern = r'=HYPERLINK\("https://docs\.google\.com/spreadsheets/d/(.*?)","(.*?)"\)'

    for i, value in enumerate(values):

        if i < offset:
            continue

        if i >= limit:
            break

        if not value[4] or not value[5]:
            continue

        match = re.match(pattern, value[4])
        if match:
            value[4] = match.group(2)

        match = re.match(pattern, value[5])
        if match:
            value[5] = match.group(2)

        for j in range(judge_num):

            if not value[6+j]:
                continue

            match = re.match(pattern, value[6+j])
            if not match:

                new_book = gc.copy(ballot_config['template'])
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
                vote = [''] * 11
                vote[0] = f"'{value[0]}"
                vote[1] = j
                vote[2] = value[6+j]
                vote[4] = f'=IF({gsutils.rowcol_to_a1(row,10)}="肯定",1,0)'
                vote[7] = f'=IF({gsutils.rowcol_to_a1(row,10)}="否定",1,0)'
                for link in ballot_config['to_vote']:
                    vote[link[1]] = f'=IMPORTRANGE("{new_book.id}","{new_sheet.title}!{link[0]}")'
                    pass

                start = gsutils.rowcol_to_a1(row, 1)
                end = gsutils.rowcol_to_a1(row, 11)
                sheet_vote.batch_update([
                    {'range': f'{start}:{end}', 'values': [vote]}
                ], value_input_option='USER_ENTERED')

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

                sheet_matches.update_cell(3+i, 7+j, f'=HYPERLINK("{new_book.url}","{value[6+j]}")')

                print(f"{ballot_config['title']} {value[0]} #{j}")

    pass


def main():
    """メイン関数
    """
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('-c', '--config', type=str, default='config.yaml', help='Config file')
    parser.add_argument('-k', '--key', type=str, default='zoom-key.yaml', help='Zoom Config file')
    parser.add_argument('-s', '--settings', type=str, default='zoom-setting.yaml', help='Zoom Config file')
    parser.add_argument('command', type=str, choices=[
        'generate-room',
        'clear-room',
        'generate-ballot',
        'generate-member-list',
        'generate-aggregate',
        'generate-advice',
        'update-live',
        'update-ballot',
    ], help='Command')
    parser.add_argument('-o', '--offset', type=int, default=0)
    parser.add_argument('-l', '--limit', type=int, default=sys.maxsize)
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
            generate_room(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['prefix'], cfg['judge_num'], cfg['staff_num'], key, settings, offset=args.offset, limit=args.limit)
        elif args.command == 'clear-room':
            clear_room(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['judge_num'], cfg['staff_num'], key, args.offset, limit=args.limit)
        elif args.command == 'generate-ballot':
            generate_ballot(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['sheets']['vote'], cfg['judge_num'], cfg['ballot'], offset=args.offset, limit=args.limit)
        elif args.command == 'generate-member-list':
            generate_member_list(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['member_list'], offset=args.offset, limit=args.limit)
        elif args.command == 'generate-aggregate':
            generate_aggregate(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['judge_num'], cfg['staff_num'], cfg['aggregate'], offset=args.offset, limit=args.limit)
        elif args.command == 'generate-advice':
            generate_advice(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['judge_num'], cfg['staff_num'], cfg['advice'], offset=args.offset, limit=args.limit)
        elif args.command == 'update-live':
            update_live(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['judge_num'], cfg['staff_num'], key, offset=args.offset, limit=args.limit)
        elif args.command == 'update-ballot':
            update_ballot(credentials, cfg['file_id'], cfg['sheets']['matches'], cfg['sheets']['vote'], cfg['judge_num'], cfg['ballot'], offset=args.offset, limit=args.limit)

        print('Complete.')


if __name__ == "__main__":
    main()
