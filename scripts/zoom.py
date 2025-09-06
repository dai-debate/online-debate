from typing import List, Dict, Any

import requests
import base64


class Zoom:

    BASE_URL = 'https://zoom.us'
    API_URL = 'https://api.zoom.us/v2'

    def __init__(
            self,
            client_id: str, client_secret: str,
            account_id: str):

        auth = base64.b64encode(f'{client_id}:{client_secret}'.encode()).decode()
        url = f'{Zoom.BASE_URL}/oauth/token'
        params = {
            'grant_type': 'account_credentials',
            'account_id': account_id,
        }
        header = {
            'Authorization': f'Basic {auth}'
        }
        response = requests.post(url, params=params, headers=header)
        if response.ok:
            js = response.json()
            self.token = js['access_token']
        else:
            response.raise_for_status()

    def get_users(self, **kwargs) -> List[Dict[str, Any]]:
        users = []
        url = f'{Zoom.API_URL}/users'
        header = {
            'Authorization': f'Bearer {self.token}'
        }
        params = {}

        if 'status' in kwargs:
            params['status'] = kwargs['status']

        page_count = None
        page_number = None

        while page_count is None or page_number < page_count:
            if page_count is not None and page_number is not None:
                page_number += 1
                params['page_number'] = page_number

            if len(params) > 0:
                response = requests.get(url, params=params, headers=header)
            else:
                response = requests.get(url, headers=header)

            if response.ok:
                js = response.json()
                page_count = js['page_count']
                page_number = js['page_number']
                users.extend(js['users'])
            else:
                response.raise_for_status()

        return users

    def get_meeting(self, id: str) -> List[Dict[str, Any]]:
        url = f'{Zoom.API_URL}/meetings/{id}'
        header = {
            'Authorization': f'Bearer {self.token}'
        }
        response = requests.get(url, headers=header)
        if response.ok:
            return response.json()
        else:
            response.raise_for_status()

    def create_meeting(self, user_id: str, body: Dict[str, Any]) -> requests.Response:
        url = f'{Zoom.API_URL}/users/{user_id}/meetings'
        header = {
            'Authorization': f'Bearer {self.token}',
        }
        response = requests.post(url, headers=header, json=body)
        if response.ok:
            return response
        else:
            response.raise_for_status()

    def delete_meeting(self, id: str) -> bool:
        url = f'{Zoom.API_URL}/meetings/{id}'
        header = {
            'Authorization': f'Bearer {self.token}'
        }
        response = requests.delete(url, headers=header)
        if response.ok:
            return True
        else:
            response.raise_for_status()

    def update_livestream(
            self, meeting_id: str,
            stream_url: str, stream_key: str, page_url: str) -> bool:

        url = f'{Zoom.API_URL}/meetings/{meeting_id}/livestream'
        header = {
            'Authorization': f'Bearer {self.token}'
        }
        body = {
            'stream_url': stream_url,
            'stream_key': stream_key,
            'page_url': page_url,
        }
        response = requests.patch(url, headers=header, json=body)
        if response.ok:
            return True
        else:
            response.raise_for_status()
