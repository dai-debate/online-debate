from typing import Union, List, Dict, Any
from gspread.models import Worksheet
import gspread.utils as gsutils


class WorksheetEx(Worksheet):
    """gspread.models.Worksheet の拡張"""

    def __init__(self, spreadsheet, properties):
        super.__init__(spreadsheet, properties)

    @classmethod
    def cast(cls, obj: Worksheet):
        """Worksheet オブジェクトを WorksheetEx に拡張する

        :param obj: 変換元のオブジェクト
        :type obj: Worksheet
        :return: 拡張されたオブジェクト
        :rtype: WorksheetEx
        """
        obj.__class__ = cls
        return obj

    @gsutils.cast_to_a1_notation
    def add_protected_range(
        self,
        name: str,
        editor_users_emails: List[str] = None,
        editor_groups_emails: List[str] = None,
        description: str = None,
        warning_only: bool = False,
        requesting_user_can_edit: bool = False,
    ):
        """保護された範囲を追加する

        :param name: 範囲の名前
        :type name: str
        :param editor_users_emails: 編集を許可するユーザーのメールアドレスの list, defaults to None
        :type editor_users_emails: List[str], optional
        :param editor_groups_emails: 編集を許可するグループのメールアドレスの list, defaults to None
        :type editor_groups_emails: List[str], optional
        :param description: 説明, defaults to None
        :type description: str, optional
        :param warning_only: 編集を警告するだけか, defaults to False
        :type warning_only: bool, optional
        :param requesting_user_can_edit: ユーザーが編集をリクエストできるか, defaults to False
        :type requesting_user_can_edit: bool, optional
        :raises PermissionError: spreadsheet に対して権限のないユーザーを指定しようとした場合に例外を発生
        """
        permitted_email_address = [
            permission.get('emailAddress')
            for permission in self.client.list_permissions(self.spreadsheet.id)
            if permission.get('emailAddress')
        ]

        editors_emails = editor_users_emails or []
        for email in editors_emails:
            if email not in permitted_email_address:
                raise PermissionError(f'{email} is not permitted to edit this spreadsheet.')

        editor_groups_emails = editor_groups_emails or []
        for email in editor_groups_emails:
            if email not in permitted_email_address:
                raise PermissionError(f'{email} is not permitted to edit this spreadsheet.')

        grid_range = gsutils.a1_range_to_grid_range(name, self.id)

        body = {
            "requests": [
                {
                    "addProtectedRange": {
                        'protectedRange': {
                            "range": grid_range,
                            "description": description,
                            "warningOnly": warning_only,
                            "requestingUserCanEdit": requesting_user_can_edit,
                            "editors": {
                                "users": editors_emails,
                                "groups": editor_groups_emails,
                            },
                        }
                    }
                }
            ]
        }

        self.spreadsheet.batch_update(body)

    def get_protected_ranges(self) -> List[Dict[str, Any]]:
        """保護された範囲のリストを取得する

        :return: 保護された範囲のリスト
        :rtype: List[Dict[str, Any]]
        """
        metadata = self.spreadsheet.fetch_sheet_metadata()
        ranges = metadata['sheets'][self.id]['protectedRanges'] if 'protectedRanges' in metadata['sheets'][self.id] else []
        result = []
        for r in ranges:
            if 'range' in r:
                r['range'] = (f'{gsutils.rowcol_to_a1(1+r["range"]["startRowIndex"],1+r["range"]["startColumnIndex"])}'
                              + ':'
                              + f'{gsutils.rowcol_to_a1(r["range"]["endRowIndex"],1+r["range"]["endColumnIndex"])}')
            if 'unprotectedRanges' in r:
                del r['unprotectedRanges']
            result.append(r)
        return result

    def clear_protected_ranges(self, ids: Union[List[str], None] = None):
        """保護された範囲を解除する

        :param ids: 保護された範囲のIDのリスト. None を指定すると全ての範囲を解除する, defaults to None
        :type ids: Union[List[str], None], optional
        """
        if ids is None:
            ids = [range['protectedRangeId'] for range in self.get_protected_ranges()]
        if len(ids) > 0:
            self.spreadsheet.batch_update({
                "requests": [{"deleteProtectedRange": {"protectedRangeId": id}} for id in ids]
            })

    class ConditionType:
        """条件付き書式のタイプ定義"""

        def __init__(self):
            pass

        NUMBER_GREATER = 'NUMBER_GREATER'
        '''
        The cell's value must be greater than the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        NUMBER_GREATER_THAN_EQ = 'NUMBER_GREATER_THAN_EQ'
        '''
        The cell's value must be greater than or equal to the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        NUMBER_LESS = 'NUMBER_LESS'
        '''
        The cell's value must be less than the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        NUMBER_LESS_THAN_EQ = 'NUMBER_LESS_THAN_EQ'
        '''
        The cell's value must be less than or equal to the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        NUMBER_EQ = 'NUMBER_EQ'
        '''
        The cell's value must be equal to the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        NUMBER_NOT_EQ = 'NUMBER_NOT_EQ'
        '''
        The cell's value must be not equal to the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        NUMBER_BETWEEN = 'NUMBER_BETWEEN'
        '''
        The cell's value must be between the two condition values.
        Supported by data validation, conditional formatting and filters.
        Requires exactly two ConditionValues .
        '''
        NUMBER_NOT_BETWEEN = 'NUMBER_NOT_BETWEEN'
        '''
        The cell's value must not be between the two condition values.
        Supported by data validation, conditional formatting and filters.
        Requires exactly two ConditionValues .
        '''
        TEXT_CONTAINS = 'TEXT_CONTAINS'
        '''
        The cell's value must contain the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        TEXT_NOT_CONTAINS = 'TEXT_NOT_CONTAINS'
        '''
        The cell's value must not contain the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        TEXT_STARTS_WITH = 'TEXT_STARTS_WITH'
        '''
        The cell's value must start with the condition's value.
        Supported by conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        TEXT_ENDS_WITH = 'TEXT_ENDS_WITH'
        '''
        The cell's value must end with the condition's value.
        Supported by conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        TEXT_EQ = 'TEXT_EQ'
        '''
        The cell's value must be exactly the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        TEXT_IS_EMAIL = 'TEXT_IS_EMAIL'
        '''
        The cell's value must be a valid email address.
        Supported by data validation.
        Requires no ConditionValues .
        '''
        TEXT_IS_URL = 'TEXT_IS_URL'
        '''
        The cell's value must be a valid URL.
        Supported by data validation.
        Requires no ConditionValues .
        '''
        DATE_EQ = 'DATE_EQ'
        '''
        The cell's value must be the same date as the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        DATE_BEFORE = 'DATE_BEFORE'
        '''
        The cell's value must be before the date of the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue that may be a relative date .
        '''
        DATE_AFTER = 'DATE_AFTER'
        '''
        The cell's value must be after the date of the condition's value.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue that may be a relative date .
        '''
        DATE_ON_OR_BEFORE = 'DATE_ON_OR_BEFORE'
        '''
        The cell's value must be on or before the date of the condition's value.
        Supported by data validation.
        Requires a single ConditionValue that may be a relative date .
        '''
        DATE_ON_OR_AFTER = 'DATE_ON_OR_AFTER'
        '''
        The cell's value must be on or after the date of the condition's value.
        Supported by data validation.
        Requires a single ConditionValue that may be a relative date .
        '''
        DATE_BETWEEN = 'DATE_BETWEEN'
        '''
        The cell's value must be between the dates of the two condition values.
        Supported by data validation.
        Requires exactly two ConditionValues .
        '''
        DATE_NOT_BETWEEN = 'DATE_NOT_BETWEEN'
        '''
        The cell's value must be outside the dates of the two condition values.
        Supported by data validation.
        Requires exactly two ConditionValues .
        '''
        DATE_IS_VALID = 'DATE_IS_VALID'
        '''
        The cell's value must be a date.
        Supported by data validation.
        Requires no ConditionValues .
        '''
        ONE_OF_RANGE = 'ONE_OF_RANGE'
        '''
        The cell's value must be listed in the grid in condition value's range.
        Supported by data validation.
        Requires a single ConditionValue , and the value must be a valid range in A1 notation.
        '''
        ONE_OF_LIST = 'ONE_OF_LIST'
        '''
        The cell's value must be in the list of condition values.
        Supported by data validation.
        Supports any number of condition values , one per item in the list.
        Formulas are not supported in the values.
        '''
        BLANK = 'BLANK'
        '''
        The cell's value must be empty.
        Supported by conditional formatting and filters.
        Requires no ConditionValues .
        '''
        NOT_BLANK = 'NOT_BLANK'
        '''
        The cell's value must not be empty.
        Supported by conditional formatting and filters.
        Requires no ConditionValues .
        '''
        CUSTOM_FORMULA = 'CUSTOM_FORMULA'
        '''
        The condition's formula must evaluate to true.
        Supported by data validation, conditional formatting and filters.
        Requires a single ConditionValue .
        '''
        BOOLEAN = 'BOOLEAN'
        '''
        The cell's value must be TRUE/FALSE or in the list of condition values.
        Supported by data validation.
        Renders as a cell checkbox.
        Supports zero, one or two ConditionValues .
        No values indicates the cell must be TRUE or FALSE, where TRUE renders as checked and FALSE renders as unchecked.
        One value indicates the cell will render as checked when it contains that value and unchecked when it is blank.
        Two values indicate that the cell will render as checked when it contains the first value and unchecked when it contains the second value.
        For example, ["Yes","No"] indicates that the cell will render a checked box when it has the value "Yes" and an unchecked box when it has the value "No".
        '''

    conditiontype = ConditionType()

    class RelativeDate:
        """相対日時のタイプ"""

        def __init__(self):
            pass

        PAST_YEAR = 'PAST_YEAR'
        '''
        The value is one year before today.
        '''
        PAST_MONTH = 'PAST_MONTH'
        '''
        The value is one month before today.
        '''
        PAST_WEEK = 'PAST_WEEK'
        '''
        The value is one week before today.
        '''
        YESTERDAY = 'YESTERDAY'
        '''
        The value is yesterday.
        '''
        TODAY = 'TODAY'
        '''
        The value is today.
        '''
        TOMORROW = 'TOMORROW'
        '''
        The value is tomorrow.
        '''

    relativedate = RelativeDate()

    @gsutils.cast_to_a1_notation
    def add_conditional_format(self, name: str, cond_type: str, cond_values: List[Union[int, float, str]], cond_format: dict):
        """条件付き書式を追加する

        :param name: 条件付き書式を設定するレンジ
        :type name: str
        :param cond_type: 条件のタイプ
        :type cond_type: str
        :param cond_values: 条件の値
        :type cond_values: List[Union[int, float, str]]
        :param cond_format: 書式
        :type cond_format: dict
        """
        grid_range = gsutils.a1_range_to_grid_range(name, self.id)
        cv = []
        for v in cond_values:
            if v in WorksheetEx.relativedate.__dict__.keys():
                cv.append({'relativeDate': v})
            elif type(v) is str:
                cv.append({'userEnteredValue': v})
            else:
                cv.append({'userEnteredValue': str(v)})

        self.spreadsheet.batch_update({
            "requests": [{
                "addConditionalFormatRule": {
                    "rule": {
                        'ranges': [grid_range],
                        'booleanRule': {
                            'condition': {
                                'type': cond_type,
                                'values': cv
                            },
                            'format': cond_format
                        }
                    },
                    'index': 0
                }
            }]
        })

    @gsutils.cast_to_a1_notation
    def add_gradient_format(self, name: str, grad_rule: Dict[str, Any]):
        grid_range = gsutils.a1_range_to_grid_range(name, self.id)

        self.spreadsheet.batch_update({
            "requests": [{
                "addConditionalFormatRule": {
                    "rule": {
                        'ranges': [grid_range],
                        'gradientRule': grad_rule
                    },
                    'index': 0
                }
            }]
        })

    @gsutils.cast_to_a1_notation
    def set_data_validation(self, name: str, cond_type: str, cond_values: List[Union[int, float, str]],
                            message: Union[str, None] = None, strict: bool = False, custom_ui: bool = False):
        """データの入力規則を追加する

        :param name: データの入力規則を設定するレンジ
        :type name: str
        :param cond_type: 条件のタイプ
        :type cond_type: str
        :param cond_values: 条件の値
        :type cond_values: List[Union[int, float, str]]
        :param message: 入力時に表示するメッセージ, defaults to None
        :type message: Union[str, None], optional
        :param strict: 条件に一致しない値を拒否する, defaults to False
        :type strict: bool, optional
        :param custom_ui: プルダウンリストを表示する, defaults to False
        :type custom_ui: bool, optional
        """
        grid_range = gsutils.a1_range_to_grid_range(name, self.id)
        cv = []
        for v in cond_values:
            if v in WorksheetEx.relativedate.__dict__.keys():
                cv.append({'relativeDate': v})
            elif type(v) is str:
                cv.append({'userEnteredValue': v})
            else:
                cv.append({'userEnteredValue': str(v)})

        rule = {
            'condition': {
                'type': cond_type,
                'values': cv
            },
            'strict': strict,
            'showCustomUi': custom_ui
        }
        if message is not None:
            rule['inputMessage'] = message

        self.spreadsheet.batch_update({
            "requests": [{
                "setDataValidation": {
                    'range': grid_range,
                    "rule": rule
                }
            }]
        })

    def set_dimention_size(self, dimention: str, start: int, end: int, size: int):
        self.spreadsheet.batch_update({
            'requests': [{
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': self.id,
                        'dimension': dimention,
                        'startIndex': start,
                        'endIndex': end+1
                    },
                    'fields': '*',
                    'properties': {
                        'pixelSize': size
                    }
                }
            }]
        })

    def clear_freeze(self):
        self.spreadsheet.batch_update({
            'requests': [{
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': self.id,
                        'gridProperties': {
                            'frozenRowCount': 0,
                            'frozenColumnCount': 0
                        }
                    },
                    'fields': 'gridProperties/frozenRowCount,gridProperties/frozenColumnCount'
                }
            }]
        })
