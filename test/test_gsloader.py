from unittest import TestCase

from mock import patch, call, Mock
from nose.plugins.attrib import attr
import gspread

from module.gsloader import SpreadSheetLoader


@attr('unit')
class SpreadSheetLoaderTest(TestCase):
    def setUp(self):
        self.loader = SpreadSheetLoader('')
        self.window = 7
        self.spread_sheet_id = "test-id"

    def create_patch(self, *args):
        get_patcher = patch if len(args) == 1 else patch.object
        patcher = get_patcher(*args)
        res = patcher.start()
        self.addCleanup(patcher.stop)
        return res

    def test_calculate_moving_average_no_data(self):
        get_data = self.create_patch(self.loader, 'get_data_from_worksheet')
        sort_data = self.create_patch(self.loader, 'sort_data_by_date')
        get_col = self.create_patch(self.loader, 'get_column_values')
        ma = self.create_patch(self.loader, 'moving_average')
        add_col = self.create_patch(self.loader, 'add_new_column')
        get_data.return_value = "", [], []

        result = self.loader.calculate_moving_average()
        self.assertIsNone(result)
        sort_data.assert_not_called()
        get_col.assert_not_called()
        ma.assert_not_called()
        add_col.assert_not_called()

    def test_calculate_moving_average_data_present(self):
        test_headers = ['Numbers', 'Dates']
        test_data = [['1', 'Date1'], ['2', 'Date2'], ['3', 'Date3'],
                     ['4', 'Date4'], ['5', 'Date5'], ['6', 'Date6']]
        get_data = self.create_patch(self.loader, 'get_data_from_worksheet')
        get_data.return_value = '', test_headers, test_data
        sort_data = self.create_patch(self.loader, 'sort_data_by_date')
        sort_data.return_value = test_data
        get_col = self.create_patch(self.loader, 'get_column_values')
        get_col.return_value = []
        ma = self.create_patch(self.loader, 'moving_average')
        ma.return_value = []
        add_col = self.create_patch(self.loader, 'add_new_column')

        result = self.loader.calculate_moving_average()
        self.assertIsNone(result)
        sort_data.assert_called_once_with(test_data, test_headers)
        get_cal_calls = [call(test_data, test_headers, 'Visitors', int), call(test_data, test_headers, 'Date')]
        get_col.assert_has_calls(get_cal_calls)
        ma.assert_called_once_with([], 5)
        add_col_calls = [call('', 4, [], 'Sorted Date'), call('', 3, [], 'Moving Average of period 5')]
        add_col.assert_has_calls(add_col_calls, any_order=True)

    def test_get_data_from_worksheet_no_data(self):
        creds = self.create_patch('oauth2client.service_account.ServiceAccountCredentials.from_json_keyfile_name')
        creds.return_value = ''
        gc = self.create_patch('gspread.authorize')
        gc.return_value.open_by_key.return_value.get_worksheet.return_value.get_all_values.return_value = []

        self.assertRaises(ValueError, self.loader.get_data_from_worksheet, 'www111')

    def test_get_data_from_worksheet_no_spreadsheet(self):
        creds = self.create_patch('oauth2client.service_account.ServiceAccountCredentials.from_json_keyfile_name')
        creds.return_value = ''
        gc = self.create_patch('gspread.authorize')
        gc.return_value.open_by_key.side_effect = gspread.exceptions.SpreadsheetNotFound
        self.assertRaises(gspread.exceptions.SpreadsheetNotFound, self.loader.get_data_from_worksheet, 'www111')

    def test_get_data_from_worksheet_data_present(self):
        creds = self.create_patch('oauth2client.service_account.ServiceAccountCredentials.from_json_keyfile_name')
        creds.return_value = ''
        gc = self.create_patch('gspread.authorize')
        gc.return_value.open_by_key.return_value.\
            get_worksheet.return_value.get_all_values.return_value = [['Visitors', 'Date']]
        res_ws, res_headers, res_data = self.loader.get_data_from_worksheet('www111')
        self.assertEqual(res_headers, ['Visitors', 'Date'])
        self.assertEqual(res_data, [])

    def test_sort_data_by_date(self):
        headers = ['Visitors', 'Date']
        data = [['777', '10/24/16'], ['122', '11/23/16'], ['111', '11/22/16']]
        res = self.loader.sort_data_by_date(data, headers)
        self.assertEqual(res, [['122', '11/23/16'], ['111', '11/22/16'], ['777', '10/24/16']])

    def test_sort_data_by_date_no_date_column(self):
        headers = ['Visitors', 'Datesssss']
        data = [['777', '10/24/16'], ['122', '11/23/16'], ['111', '11/22/16']]
        self.assertRaises(ValueError, self.loader.sort_data_by_date, data, headers)

    def test_cast_int(self):
        self.assertEqual(self.loader.cast_int('12', 0), 12)
        self.assertEqual(self.loader.cast_int('12.1', 0), 0)
        self.assertEqual(self.loader.cast_int('aaa', 0), 0)
        self.assertEqual(self.loader.cast_int('0', 7), 0)

    def test_moving_average(self):
        res = self.loader.moving_average([1, 1, 1, 1, 2, 2, 2], 5)
        self.assertListEqual(res, [1.2, 1.4, 1.6])
        res = self.loader.moving_average([1, 100, 1, 100, 1, 100, 1], 2)
        self.assertListEqual(res, [50.5, 50.5, 50.5, 50.5, 50.5, 50.5])

    def test_col_name_spreasheet_col_name(self):
        self.assertEqual('A', self.loader.col_name_to_spreadsheet_col_name(1))
        self.assertEqual('E', self.loader.col_name_to_spreadsheet_col_name(5))
        self.assertEqual('Z', self.loader.col_name_to_spreadsheet_col_name(26))
        self.assertEqual('AZ', self.loader.col_name_to_spreadsheet_col_name(2*26))
        self.assertEqual('ALL', self.loader.col_name_to_spreadsheet_col_name(1000))

    def test_add_new_column(self):
        worksheet = Mock()
        test_cells_update = ['D1', 'D2', 'D3']
        worksheet.range.return_value = test_cells_update
        col_to_symbol = self.create_patch(self.loader, 'col_name_to_spreadsheet_col_name')
        col_to_symbol.return_value = 'D'
        self.create_patch(self.loader, 'setcell')
        self.loader.add_new_column(worksheet, 0, [10]*10, 'Field')
        worksheet.add_cols.assert_called_once_with(1)
        worksheet.range.assert_called_once_with('D1:D11')
        worksheet.update_cells.assert_called_once_with(test_cells_update)
