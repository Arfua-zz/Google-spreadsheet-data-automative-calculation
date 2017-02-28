import numpy as np
from datetime import datetime
import ConfigParser

import gspread
from oauth2client.service_account import ServiceAccountCredentials


class SpreadSheetLoader(object):
    def __init__(self, config_folder):
        self.configParser = ConfigParser.RawConfigParser()
        configFilePath = config_folder + 'settings.txt'
        self.configParser.read(configFilePath)
        self.scope = self.configParser.get('config', 'scope')
        self.json_file = config_folder + self.configParser.get('config', 'jsonCredentialsFile')
        self.window = int(self.configParser.get('config', 'movingAverageWindow'))
        self.spread_sheet_id = self.configParser.get('config', 'spreadSheetId')

    def calculate_moving_average(self):
        # must share spreadsheet with "email from loaded json file"
        worksheet, headers, list_of_lists = self.get_data_from_worksheet(self.spread_sheet_id)
        if len(list_of_lists) <= self.window:
            print "Not enough data to calculate moving average with window %s" % (self.window)
            return
        list_of_lists = self.sort_data_by_date(list_of_lists, headers)
        visitor_number = self.get_column_values(list_of_lists, headers, 'Visitors', int)
        dates = self.get_column_values(list_of_lists, headers, 'Date')
        visitor_number_ma = self.moving_average(visitor_number, self.window)
        self.add_new_column(worksheet, len(list_of_lists[0]) + 1, visitor_number_ma,
                            'Moving Average of period {}'.format(self.window))
        self.add_new_column(worksheet, len(list_of_lists[0]) + 2, dates, 'Sorted Date')

    def get_data_from_worksheet(self, spread_sheet_id):
        try:
            credentials = ServiceAccountCredentials.from_json_keyfile_name(self.json_file, self.scope)
            gc = gspread.authorize(credentials)
            wks = gc.open_by_key(spread_sheet_id)
            worksheet = wks.get_worksheet(0)
            list_of_lists = worksheet.get_all_values()
        except gspread.exceptions.SpreadsheetNotFound as e:
            print "No spreadsheet was found with id: {}".format(spread_sheet_id)
            raise e

        if len(list_of_lists) == 0:
            raise ValueError('No data in spreadsheet with id: {}'.format(spread_sheet_id))
        else:
            headers = list_of_lists[0]
        return worksheet, headers, list_of_lists[1:]

    def sort_data_by_date(self, list_of_lists, headers):
        # works for only specific format '%m/%d/%Y', e.g. '10/23/17'
        # other formats can be added separately
        i = headers.index('Date')
        return sorted(list_of_lists, key=lambda x: datetime.strptime(x[i][:-2]+'20'+x[i][-2:], '%m/%d/%Y'),
                      reverse=True)

    def get_column_values(self, list_of_lists, headers, column_name, value_type=str):
        if value_type not in [int, str]:
            raise AssertionError('Wrong type of values called in loaded column!')
        i = headers.index(column_name)
        list_of_values = [value[i] for value in list_of_lists]
        if value_type == int:
            return map(lambda x: self.cast_int(x, 0), list_of_values)
        return list_of_values

    def cast_int(self, x, default_int_value):
        try:
            return int(x)
        except ValueError as e:
            return default_int_value

    def moving_average(self, values, window):
        weights = np.repeat(1.0, window) / window
        sma = np.convolve(values, weights, 'valid')
        sma = list([round(x, 2) for x in sma])
        return sma

    def add_new_column(self, worksheet, number_of_cols, data_to_add, column_name):
        worksheet.add_cols(1)
        number_of_rows = len(data_to_add) + 1
        symbol = self.col_name_to_spreadsheet_col_name(number_of_cols)
        cell_list_to_update = worksheet.range("%s%s:%s%s" % (symbol, 1, symbol, number_of_rows))
        map(self.setcell, cell_list_to_update, [column_name] + [str(v) for v in data_to_add])
        worksheet.update_cells(cell_list_to_update)

    def setcell(self, c, v):
        c.value = v

    def col_name_to_spreadsheet_col_name(self, col): # col is 1 based
        spreadsheet_col = str()
        div = col
        while div:
            (div, mod) = divmod(div-1, 26) # will return (x, 0 .. 25)
            spreadsheet_col = chr(mod + 65) + spreadsheet_col

        return spreadsheet_col


if __name__ == '__main__':
    SpreadSheetLoader('module/').calculate_moving_average()
