#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Curtis Yu
@contact: cuyu@splunk.com
@since: 21/11/2017
"""
import re
import xlwings


class MyCell(object):
    def __init__(self, value, rows):
        self.value = value  # total value
        self.rows = [rows]  # list of row number


def _clean_match_column1(match_cells, account_column, regex):
    """
    Extract content according to regex and transfer column into a dict.
    Merge the rows into one if they are the same (add the account together).
    :return: a dict which key is the extracted content (a list of id) and value is account number
    """
    result = dict()
    for cell in match_cells:
        if cell.value:
            match = re.search(regex, cell.value)
            if match:
                # get the IDs, e.g. '1344,3233,6764'
                id_string = match.group(1)
                id_list = tuple(re.split(r'[,/\s]\s*', id_string))  # todo:handle if not matched by regex split
                if id_list in result:
                    result[id_list].value += match_cells.sheet.range(
                        '{0}{1}'.format(account_column, str(cell.row))).value
                    result[id_list].rows.append(cell.row)
                else:
                    result[id_list] = MyCell(
                        match_cells.sheet.range('{0}{1}'.format(account_column, str(cell.row))).value, cell.row)
    return result


def _clean_match_column2(match_cells, account_column):
    """
    :return: a dict which key is the id and value is account number
    """
    result = dict()
    for cell in match_cells:
        if cell.value:
            if isinstance(cell.value, float):
                value = str(int(cell.value))
            else:
                value = str(cell.value)
            assert value not in result
            result[value] = MyCell(match_cells.sheet.range('{0}{1}'.format(account_column, str(cell.row))).value,
                                   cell.row)
    return result


def match_account(excel_path, sheet_name, match_column1, match_column2, account_column1, account_column2, regex1=None,
                  regex2=None, mark_color=(0, 200, 100,), last_row=1000):
    """
    Match the account and remove the matched ones in a new excel file.
    :param excel_path: full path of the excel file
    :param sheet_name: name of the sheet tab
    :param match_column1: patterns extracted in this column are used to match patterns in match_column2, e.g. 'F'
    :param match_column2: patterns extracted in this column are used to match patterns in match_column1
    :param account_column1: account in this column and in the same row of match_column1 is compared to account_column2
    :param account_column2: account in this column and in the same row of match_column2 is compared to account_column1
    :param regex1: regex to extract patterns in match_column1
    :param regex2: regex to extract patterns in match_column2
    :param mark_color: the RGB color to mark the matched cell's background
    :param last_row: the number of last row
    :return:
    """
    wb = xlwings.Book(excel_path)
    sheet = wb.sheets[sheet_name]
    match_cells1 = sheet.range('{0}1:{1}{2}'.format(match_column1, match_column1, str(last_row)))
    match_cells2 = sheet.range('{0}1:{1}{2}'.format(match_column2, match_column2, str(last_row)))
    match_dict1 = _clean_match_column1(match_cells1, account_column1, regex1)
    # todo: use regex2 to clean the match_cells2 first
    match_dict2 = _clean_match_column2(match_cells2, account_column2)
    match_num = 0
    for id_list in match_dict1:
        account_value2 = 0
        match_rows2 = []
        for id in id_list:
            my_cell = match_dict2.get(id)
            if my_cell:
                account_value2 += my_cell.value
                match_rows2 += my_cell.rows

        if int(match_dict1[id_list].value) == int(account_value2):
            for row in match_dict1[id_list].rows:
                sheet.range('{0}{1}'.format(match_column1, row)).color = mark_color
            for row in match_rows2:
                sheet.range('{0}{1}'.format(match_column2, row)).color = mark_color
            print(id_list)
            match_num += 1

    print('===========Total matched===========')
    print(match_num)


if __name__ == '__main__':
    match_account('/Users/CYu/Downloads/苹果调节表201710(1).xlsx', '入库差异', 'F', 'J', 'E', 'L', '\(([\d,/\s]+)\)', last_row=3000)
    # match_account('/Users/CYu/Downloads/苹果调节表201710(1).xlsx', '入库差异', 'F', 'J', 'E', 'L', 'DN\:([\d,/\s]+)',
    #               mark_color=(0, 100, 200), last_row=3000)
