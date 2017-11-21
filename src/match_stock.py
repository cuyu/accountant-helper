#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: Curtis Yu
@contact: cuyu@splunk.com
@since: 21/11/2017
"""
import re
import xlwings


def match_account(excel_path, sheet_name, match_column1, match_column2, account_column1, account_column2, regex1=None, regex2=None, last_row=1000):
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
    :param last_row: the number of last row
    :return:
    """
    wb = xlwings.Book(excel_path)
    sheet = wb.sheets[sheet_name]
    match_cells1 = sheet.range('{0}1:{1}{2}'.format(match_column1, match_column1, str(last_row)))
    match_cells2 = sheet.range('{0}1:{1}{2}'.format(match_column2, match_column2, str(last_row)))
    # todo: use regex2 to clean the match_cells2 first
    match_num = 0
    for cell in match_cells1:
        if cell.value:
            result = re.search(regex1, cell.value)
            if result:
                # get the IDs, e.g. '1344,3233,6764'
                id_string = result.group(1)
                id_list = re.split(r'[,/\s]\s*', id_string)  # todo:handle if not matched by regex split
                account_value1 = sheet.range('{0}{1}'.format(account_column1, str(cell.row))).value
                account_value2 = 0
                for id in id_list:
                    matched_row = match_string_in_column(match_cells2, id)
                    if matched_row:
                        account_value2 += sheet.range('{0}{1}'.format(account_column2, str(matched_row))).value

                # mark these lines if the account is matched?
                if account_value1 == account_value2:
                    print(id_string)
                    match_num+=1

    print('======================')
    print(match_num)


def match_string_in_column(column_cells, id):
    for cell in column_cells:
        if cell.value:
            if isinstance(cell.value, float):
                value = str(int(cell.value))
            else:
                value = str(cell.value)

            if value == id:
                return cell.row
    else:
        return None


if __name__ == '__main__':
    match_account('/Users/CYu/Downloads/苹果调节表201710.xlsx', '入库差异', 'F', 'J', 'E', 'L', '\(([\d,/\s]+)\)', last_row=2000)
