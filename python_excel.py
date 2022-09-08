import datetime
import os
from typing import List, Dict

import xlsxwriter


def excel_create(filename: str, data: List[Dict]) -> None:

    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    for col, name in enumerate(data[0].keys()):
        worksheet.write(0, col, name)

    for row, item in enumerate(data, start=1):
        items_dict = item.values()
        for col, value in enumerate(items_dict):
            worksheet.write(row, col, value)

    workbook.close()


def test_case():
    data = [
        {
            'id': 0,
            'date': datetime.datetime.now().strftime("%d.%m.%YT%H:%M:%S"),
            'name': 'Some name',
            'status': 'Success',
            'message': '',
        }, {
            'id': 1,
            'date': datetime.datetime.now().strftime("%d.%m.%YT%H:%M:%S"),
            'name': 'Task force',
            'status': 'Fail',
            'message': 'Error in host',
        }, {
            'id': 2,
            'date': datetime.datetime.now().strftime("%d.%m.%YT%H:%M:%S"),
            'name': 'Task',
            'status': 'Fail',
            'message': 'Error. Did not send',
        }, {
            'id': 3,
            'date': datetime.datetime.now().strftime("%d.%m.%YT%H:%M:%S"),
            'name': 'Task',
            'status': 'Success',
            'message': 'Field in my.py not found',
        },
    ]
    this_folder = os.path.dirname(os.path.abspath(__file__))
    my_file = os.path.join(this_folder, 'static/test.xlsx')
    excel_create(my_file, data)


test_case()
