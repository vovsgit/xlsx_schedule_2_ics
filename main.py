import os
import sys
from typing import List
import openpyxl.workbook


def check_args(args: List[str]) -> List[str]:
    if len(args) == 3:
        return [args[1], args[2]]
    else:
        raise ValueError("Invalid number of arguments")


def get_files(group: str) -> List[openpyxl.workbook]:
    dir_path = './schedules'

    if group == 'all':
        files = [os.path.join(dir_path, f) for f in os.listdir(dir_path) if '.xlsx' in f]
        return [openpyxl.load_workbook(file) for file in files]
    else:
        files = [os.path.join(dir_path, f) for f in os.listdir(dir_path) if group in f]
        return [openpyxl.load_workbook(files[0])]

def parse_workbook(workbook: openpyxl.workbook) -> List[List[str]]:
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
        data.append(row)
    return data

if __name__ == "__main__":
    args = sys.argv
    group, english_group = check_args(args)
    workbooks = get_files(group)
