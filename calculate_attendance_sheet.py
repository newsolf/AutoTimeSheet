#!/usr/bin/python3
import xlrd
import xlwt
import os
from pathlib import Path
import datetime

# import win32com.client

ADVANCE_START = '05:01'
ADVANCE_END = '09:00'
ADVANCE_MONEY = 15
GO_LATE_START = '20:29'
GO_LATE_END = '23:59'
GO_LATE_TOMORROW_START = '00:00'
GO_LATE_TOMORROW_END = '05:00'
GO_LATE_MONEY = 25
IGNORE_NAME = '姓名'
DIR_PATH = "files"


def find_attendance_sheet_file():
    dir_path = Path(DIR_PATH)
    if 1 - dir_path.exists():
        os.mkdir(DIR_PATH)
        print("mkdir path = %s ,please set file in %s " % dir_path)
        return ""

    if 1 - dir_path.is_dir():
        os.remove(dir_path)
        os.mkdir(DIR_PATH)
        print("mkdir path = %s ,please set file in %s " % DIR_PATH, DIR_PATH)
        return None

    file_list = os.listdir(DIR_PATH)
    if len(file_list) == 0:
        print("please set file in %s " % DIR_PATH)
        return None

    file_name = ""
    for index in range(len(file_list)):
        file = file_list[index]
        if file.endswith("result.xls"):
            print('need delete file %s' % file)
            os.remove(os.path.join(DIR_PATH, file))

        if file.endswith(".xls") and 1 - file.endswith("result.xls"):
            file_name = str(file)
            return os.path.join(DIR_PATH, file)


def pwd_xlsx(old_filename, new_filename, pwd_str, pw_str=''):
    xcl = win32com.client.Dispatch("Excel.Application")
    # pw_str为打开密码, 若无 访问密码, 则设为 ''
    wb = xcl.Workbooks.Open(old_filename, False, False, None, pw_str)
    xcl.DisplayAlerts = False

    # 保存时可设置访问密码.
    wb.SaveAs(new_filename, None, pwd_str, '')


global result_sheet  # result_sheet


def get_current_time():
    return datetime.datetime.now().strftime('%Y%m%d%H%M%S')


def calculate_attendance_sheet():
    import time
    result_cols = 0
    start_time = time.time()
    # print('start_time = %s' % start_time)
    file_name = find_attendance_sheet_file()
    if file_name == "" or file_name is None:
        print("no file,break")
        return
    print("file is %s" % file_name)

    data = xlrd.open_workbook(file_name)  # 读取数据

    result = xlwt.Workbook()
    result_title = ['姓名', '起早', '金额（15元/天）', '贪黑', '金额（25元/天）', '起早贪黑合计补助（元）']
    page = len(data.sheets())  # 获取sheet的数量
    for i in range(page):
        table = data.sheets()[i]
        # print(table.name, i)
        result_sheet = result.add_sheet(table.name)  # 写sheet name
        result_cols = 0
        for row_result in range(len(result_title)):
            result_sheet.write(result_cols, row_result, result_title[row_result])  # 写标题

        n_rows = table.nrows  # 获取总行数
        n_cols = table.ncols  # 获取总列数
        for row in range(n_rows):
            advance = 0
            go_late = 0
            name = table.cell_value(row, 0)
            # print('name = %s , is name = %s' % (name, (name == IGNORE_NAME)))
            if name == IGNORE_NAME:
                continue
            for col in range(n_cols):
                cell_data = table.cell_value(row, col)
                string_date = str.strip(cell_data)
                string_date = string_date.replace(' ', '')
                string_date = string_date.replace('\'', '')
                split_date = string_date.split('\n')
                advance_flag = 0
                go_late_flag = 0
                for index in range(len(split_date)):
                    time = split_date[index]
                    # if 'NeWolf' == name and len(time) > 2:
                    #     print('time = %s' % time)

                    if len(time) > 6:
                        # print('len(time) > 6 time = %s' % time)
                        continue

                    if not time.__contains__(":"):
                        # if not '' == time:
                        # print('not time = %s' % time)
                        continue

                    # print('normal time = %s' % time)

                    is_advance = ADVANCE_START <= time <= ADVANCE_END
                    if is_advance:
                        advance_flag = 1
                        # print('is_advance = %s time = %s' % (is_advance, time))

                    is_late = GO_LATE_START <= time <= GO_LATE_END or \
                              GO_LATE_TOMORROW_START <= time <= GO_LATE_TOMORROW_END
                    if is_late:
                        go_late_flag = 1
                        # print('is_late = %s time = %s' % (is_late, time))

                advance += advance_flag
                go_late += go_late_flag

                # print(split_date)

            if advance != 0 or go_late != 0:
                total_advance_money = advance * ADVANCE_MONEY
                total_go_late_money = go_late * GO_LATE_MONEY
                total_money = total_advance_money + total_go_late_money

                # print(
                #     'name %s , advance = %s, total_advance_money = %s ,go_late = %s , total_go_late_money = %s , '
                #     'total_money = %s' % (
                #         name, advance, total_advance_money, go_late, total_go_late_money, total_money))
                person_stats = [name, advance, total_advance_money, go_late, total_go_late_money, total_money]
                result_cols += 1
                for row_result in range(len(person_stats)):
                    result_sheet.write(result_cols, row_result, person_stats[row_result])

    result_cols += 3
    import time
    use_time = time.time() - start_time

    result_sheet.write(result_cols, 4, "use time = %.2f s, by NeWolf" % use_time)
    file_name = file_name.split('.')
    # print(file_name[0])
    result_file_name = '%s_%s_result.xls' % (file_name[0], get_current_time())
    result.save(result_file_name)
    print("use time = %.2f s , result file is %s" % (use_time, result_file_name))
    # pwd_xlsx(result_file_name, result_file_name, "NeWolf")


if __name__ == '__main__':
    calculate_attendance_sheet()
    # write_excel()
