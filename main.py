import xlrd
from os.path import join, normpath, basename
from os import listdir, mkdir
from datetime import datetime
from ets.ets_ds_lib import get_party_check
from time import sleep
from shutil import move
from config import *
from codes import *

input_dir = normpath(input_dir)
out_dir = normpath(out_dir)


def write_log(file, text):
    print(text)
    with open(file, mode='a', encoding='utf8') as log:
        log.write(str(text) + '\r\n')


while True:
    input_files = [file for file in listdir(input_dir) if (file.endswith('.xlsx') or file.endswith('.xls'))]
    if input_files:
        start_datetime = datetime.now()

        exec_out_dir = join(out_dir, start_datetime.strftime('%Y%m%d_%H%M%S'))
        mkdir(exec_out_dir)

        log_file = join(exec_out_dir, log_name)

        event = 'Starting %s' % start_datetime
        write_log(log_file, event)

        for file in input_files:
            file = join(input_dir, file)

            event = 'Working %s' % file
            write_log(log_file, event)

            excel_file = xlrd.open_workbook(file)
            excel_sheet = excel_file.sheet_by_index(0)

            parsed_data = list(excel_sheet.get_rows())[2:]

            for s in parsed_data:

                s = [str(v.value).strip().replace('.0', '') for v in s][0:5]
                s = map(lambda x: bank_codes.get(x, x), s)
                s = list(map(lambda x: type_codes.get(x, x), s))

                if not s[0]:
                    continue

                if not (s[3] and s[4]):
                    event = '%s: have not all data. Ignored' % s[0]
                    write_log(log_file, event)
                    continue

                guid, error = get_party_check(s[3], s[4], s[0], kpp=s[1], ogrn=s[2])

                if error:
                    event = '%s: send to %-7s with error %s' % (s[0], s[3], error)
                else:
                    event = '%s: send to %-7s with GUID %s' % (s[0], s[3], guid)

                write_log(log_file, event)

                sleep(send_awaiting_time)

            move(file, exec_out_dir)
            event = 'File %s is done' % basename(file)
            write_log(log_file, event)

        end_datetime = datetime.now()
        event = 'Finished %s' % end_datetime
        write_log(log_file, event)

    sleep(daemon_awaiting_time)






