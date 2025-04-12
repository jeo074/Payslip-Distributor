import gc
import csv
import fpdf
import queue
import string
import requests
import traceback
import pandas as pd
import fnmatch as f
import datetime as d
from constant import *
from pathlib import Path
from copy import deepcopy
from fpdf.enums import XPos, YPos
from googleapiclient.discovery import build
from messenger_api import send_message, send_attachment, get_psids
from google_fnx import goog_auth, get_spreadsheet_data, get_sheet_values
from google_fnx import get_file_id, create_file, send_email, update_spreadsheet
from dotenv import load_dotenv

# logs_folder_id & tool_logs_folder_id - Google Drive folders for logs
# spreadsheet_id - Google Spreadsheet to get data like column index
# PSID_sheet_id - messenger PSIDs

load_dotenv()
sender_email = os.getenv("sender_email")
PSID_sheet_id = os.getenv("PSID_sheet_id")
logs_folder_id = os.getenv("logs_folder_id")
spreadsheet_id = os.getenv("spreadsheet_id")
tool_logs_folder_id = os.getenv("tool_logs_folder_id")


def get_headers(spreadsheet_data, sheets, sheet_name):
    try:
        headers = spreadsheet_data[sheets[sheet_name]]['data'][0]['rowData'][0]['values']
        headers_dict = {}
        for i in range(0, len(headers)):
            if 'formattedValue' not in headers[i]:
                break
            header_name = headers[i]['formattedValue']
            if header_name in headers_dict:
                headers_dict[header_name].append(i)
            else:
                headers_dict[header_name] = [i]
        return headers_dict
    except Exception as get_header_e:
        raise Exception('Error getting header data from Pay & Deduct sheets'.format(get_header_e))


def get_columns(spreadsheet_data, sheet_index, headers_dict, queryheader, returnheader):
    try:
        data = spreadsheet_data[sheet_index]['data'][0]['rowData']
        return_dict = {}
        sheet_name = []
        for rows in data:
            cells = rows['values']
            query_pos = headers_dict[queryheader]
            for i in range(0, len(query_pos)):
                return_pos = headers_dict[returnheader]
                for j in range(0, len(return_pos)):
                    cell = return_pos[j]
                    if len(cells) > cell:
                        if 'formattedValue' in cells[cell]:
                            param = cells[query_pos[i]]['formattedValue']
                            if param == 'Company':
                                continue
                            value = cells[cell]['formattedValue']
                            if param == 'Sheet Tab':
                                sheet_name.append(value)
                            if sheet_name[j] not in return_dict:
                                return_dict[sheet_name[j]] = {}
                            else:
                                return_dict[sheet_name[j]][param] = value
        return return_dict
    except Exception as get_columns_e:
        raise Exception('Error getting column({})'.format(str(get_columns_e)))


def get_sheet_headers(spreadsheet_data, sheet_index):
    data = spreadsheet_data[sheet_index]['data'][0]['rowData'][0]['values']
    mylist = [x['formattedValue'] for x in data if 'formattedValue' in x]
    mylist = list(set(mylist))
    mylist.remove('Company')
    return mylist


def get_sheet_data(spreadsheet_data, sheet_index):
    data = spreadsheet_data[sheet_index]['data'][0]['rowData']
    value = ''
    return_dict = {}
    for row in data:
        if 'formattedValue' in row['values'][0]:
            param = row['values'][0]['formattedValue']
            if 'formattedValue' in row['values'][1]:
                value = row['values'][1]['formattedValue']
            if param == 'PARAMETER':
                continue
            return_dict[param] = value
    return return_dict


def get_data(sheets, input_file):
    vals_dict = {}
    sheet_names = pd.ExcelFile(input_file).sheet_names
    df1 = None
    df = None
    for sheet_name in sheet_names:
        for sheet, cols in sheets.items():
            if not f.fnmatch(sheet_name.lower(), sheet.lower()):
                continue
            df1 = pd.read_excel(input_file, sheet_name=sheet_name.strip())
            df = pd.DataFrame.from_dict(df1)
            start_row = 0
            last_row = df.index[-1]
            names = []
            for key, col_alpha in cols.items():
                if key == 'Start Row':
                    start_row = int(col_alpha) - 2
                elif key == 'Payroll Period':
                    temp_col = temp_row = ''
                    for char in col_alpha:
                        if char.isalpha():
                            temp_col = '{}{}'.format(temp_col, char)
                        elif char.isnumeric():
                            temp_row = '{}{}'.format(temp_row, char)
                    row = int(temp_row) - 2
                    col = col2num(temp_col) - 1
                    if ':' in df.iloc[row, col]:
                        vals_dict[key] = str(df.iloc[row, col]).split(':')[1].strip()
                    else:
                        vals_dict[key] = ' '.join([x for x in df.iloc[row, col].split() if
                                                   x.lower() not in ['payroll', 'cut', 'off', 'cutoff', 'cut-off',
                                                                     'period']]).strip()
                    if vals_dict[key] == 'nan':
                        vals_dict[key] = ''
                elif key not in ['Company', 'Sheet Tab']:
                    if col_alpha == 'nan':
                        continue
                    col = col2num(col_alpha) - 1
                    if len(df.columns) < col + 1:
                        raise Exception("'{}' data not found in column {}! Please update 'Payment and Deductions' spreadsheet.".format(key, col_alpha))
                    for j in range(start_row, last_row + 1):
                        try:
                            val = str(df.iloc[j, col])
                        except:
                            raise Exception("'{}' data not found in cell {}{}! Please update 'Payment and Deductions' spreadsheet.".format(key, col_alpha, j))
                        if val == 'nan':
                            val = ''
                        if key == 'Name':
                            if val == '':
                                a_cell = str(df.iloc[j, 0]).lower().strip()
                                # identify separation header rows
                                if a_cell == 'adjustment' or a_cell == 'training allowance':
                                    val = 'skip this row'
                                    if val in vals_dict:
                                        val = val + '!'
                                    names.append(val)
                                    vals_dict[val] = {}
                                else:
                                    last_row = j
                                    break
                            else:
                                val = val.upper()
                                while val in vals_dict:
                                    val = val + '!'
                                names.append(val)
                                vals_dict[val] = {}
                        elif j - start_row > len(names)-1:
                            break
                        elif names[j - start_row] in vals_dict and val != '':
                            if key in ['Email', 'Branch', 'ID No.']:
                                pass
                            else:
                                if val.strip() == '-':
                                    val = 0
                                elif val.replace('.', '').isnumeric():
                                    if float(val) == 0 and key not in ['Gross', 'Total', 'Net']:
                                        continue
                            vals_dict[names[j - start_row]][key] = val
    del df
    del df1
    gc.collect()
    return vals_dict


def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def create_table(table_data, title='', data_size=10, title_size=12, align_data='L', align_header='L', cell_width='even',
                 x_start='x_default', emphasize_data=None, emphasize_style=None, emphasize_size=10,
                 emphasize_color=(0, 0, 0), bottom_line=True, with_headers=True):

    if emphasize_data is None:
        emphasize_data = {}
    default_style = pdf.font_style
    if emphasize_style is None:
        emphasize_style = default_style

    # Get Width of Columns
    def get_col_widths():
        col_width = cell_width
        if col_width == 'even':
            col_width = pdf.epw / len(data[0]) - 1

        elif col_width == 'uneven':
            col_widths = []
            if isinstance(table_data, dict):
                # search for largest sized cell
                for head, col in table_data.items():
                    longest = pdf.get_string_width(head)
                    for row in col:
                        value_length = pdf.get_string_width(row)
                        if value_length > longest:
                            longest = value_length
                    col_widths.append(longest + 20)
                col_width = col_widths
            elif isinstance(table_data, list):
                for col in range(len(table_data[0])):
                    longest = 0
                    for row in range(len(table_data)):
                        cell_value = str(table_data[row][col])
                        value_length = pdf.get_string_width(cell_value)
                        if value_length > longest:
                            longest = value_length
                    col_widths.append(longest + 4)
                col_width = col_widths

        elif isinstance(cell_width, list):
            col_width = cell_width
        else:
            col_width = int(col_width)
        return col_width

    pdf.set_line_width(0.4)
    pdf.set_draw_color(220)

    if isinstance(table_data, dict):
        header = [key for key in table_data]
        data = []
        for key in table_data:
            value = table_data[key]
            data.append(value)
        # need to zip so data is in correct format (first, second, third --> not first, first, first)
        data = [list(a) for a in zip(*data)]

    else:
        header = table_data[0]
        data = table_data[1:]

    line_height = pdf.font_size * 1.75
    col_width = get_col_widths()
    pdf.set_font('Helvetica', size=title_size)

    # Determine width of table to get x starting point for centred table
    if x_start == 'C':
        table_width = 0
        if isinstance(col_width, list):
            for width in col_width:
                table_width += width
        else:
            table_width = col_width * len(table_data[0])
        margin_width = pdf.w - table_width

        center_table = margin_width / 2
        x_start = center_table
        pdf.set_x(x_start)
    elif isinstance(x_start, int):
        pdf.set_x(x_start)

    # TABLE CREATION #
    # add title
    if title != '':
        pdf.multi_cell(0, line_height, title, border=0, align='j', new_x=XPos.RIGHT, new_y=YPos.TOP,
                       max_line_height=pdf.font_size)
        pdf.ln(line_height)  # move cursor back to the left margin

    pdf.set_font('Helvetica', size=data_size)
    # add header
    y1 = pdf.get_y()
    if x_start:
        x_left = x_start
    else:
        x_left = pdf.get_x()
    x_right = pdf.epw + x_left
    if not isinstance(col_width, list):
        if x_start:
            pdf.set_x(x_start)
        if with_headers:
            for datum in header:
                pdf.set_font(style='B')
                pdf.multi_cell(col_width, line_height, datum, border=0, align=align_header, new_x=XPos.RIGHT,
                               new_y=YPos.TOP, max_line_height=pdf.font_size)
                pdf.set_font(style=default_style)
                x_right = pdf.get_x()
            pdf.ln(line_height)  # move cursor back to the left margin
            y2 = pdf.get_y()
            pdf.line(x_left, y1, x_right, y1, )
            pdf.line(x_left, y2, x_right, y2)

        for row in data:
            if x_start:
                pdf.set_x(x_start)
            for datum in row:
                if datum in emphasize_data:
                    pdf.set_text_color(*emphasize_color)
                    pdf.set_font(style=emphasize_style, size=emphasize_size)
                    pdf.multi_cell(col_width, line_height, datum, border=0, align=align_data, new_x=XPos.RIGHT,
                                   new_y=YPos.TOP, max_line_height=pdf.font_size)
                    pdf.set_text_color(0, 0, 0)
                    pdf.set_font(style=default_style)
                else:
                    pdf.multi_cell(col_width, line_height, datum, border=0, align=align_data, new_x=XPos.RIGHT,
                                   new_y=YPos.TOP, max_line_height=pdf.font_size)
            pdf.ln(line_height)  # move cursor back to the left margin

    else:
        if x_start:
            pdf.set_x(x_start)
        if with_headers:
            for i in range(len(header)):
                datum = header[i]
                pdf.set_font(style='B')
                pdf.multi_cell(col_width[i], line_height, datum, border=0, align=align_header, new_x=XPos.RIGHT,
                               new_y=YPos.TOP, max_line_height=pdf.font_size)
                pdf.set_font(style=default_style)
                x_right = pdf.get_x()
            pdf.ln(line_height)  # move cursor back to the left margin
            y2 = pdf.get_y()
            pdf.line(x_left, y1, x_right, y1)
            pdf.line(x_left, y2, x_right, y2)

        for i in range(len(data)):
            if x_start:
                pdf.set_x(x_start)
            row = data[i]
            for j in range(len(row)):
                datum = row[j]
                if not isinstance(datum, str):
                    datum = str(datum)
                adjusted_col_width = col_width[j]
                if datum in emphasize_data:
                    pdf.set_text_color(*emphasize_color)
                    pdf.set_font(style=emphasize_style, size=emphasize_size)
                    pdf.multi_cell(adjusted_col_width, line_height, datum, border=0, align=align_data, new_x=XPos.RIGHT,
                                   new_y=YPos.TOP, max_line_height=pdf.font_size)
                    pdf.set_text_color(0, 0, 0)
                    pdf.set_font(style=default_style)
                else:
                    pdf.multi_cell(adjusted_col_width, line_height, datum, border=0, align=align_data, new_x=XPos.RIGHT,
                                   new_y=YPos.TOP, max_line_height=pdf.font_size)
            pdf.ln(line_height)  # move cursor back to the left margin
    y3 = pdf.get_y()
    if bottom_line:
        pdf.line(x_left, y3, x_right, y3)


def Header(self, logo_file, title):
    # Logo
    self.image(logo_file, 10, 6, 30)
    # Helvetica bold 15
    self.set_font('Helvetica', 'B', 15)
    # Move to the right
    self.cell(80)
    # Title
    self.cell(30, 10, title, 0, 1, 'C')
    # Line break
    self.ln(20)


def payslip_pdf(logo_file, title, id_no, branch, payroll_period, pay_table, deduct_table, employee_name, gross_pay,
                total_deductions, net_pay, output_file_path):
    global pdf
    pdf = fpdf.FPDF(format='legal')
    pdf.add_page()

    # Header(pdf, logo_file, title)
    pdf.image(logo_file, 20, 6, 30)
    pdf.set_font('Helvetica', size=17, style='B')
    pdf.cell(200, 10, 'PAYSLIP', new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    pdf.set_font('Helvetica', size=10)
    pdf.cell(200, 5, company_name, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    pdf.cell(200, 5, 'No. 123 Kalye Street, Mandaluyong City, Metro Manila, 1551, Philippines', new_x=XPos.LMARGIN,
             new_y=YPos.NEXT, align='C')
    pdf.cell(200, 5, 'Office: (02) 123 4567', new_x=XPos.LEFT, new_y=YPos.NEXT, align='C')
    pdf.ln()
    pdf.set_font(style='B')
    pdf.cell(200, 5, 'Payroll Period: {}'.format(payroll_period), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    pdf.set_font("Helvetica", size=10)
    pdf.ln()
    create_table([['', ''], ['Name', employee_name], ['ID', id_no], ['Branch', branch]], '',
                 emphasize_data=[employee_name, branch], emphasize_style='B', cell_width=[35, 105], x_start=40,
                 align_header='C', align_data='L', with_headers=False, bottom_line=False)
    pdf.ln()
    create_table(pay_table, '', cell_width=[65, 55, 20], x_start=40, align_header='L', align_data='L')
    create_table([['', ''], ['Gross Pay:', gross_pay]], '', emphasize_data=['Gross Pay:', gross_pay],
                 emphasize_style='B', emphasize_size=12, cell_width=[40, 25], x_start=115, align_header='L',
                 align_data='L', with_headers=False, bottom_line=False)
    pdf.set_font("Helvetica", size=10)
    pdf.ln()

    create_table(deduct_table, 'Deductions', title_size=11, cell_width=[65, 55, 20], x_start=40, align_header='L',
                 align_data='L')

    create_table([['Deductions', total_deductions], ['NET PAY:', net_pay]], '',
                 emphasize_data=['NET PAY:', net_pay], emphasize_style='B', emphasize_size=12, cell_width=[40, 25],
                 x_start=115, align_header='L', align_data='L', bottom_line=False)
    pdf.ln()
    pdf.ln()
    pdf.set_font("Helvetica", size=10)
    pdf.cell(200, 5, 'Please examine your payslip immediately upon receipt.', new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             align='C')
    pdf.cell(200, 5, 'If no error is reported within 7 days, the account is considered accurate.', new_x=XPos.LMARGIN,
             new_y=YPos.NEXT, align='C')
    pdf.cell(200, 5, 'Your payslip is system generated. Signature is not necessary unless altered.', new_x=XPos.LMARGIN,
             new_y=YPos.NEXT,
             align='C')
    pdf.output(output_file_path)


def get_payroll_datetime(payroll_period):
    try:
        year = payroll_period.strip()[-4:]
        start_month = payroll_period.strip()[:3]
        day_split = payroll_period.strip().replace('-', ' to ').split(' to ')
        start_day = day_split[0].strip().split(' ')[1].replace(',', '')
        if len(day_split) > 1:
            if ',' in day_split[1]:
                end_day = day_split[1].split(',')[0].strip()
            else:
                end_day = day_split[1].split(' ')[-2]
            if end_day.isnumeric():
                end_month = start_month
            else:
                end_day = day_split[1].strip().split(' ')[1].replace(',', '')
                end_month = day_split[1].strip()[:3]
        else:
            end_day = start_day
            end_month = start_month
        payroll_start = d.datetime.strptime('{} {}, {}'.format(start_month, start_day, year), '%b %d, %Y').date()
        payroll_end = d.datetime.strptime('{} {}, {}'.format(end_month, end_day, year), '%b %d, %Y').date()
        return [payroll_start, payroll_end]
    except Exception:
        return [None, None]


def confirm_period(request_queue, period_queue, company, payroll_period, start_date, end_date):
    payroll_start, payroll_end = get_payroll_datetime(payroll_period)
    if payroll_start is None:
        return True
    if start_date > payroll_start or end_date < payroll_end:
        request_queue.put(['prompt', '{} payroll period of {}\nis not in between chosen scope\n\nStill continue?'.format(company, payroll_period)])
        return period_queue.get(block=True)
    else:
        return True


def add_logs(fPath, text):
    fPath = fPath + '.txt'
    text = ['{}: {}'.format(d.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), text)]
    if os.path.exists(fPath):
        with open(fPath, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(text)
    else:
        with open(fPath, 'w', newline='', encoding='utf-8') as f:
            f.write(u'\ufeff')
            writer = csv.writer(f)
            writer.writerow(text)


def record_employee(fPath, data, header):
    if os.path.exists(fPath):
        with open(fPath, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(data)
    else:
        with open(fPath, 'w', newline='', encoding='utf-8') as f:
            f.write(u'\ufeff')
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerow(data)


def check_exit(exit_queue):
    try:
        exit_result = exit_queue.get_nowait()
    except queue.Empty:
        pass
    else:
        if exit_result == 'exit':
            exit()


def main(start_date, end_date, request_queue, period_queue, exit_queue, enable_sending, payslip_type):
    # Start
    try:
        title = ''
        employee_logs_file = None

        logo_file = resource_path + r'\Data\Images\logo.png'
        file_path = payroll_path + r'\Files'
        if not os.path.exists(file_path):
            os.makedirs(file_path)
        log_path = payroll_path + '\\Tool Logs\\'
        if not os.path.exists(log_path):
            os.makedirs(log_path)
        SERVICE_ACCOUNT_KEY_FILE = resource_path + r'\Data\service_account_key.json'
        # CLIENT_SECRET = resource_path + r'\Data\client_secret.json'
        add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), 'Tool started by user {}.'.format(windows_user))
        check_exit(exit_queue)

        SCOPES = ['https://www.googleapis.com/auth/gmail.send',
                  'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file']
        creds = goog_auth(SCOPES, SERVICE_ACCOUNT_KEY_FILE)

        sheets_service = build("sheets", "v4", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)
        check_exit(exit_queue)

        # get folder ID or create if not existing
        request_queue.put(['print', 'Getting drive folder ID...'])
        year_folder_id = get_file_id(drive_service, start_date.year, 'folder', logs_folder_id)
        check_exit(exit_queue)
        if year_folder_id is None:
            request_queue.put(['print', 'Creating new drive folder...'])
            year_folder_id = create_file(drive_service, start_date.year, 'folder', logs_folder_id)

        # employee distribution logs
        rec_log_path = payroll_path + r'\Logs'
        if not os.path.exists(rec_log_path):
            os.makedirs(rec_log_path)
        log_sheet_name = '{} - {}.csv'.format(d.datetime.strftime(start_date, '%b %d'), d.datetime.strftime(end_date, '%b %d, %Y'))
        sheet_header = ['Company', 'Payroll Period', 'Employee Name', 'Email Address', 'Email Sending', 'Messenger Name', 'Messenger Sending', 'Date and Time']
        check_exit(exit_queue)

        # prepare messenger httprequest object
        msgr = requests.session()

        request_queue.put(['print', 'Getting columns and rows data...'])
        spreadsheet_data = get_spreadsheet_data(sheets_service, spreadsheet_id)
        sheets = {}
        for i in range(0, len(spreadsheet_data)):
            sheet_name = spreadsheet_data[i]['properties']['title']
            sheets[sheet_name] = i
        check_exit(exit_queue)

        pay_headers = get_headers(spreadsheet_data, sheets, 'Pay')
        deduction_headers = get_headers(spreadsheet_data, sheets, 'Deductions')
        durations = get_sheet_data(spreadsheet_data, sheets['Durations'])

        if payslip_type == 'Standard':
            param_data = get_sheet_data(spreadsheet_data, sheets['Fixed Parameters Standard'])
        elif payslip_type == 'Monthly':
            param_data = get_sheet_data(spreadsheet_data, sheets['Fixed Parameters Monthly'])
        pay_params = {}
        deduct_params = {}
        for key, val in param_data.items():
            if val.lower() == 'pay':
                pay_params[key] = {'Description': '', 'Amount': ''}
            elif val.lower() == 'deductions':
                deduct_params[key] = {'Description': '', 'Amount': ''}
        check_exit(exit_queue)

        request_queue.put(['print', 'Creating new folder for payslips...'])
        # create finished folder
        today = d.datetime.today().strftime('%Y-%m-%d')
        finished_path = file_path + '\\Finished\\{}'.format(today)

        # Get downloaded files
        file_num = 0
        request_queue.put(['print', 'Getting Excel files...'])
        inp_files_return = Path(file_path).glob('*.xls*')
        for inp in inp_files_return:
            file_num += 1
            file_name = str(inp).split('\\')[-1]
            request_queue.put(['add', [1, [file_num, file_name, '']]])
        del inp_files_return
        check_exit(exit_queue)

        if not os.path.exists(finished_path):
            os.makedirs(finished_path)
    except Exception as err:
        traceback.print_exc()
        request_queue.put(['log', ['Error: {}'.format(str(err)), 'red']])
        add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), str(err))
        exit()

    companies = get_sheet_headers(spreadsheet_data, sheets['Pay'])

    file_queue = 0
    input_files_init = Path(file_path).glob('*.xls*')
    input_files = {}
    for input_file in input_files_init:
        for company_head in companies:
            file_name = str(input_file).split('\\')[-1]
            if company_head.upper().strip() in file_name.upper():
                input_files[input_file] = company_head

    for input_file, company in input_files.items():
        try:
            progress = 0
            queue_num = 0
            request_queue.put(['update', [1, 0]])
            file_queue += 1
            file_name = str(input_file).split('\\')[-1]
            check_exit(exit_queue)

            request_queue.put(['print', 'Processing {}...'.format(file_name)])
            pay_cols = get_columns(spreadsheet_data, sheets['Pay'], pay_headers, 'Company', company)
            ded_cols = get_columns(spreadsheet_data, sheets['Deductions'], deduction_headers, 'Company', company)
            pay_data = get_data(pay_cols, input_file)
            ded_data = get_data(ded_cols, input_file)
            try:
                payroll_period = pay_data['Payroll Period']
            except Exception:
                raise Exception('Payroll Period not found! Please make sure the SHEET TAB name and payroll cell index is correct!')

            # remove separation header row
            for datas in [pay_data, ded_data]:
                for skip in ['skip this row', 'skip this row!']:
                    if skip in datas:
                        datas.pop(skip)

            # Get & update Messenger PSIDs
            if enable_sending:
                psids = {}
                to_find_psid = {}
                psid_rows = {}
                cnt = 1
                request_queue.put(['print', 'Getting {} Messenger IDs...'.format(company)])
                psid_data = get_sheet_values(sheets_service, PSID_sheet_id, '{}!A:D'.format(company))
                employee_names = [str(x).upper().replace('!', '').strip() for x in pay_data if x != 'Payroll Period']
                for psid_row in psid_data['values']:
                    if psid_row[0].upper() in employee_names:
                        if len(psid_row) == 3:
                            psids[psid_row[0].upper()] = psid_row[2]
                        elif len(psid_row) == 2:
                            to_find_psid[psid_row[0].upper()] = psid_row[1]
                            psid_rows[psid_row[0].upper()] = cnt
                    cnt += 1
                if 'EMPLOYEE NAME' in psids:
                    del psids['EMPLOYEE NAME']
                check_exit(exit_queue)
                if to_find_psid:
                    new_psids = get_psids(msgr, to_find_psid)
                else:
                    new_psids = {}
                check_exit(exit_queue)
                for key, new_psid in new_psids.items():
                    update_spreadsheet(sheets_service, [[new_psid]], PSID_sheet_id,
                                       '{}!C{}'.format(company, psid_rows[key]))
                psids.update(new_psids)

            request_queue.put(['print', 'Processing {}...'.format(file_name)])

            prompt_result = confirm_period(request_queue, period_queue, company, payroll_period, start_date, end_date)
            if not prompt_result:
                request_queue.put(['log', ['Sending for file {} for payroll period {} skipped by user.'.format(file_name, payroll_period), None]])
                add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), 'Sending for file {} for payroll period {} skipped by user.'.format(file_name, payroll_period))
                continue

        except Exception as e:
            traceback.print_exc()
            request_queue.put(['log', ['Error: {}'.format(str(e)), 'red']])
            request_queue.put(['edit', [1, 'Failed', file_queue, 2, 'red']])
            add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), str(e))
            continue

        for employee, employee_data in pay_data.items():
            fb_name = ''
            loop = True
            while loop:
                try:
                    check_exit(exit_queue)
                    messenger_sent = 'Not sent'
                    email_sent = 'Not sent'
                    if employee == 'Payroll Period':
                        loop = False
                        continue
                    queue_num += 1

                    calculated_gross = 0
                    data = deepcopy(pay_params)
                    for param_name, payments in employee_data.items():
                        if ' desc' not in param_name and param_name not in ['Gross', 'Net', 'Email', 'Branch', 'ID No.']:
                            try:
                                float_amount = float(payments)
                            except Exception:
                                raise Exception('Entry should be numeric under "{}" column!\nNon-numeric entry: {}\nPlease check column alignment!'.format(param_name, payments))
                            if '()' in param_name:
                                base_name = param_name.replace('()', '').strip()
                                str_amount = "({0:,.2f})".format(float_amount)
                                calculated_gross -= float_amount
                            else:
                                base_name = param_name.strip()
                                str_amount = "{0:,.2f}".format(float_amount)
                                if base_name not in ['Daily Rate', 'ECOLA']:
                                    calculated_gross += float_amount
                            if base_name in data:
                                data[base_name]['Amount'] = str_amount
                            else:
                                data[base_name] = {'Description': '', 'Amount': str_amount}
                        elif param_name not in ['Gross', 'Net', 'Email', 'Branch', 'ID No.']:
                            base_name = param_name.strip()[:-5]
                            duration_param = ''
                            try:
                                float_amount = float(payments)
                            except Exception:
                                raise Exception('Entry should be numeric under "{}" column!\nNon-numeric entry: {}\nPlease check column alignment!'.format(param_name, payments))
                            str_amount = "{0:,.2f}".format(float_amount)
                            if base_name in durations:
                                duration_param = durations[base_name]
                            if base_name in data:
                                data[base_name]['Description'] = '{} {}'.format(str_amount, duration_param)
                            else:
                                data[base_name] = {'Description': '{} {}'.format(str_amount, duration_param), 'Amount': ''}
                    param_list = [x for x in data]
                    desc_list = [x['Description'] for x in data.values()]
                    amt_list = [x['Amount'] for x in data.values()]
                    pay_table = {'Pay Type': param_list, 'Description': desc_list, 'Amount': amt_list}

                    param_list = []
                    amt_list = []
                    desc_list = []
                    if employee not in ded_data:
                        raise Exception(
                            'Deductions for {} not found! Please double check data file and payroll file.'.format(
                                employee))
                    calculated_deductions = 0
                    data = deepcopy(deduct_params)
                    for param_name, deductions in ded_data[employee].items():
                        if param_name in ['Total']:
                            loop = False
                            continue
                        if ' desc' in param_name:
                            base_name = param_name.strip()[:-4].strip()
                            try:
                                float_amount = float(deductions)
                            except Exception:
                                raise Exception('Entry should be numeric under "{}" column!\nNon-numeric entry: {}\nPlease check column alignment!'.format(param_name, deductions))
                            deductions = "({0:,.2f})".format(float_amount)
                            if base_name in data:
                                data[base_name]['Description'] = '{} {}'.format(deductions, durations[base_name])
                            else:
                                data[base_name] = {'Description': '{} {}'.format(deductions, durations[base_name]), 'Amount': ''}
                        elif param_name not in ['Total']:
                            try:
                                float_amount = float(deductions)
                            except Exception:
                                raise Exception('Entry should be numeric under "{}" column!\nNon-numeric entry: {}\nPlease check column alignment!'.format(param_name, deductions))
                            if '()' in param_name:
                                base_name = param_name.replace('()', '').strip()
                                str_amount = "({0:,.2f})".format(float_amount)
                                calculated_deductions -= float_amount
                            else:
                                base_name = param_name.strip()
                                str_amount = "{0:,.2f}".format(float_amount)
                                if base_name not in ['Daily Rate', 'ECOLA']:
                                    calculated_deductions += float_amount
                            if base_name in data:
                                data[base_name]['Amount'] = str_amount
                            else:
                                data[base_name] = {'Description': '', 'Amount': str_amount}
                    param_list = [x for x in data]
                    desc_list = [x['Description'] for x in data.values()]
                    amt_list = [x['Amount'] for x in data.values()]
                    deduct_table = {'Deduction Type': param_list, 'Description': desc_list, 'Amount': amt_list}

                    employee_name = str(employee).upper().replace('!', '').strip()
                    if '\\' in employee_name or '/' in employee_name:
                        employee_name = ' '.join([x for x in employee_name.split(' ') if '\\' not in x and '/' not in x]).strip()
                    if 'cut off' in employee_name:
                        employee_name = employee_name.split('cut off')[0].strip()
                    if ',' in employee_name:
                        name_split = employee_name.split(',')[1].strip().split()
                        first_name = ' '.join([x for x in name_split if x != name_split[-1] or ' ' not in name_split]).title()
                    employee_email = ''
                    if 'Email' in pay_data[employee]:
                        employee_email = pay_data[employee]['Email']
                    branch = ''
                    if 'Branch' in pay_data[employee]:
                        branch = pay_data[employee]['Branch']
                    id_no = ''
                    if 'ID No.' in pay_data[employee]:
                        id_no = pay_data[employee]['ID No.']

                    # Create folders for payslip
                    payslip_folder = '{}\\Payslips'.format(payroll_path)
                    if not os.path.exists(payslip_folder):
                        os.makedirs(payslip_folder)
                    company_folder = '{}\\{}'.format(payslip_folder, company)
                    if not os.path.exists(company_folder):
                        os.makedirs(company_folder)
                    payroll_period_folder = '{}\\{}'.format(company_folder, payroll_period)
                    if not os.path.exists(payroll_period_folder):
                        os.makedirs(payroll_period_folder)
                    output_file_name = '{}_{}.pdf'.format(payroll_period, employee_name)
                    payroll_period_text = payroll_period.capitalize().replace('-', ' to ')
                    if '!' in employee:
                        file_num = employee.count('!') + 1
                        output_file_name = '{}_{}_{}.pdf'.format(payroll_period, employee_name, file_num)

                    # Check calculations and create PDF
                    gross_pay = "{0:,.2f}".format(float(calculated_gross))
                    total_deductions = "{0:,.2f}".format(float(calculated_deductions))
                    calculated_net = "{0:,.2f}".format(float(calculated_gross - calculated_deductions))
                    try:
                        net_pay = "{0:,.2f}".format(float(pay_data[employee]['Net']))
                    except Exception:
                        raise Exception("{}'s 'Net' not found! Please check column alignment.".format(employee_name))
                    discrepancy = False
                    difference = "{0:,.2f}".format(abs(float(calculated_net.replace(',', '')) - float(net_pay.replace(',', ''))))
                    if calculated_net != net_pay and difference != '0.01':
                        # Revert gross and deductions to equal excel values.
                        try:
                            excel_gross = "{0:,.2f}".format(float(pay_data[employee]['Gross']))
                            excel_deductions = "{0:,.2f}".format(float(ded_data[employee]['Total']))
                        except Exception as not_found_e:
                            raise Exception("{}'s {} not found! Please check column alignment.".format(employee_name, str(not_found_e)))
                        discrepancy = True
                        warning_text = '{} employee {} with pay discrepancies. Please double check payslip PDF before sending manually.'.format(
                            company, employee_name)
                        request_queue.put(['log', [warning_text, 'yellow']])
                        add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), warning_text)

                        discrepancy_folder = '{}\\Discrepancies'.format(payroll_period_folder)
                        if not os.path.exists(discrepancy_folder):
                            os.makedirs(discrepancy_folder)
                        for file_prefix, sub_total in {'Tool Calculated_': [gross_pay, total_deductions, calculated_net], 'Excel Values_': [excel_gross, excel_deductions, net_pay]}.items():
                            output_file_path = '{}\\{}{}'.format(discrepancy_folder, file_prefix, output_file_name)
                            payslip_pdf(logo_file, title, id_no, branch,
                                        payroll_period_text, pay_table,
                                        deduct_table, employee_name, sub_total[0], sub_total[1], sub_total[2],
                                        output_file_path)
                    else:
                        output_file_path = '{}\\{}'.format(payroll_period_folder, output_file_name)
                        payslip_pdf(logo_file, title, id_no, branch,
                                    payroll_period_text, pay_table,
                                    deduct_table, employee_name, gross_pay, total_deductions, net_pay,
                                    output_file_path)

                    if enable_sending:
                        # send pdf file via email
                        email_result = None
                        if employee_email != '':
                            file_attachments = [output_file_path]
                            to = employee_email
                            cc = sender_email
                            body = ('Hi {},\n\nKindly see attached payslip for the payroll period {}.\n\n'
                                    'Please examine your payslip immediately upon receipt.\n'
                                    'If no error is reported within 7 days, the account is considered accurate.\n'
                                    'Your payslip is system generated. Signature is not necessary unless altered.\n\n\n'
                                    'regards,\n{}').format(first_name, payroll_period, company_name)
                            subject = 'Payslip: {}'.format(payroll_period)
                            send_email(sender_email, to, subject, body, file_attachments, cc)
                            email_sent = 'Sent'

                        # send pdf file via messenger
                        if discrepancy:
                            pass
                        elif employee_name in psids:
                            fb_name = psids[employee_name]
                            msg = ('Hi {}, here is your payslip for the payroll period {}.\n\n'
                                   'Please examine your payslip immediately upon receipt.\n'
                                   'If no error is reported within 7 days, the account is considered accurate.\n'
                                   'Your payslip is system generated. Signature is not necessary unless altered.').format(first_name, payroll_period)
                            send_message(msgr, psids[employee_name], msg)
                            msgr_result = send_attachment(msgr, psids[employee_name], output_file_path)
                            if msgr_result.status_code == 200:
                                messenger_sent = 'Sent'
                            else:
                                raise Exception('Failed to send payslip of {} via Messenger!'.format(employee_name))
                        else:
                            if employee_name in to_find_psid:
                                fb_name = to_find_psid[employee_name]
                                raise Exception("{}'s Messenger conversation not found. Please check if the FB account name ({}) is correct. If correct, kindly request employee to chat the page and try again.".format(employee_name, to_find_psid[employee_name]))
                            else:
                                warning_text = "{}'s Messenger name not listed in data file. Please input employee data in 'Employee Messenger Conversation IDs' spreadsheet.".format(employee_name)
                                request_queue.put(['log', [warning_text, 'yellow']])
                                add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), warning_text)

                except Exception as e1:
                    traceback.print_exc()
                    request_queue.put(['log', ['Error: {}'.format(str(e1)), 'red']])
                    add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), str(e1))

                employee_logs = [company, payroll_period, employee_name, employee_email, email_sent, fb_name, messenger_sent, d.datetime.strftime(d.datetime.now(), '%Y-%m-%d %H:%M:%S')]
                employee_logs_file = '{}\\{}'.format(rec_log_path, log_sheet_name)
                record_employee(employee_logs_file, employee_logs, sheet_header)

                check_prog = queue_num / (len(pay_data) - 1)
                if progress + 0.02 < check_prog:
                    progress = check_prog
                    request_queue.put(['update', [1, progress]])
                break

        # move to finished folder
        request_queue.put(['print', 'Finishing file...'])
        request_queue.put(['edit', [1, 'Done', file_queue, 2, 'green']])
        request_queue.put(['update', [2, file_queue / file_num]])
        request_queue.put(['log', ['{} Processed'.format(file_name), None]])
        add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), '{} Processed'.format(file_name))
        try:
            os.replace(input_file, '{}\\{}'.format(finished_path, file_name))
        except Exception:
            warning_text = 'Warning: {} not moved to "Finished" folder. Make sure the file is not open next time.'.format(file_name)
            request_queue.put(['log', [warning_text, 'yellow']])
            add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), warning_text)
        gc.collect()

    if file_queue == 0:
        request_queue.put(['log', ['No files found in "Files" folder.', 'red']])
        add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), 'No files found in "Files" folder.')

    date_now = d.datetime.strftime(d.datetime.now(), '%Y-%m-%d')
    add_logs(log_path + date_now, 'Tool finished.')
    try:
        if employee_logs_file is not None:
            create_file(drive_service, log_sheet_name, 'text', year_folder_id, employee_logs_file)
        create_file(drive_service, date_now, 'text', tool_logs_folder_id, log_path + date_now + '.txt')
    except Exception:
        warning_text = 'Warning: Logs not uploaded in Google Drive.'
        request_queue.put(['log', [warning_text, 'yellow']])
        add_logs(log_path + d.datetime.strftime(d.datetime.now(), '%Y-%m-%d'), warning_text)
    request_queue.put(['done', None])
    print('FINISHED!')
