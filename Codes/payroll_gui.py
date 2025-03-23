# This uses open-source Custom Tkinter by Tom Schimansky
# https://github.com/TomSchimansky/CustomTkinter

from constant import *
import dateutil.relativedelta
import distribute_payslips as p
import customtkinter as ctk
import ctk_xyframe as xyframe
import datetime as d
import tkinter as tk
import PIL.Image
import tkcalendar
import traceback
import threading
import ctktable
import calendar
import queue
import gc
import os


def get_length(pos, border=0.04):
    return 1 - pos - border


def change_pos(obj, dimensions):
    obj.place(relx=dimensions[0], rely=dimensions[1], relwidth=dimensions[2], relheight=dimensions[3], anchor="nw")


def hide_company_frame(company_frame):
    global company_frame_shown
    if company_frame_shown:
        company_frame.place_forget()
        change_pos(line1, line1_dimensions)
        hide_frame_btn1.configure(image=down_arrow)
        if employee_frame_shown:
            change_pos(logs_frame, ef_dimensions['full'])
        else:
            change_pos(line2, line2_dimensions['top'])
        change_pos(hide_frame_btn2, hf2_dimensions['top'])
        change_pos(total_label, el_dimensions['top'])
        change_pos(total_progress, ep_dimensions['top'])
        company_frame_shown = False
    else:
        line1.place_forget()
        hide_frame_btn1.configure(image=up_arrow)
        if employee_frame_shown:
            change_pos(company_frame, cf_dimensions['half'])
            change_pos(logs_frame, ef_dimensions['half'])
            change_pos(hide_frame_btn2, hf2_dimensions['mid'])
            change_pos(total_label, el_dimensions['mid'])
            change_pos(total_progress, ep_dimensions['mid'])
        else:
            change_pos(company_frame, cf_dimensions['full'])
            change_pos(line2, line2_dimensions['bottom'])
            change_pos(hide_frame_btn2, hf2_dimensions['bottom'])
            change_pos(total_label, el_dimensions['bottom'])
            change_pos(total_progress, ep_dimensions['bottom'])
        company_frame_shown = True


def hide_employee_frame(logs_frame):
    global employee_frame_shown
    if employee_frame_shown:
        logs_frame.place_forget()
        hide_frame_btn2.configure(image=down_arrow)
        if company_frame_shown:
            change_pos(company_frame, cf_dimensions['full'])
            change_pos(line2, line2_dimensions['bottom'])
            change_pos(hide_frame_btn2, hf2_dimensions['bottom'])
            change_pos(total_label, el_dimensions['bottom'])
            change_pos(total_progress, ep_dimensions['bottom'])
        else:
            change_pos(line2, line2_dimensions['top'])
            change_pos(hide_frame_btn2, hf2_dimensions['top'])
            change_pos(total_label, el_dimensions['top'])
            change_pos(total_progress, ep_dimensions['top'])
        employee_frame_shown = False
    else:
        line2.place_forget()
        hide_frame_btn2.configure(image=up_arrow)
        if company_frame_shown:
            change_pos(company_frame, cf_dimensions['half'])
            change_pos(logs_frame, ef_dimensions['half'])
            change_pos(hide_frame_btn2, hf2_dimensions['mid'])
            change_pos(total_label, el_dimensions['mid'])
            change_pos(total_progress, ep_dimensions['mid'])
        else:
            change_pos(logs_frame, ef_dimensions['full'])
            change_pos(hide_frame_btn2, hf2_dimensions['top'])
            change_pos(total_label, el_dimensions['top'])
            change_pos(total_progress, ep_dimensions['top'])
        employee_frame_shown = True


def add_gui_table_row(table_index: int, values: list):
    if table_index == 1:
        table1.add_row(values)
    elif table_index == 2:
        # table2.add_row(values)
        pass


def edit_gui_cell(table_index: int, value, edit_row, edit_col, color='black'):
    if table_index == 1:
        table1.edit(row=edit_row, column=edit_col, value=value, text_color=color, font=('Helvetica', 12, 'bold'))
    elif table_index == 2:
        # table2.edit(row=edit_row, column=edit_col, value=value, text_color=color, font=('Helvetica', 12, 'bold'))
        pass


# noinspection PyUnusedLocal
def update_value(s):
    global period, start_date_str, end_date_str, start_date, end_date
    start_date = cal.get_date()
    start_date_str = d.datetime.strftime(start_date, '%B %d, %Y')
    end_date = get_end_date(start_date)
    end_date_str = d.datetime.strftime(end_date, '%B %d, %Y')
    cal2.set_date(end_date)
    period.configure(text='{}\n- {}'.format(start_date_str, end_date_str))


def update_value2(s):
    global period, start_date_str, end_date_str, end_date
    end_date = cal2.get_date()
    end_date_str = d.datetime.strftime(end_date, '%B %d, %Y')
    period.configure(text='{}\n- {}'.format(start_date_str, end_date_str))


def toggle_dark_mode(mode):
    ctk.set_appearance_mode(mode)
    update_calendar_appearance()
    f = open(payroll_path + '\\Data\\preferred_mode', 'w')
    f.write(mode)
    f.close()


def update_calendar_appearance():
    for c in [cal, cal2]:
        if ctk.get_appearance_mode() == 'Light':
            c.configure(background='#cce4f7', headersbackground='#cfcfcf', normalbackground='white',
                        weekendbackground='white', othermonthbackground='#ebebeb', othermonthwebackground='#ebebeb',
                        foreground='black', bordercolor='#ebebeb', headersforeground='gray', normalforeground='black')
        elif ctk.get_appearance_mode() == 'Dark':
            c.configure(background='#2b2b2b', headersbackground='#242424', normalbackground='#4a4d50',
                        weekendbackground='#4a4d50', othermonthbackground='#2b2b2b',
                        othermonthwebackground='#2b2b2b', foreground='white',
                        bordercolor='black', headersforeground='gray', normalforeground='white',
                        weekendforeground='#aaaaaa')


def get_end_date(start_date):
    if start_date.day < 15 or start_date.day > 18:
        return start_date + d.timedelta(days=14)
    else:
        day = calendar.monthrange(start_date.year, start_date.month)[1]
        return d.date(start_date.year, start_date.month, day)


def update_progress(index, progress_perc):
    if index == 1:
        company_progress.set(progress_perc)
    elif index == 2:
        total_progress.set(progress_perc)


def add_gui_log(text, color=None):
    time_now = d.datetime.now().strftime("%H:%M:%S")
    text = '{}: {}'.format(time_now, text)
    if color is None:
        log = ctk.CTkLabel(logs_frame, text=text, font=('Helvetica', 11), height=15)
    else:
        log = ctk.CTkLabel(logs_frame, text=text, font=('Helvetica', 11), text_color=color, height=15)
    log.grid(padx=(0, 0), pady=(0, 0), sticky="w")


def set_sending_mode(sending_var):
    global enable_sending
    enable_sending = sending_var == "Sending Enabled"
    f = open(payroll_path + '\\Data\\enable_sending', 'w')
    f.write(sending_var)
    f.close()


def set_payslip_type(payslip_type_var):
    global payslip_type
    payslip_type = payslip_type_var
    f = open(payroll_path + '\\Data\\payslip_type', 'w')
    f.write(payslip_type)
    f.close()


def on_closing():
    global exit_queue
    if t.is_alive():
        if tk.messagebox.askokcancel('Payslip Distributor', 'The tool is still running.\n'
                                                                 'Are you sure you want to stop and quit?'):
            exit_queue.put('exit')
            gui.destroy()
    else:
        gui.destroy()
        exit()


def run_main():
    global request_queue, t, loading_bar
    if not t.is_alive():
        if tk.messagebox.askokcancel('Payslip Distributor', 'Distribute payslip for payroll scope: {} - {}?'.format(start_date_str, end_date_str)):
            t.start()
            check_queue(request_queue)
            loading_bar.grid(row=4, column=0, padx=(20, 10), pady=(0, 0), sticky="ew")
            loading_bar.start()
            update_progress(1, 0)
            update_progress(2, 0)
            create_table()
            add_gui_log('Tool started.', None)
    else:
        tk.messagebox.showwarning('Payslip Distributor', 'A process is still currently running!')


def create_table():
    global table1
    table1.pack_forget()
    table1.destroy()
    header = [['Queue', '                    Payroll Excel File                    ', 'Status']]
    table1 = ctktable.CTkTable(company_frame, row=1, column=3, values=header, width=80, height=22, padx=0)
    table1.pack(expand=True, fill="both", padx=0, pady=0)
    for i in range(0, 3):
        table1.edit(row=0, column=i, font=('Helvetica', 12, 'bold'))


def check_queue(request_queue):
    global t, print_label
    try:
        queue_result = request_queue.get_nowait()
    except queue.Empty:
        pass
        if not t.is_alive():
            loading_bar.grid_forget()
            tk.messagebox.showerror('Payslip Distributor', 'An unexpected error has occured!')
            t = threading.Thread(target=lambda: p.main(start_date, end_date, request_queue, period_queue, exit_queue, enable_sending, payslip_type))
            gc.collect()
            return
    else:
        to_do, params = queue_result
        if to_do == 'add':
            table_index, values = params
            add_gui_table_row(table_index, values)
        elif to_do == 'update':
            index, progress_perc = params
            update_progress(index, progress_perc)
        elif to_do == 'edit':
            table_index, value, edit_row, edit_col, color = params
            edit_gui_cell(table_index, value, edit_row, edit_col, color)
        elif to_do == 'print':
            print_label.configure(text=params)
        elif to_do == 'log':
            text, color = params
            add_gui_log(text, color)
        elif to_do == 'done':
            print_label.configure(text='')
            loading_bar.grid_forget()
            add_gui_log('Finished.')
            update_progress(2, 1)
            t = threading.Thread(target=lambda: p.main(start_date, end_date, request_queue, period_queue, exit_queue, enable_sending, payslip_type))
            tk.messagebox.showinfo('Payslip Distributor', 'Process has finished!\nKindly double check using logs.')
            gc.collect()
            return
        elif to_do == 'prompt':
            prompt_response = tk.messagebox.askyesno('Payslip Distributor', params)
            period_queue.put(prompt_response)
        gc.collect()
    gui.after(1, lambda: check_queue(request_queue))


try:
    # GLOBAL Variables
    gui = ctk.CTk()
    period = None
    start_date_str = ''
    end_date_str = ''
    period_queue = queue.Queue()
    exit_queue = queue.Queue()
    request_queue = queue.Queue()
    t = threading.Thread(target=lambda: p.main(start_date, end_date, request_queue, period_queue, exit_queue, enable_sending, payslip_type))

    # get preferred mode
    data_folder = payroll_path + '\\Data\\'
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)
    mode_path = data_folder + 'preferred_mode'
    if os.path.exists(mode_path):
        f = open(mode_path, 'r')
        mode = f.read()
    else:
        mode = 'light'

    # get sending preference
    mode_path = payroll_path + '\\Data\\enable_sending'
    if os.path.exists(mode_path):
        f = open(mode_path, 'r')
        sending_mode = f.read()
    else:
        sending_mode = 'Create PDF Only'
    enable_sending = sending_mode == 'Sending Enabled'

    # get payslip type
    mode_path = payroll_path + '\\Data\\payslip_type'
    if os.path.exists(mode_path):
        f = open(mode_path, 'r')
        payslip_type = f.read()
    else:
        payslip_type = "Standard"

    # set appearance
    ctk.set_appearance_mode(mode)  # Modes: "system" (standard), "dark", "light"
    ctk.set_default_color_theme('green')  # Themes: "blue" (standard), "green", "dark-blue"

    # prepare gui window
    gui.title("Payslip Distributor")
    gui.protocol("WM_DELETE_WINDOW", on_closing)

    # set up window dimensions and position
    gui_w = 850
    gui_h = 650
    screen_w = gui.winfo_screenwidth()
    screen_h = gui.winfo_screenheight()
    window_x = (screen_w - gui_w) / 2
    window_y = (screen_h - gui_h) / 2
    gui.geometry('%dx%d+%d+%d' % (gui_w, gui_h, window_x, window_y))
    data_path = resource_path + '\\Data\\'
    image_path = data_path + '\\Images\\'
    if not os.path.exists(image_path):
        os.makedirs(payroll_path + '\\Data\\')
        os.makedirs(payroll_path + '\\Data\\Images\\')
        image_path = payroll_path + '\\Data\\Images\\'
    company_frame_shown = True
    employee_frame_shown = True

    # window icon
    icon_file = image_path + 'icon.png'
    icon = PIL.ImageTk.PhotoImage(file=icon_file)
    gui.after(201, lambda: gui.iconphoto(False, icon))

    # Add header logo
    light_logo_file = image_path + 'logo_light.png'
    light_logo = PIL.Image.open(light_logo_file)
    dark_logo_file = image_path + 'logo_dark.png'
    dark_logo = PIL.Image.open(dark_logo_file)
    logo = ctk.CTkImage(light_image=light_logo, dark_image=dark_logo, size=(240, 60))
    header_logo = ctk.CTkButton(gui, image=logo, text='', fg_color='transparent', bg_color='transparent', hover=False, corner_radius=0)
    header_logo.place(relx=.0001, rely=.02, relwidth=.4, relheight=.1, anchor="nw")

    # LEFT FRAME DIMENSIONS
    rely = 0.15
    border = 0.04
    left_frame_h = get_length(rely, border)
    left_frame = xyframe.CTkXYFrame(gui)
    change_pos(left_frame, [border, rely, .35, left_frame_h])

    # enable/disable sending
    sending_var = ctk.StringVar(value=sending_mode)
    toggle_sending = ctk.CTkOptionMenu(left_frame, values=["Sending Enabled", "Create PDF Only"], variable=sending_var,
                                       fg_color='#eeeeee',
                                       command=set_sending_mode, width=20, height=15, text_color='#555555',
                                       font=('helvetica', 11, 'bold'))
    toggle_sending.grid(row=0, column=0, padx=(60, 50), pady=(20, 0), sticky="nsew")
    payslip_type_var = ctk.StringVar(value=payslip_type)
    toggle_payslip_type = ctk.CTkOptionMenu(left_frame, values=["Standard", "Monthly"], variable=payslip_type_var,
                                            fg_color='#eeeeee',
                                            command=set_payslip_type, width=20, height=15, text_color='#555555',
                                            font=('helvetica', 11, 'bold'))
    toggle_payslip_type.grid(row=1, column=0, padx=(60, 50), pady=(10, 0), sticky="nsew")

    # CALENDAR FRAME
    calendar_frame = ctk.CTkFrame(left_frame, border_width=2)
    calendar_frame.grid(row=2, column=0, padx=(13, 0), pady=(20, 1), sticky="nsew")

    # Get start date
    start_date = d.date.today()
    day_today = start_date.day
    if 8 < day_today < 24:
        day = 1
    elif day_today < 9:
        start_date = start_date - dateutil.relativedelta.relativedelta(months=1)
    if day_today < 9 or day_today > 23:
        day = 16
    start_date = d.date(start_date.year, start_date.month, day)
    end_date = get_end_date(start_date)

    start_date_label = ctk.CTkLabel(calendar_frame, corner_radius=0, text='Start date:', height=10, font=('helvetica', 10))
    start_date_label.grid(row=0, column=0, padx=(50, 10), pady=(10, 0), sticky="w")
    sel = ctk.StringVar(gui)

    start_date_label = ctk.CTkLabel(calendar_frame, corner_radius=0, text='Start date:', width=10, height=10, font=('helvetica', 10))
    start_date_label.grid(row=0, column=0, padx=(50, 0), pady=(10, 0), sticky="nw")
    cal = tkcalendar.DateEntry(calendar_frame, selectmode='day', textvariable=sel, date_pattern='y-mm-dd', width=12, height=10, year=start_date.year, month=start_date.month, day=day)
    cal.grid(row=1, column=0, padx=(28, 10), pady=(4, 20), sticky="w")

    end_date_label = ctk.CTkLabel(calendar_frame, corner_radius=0, text='End date:', height=10, font=('helvetica', 10))
    end_date_label.grid(row=0, column=1, padx=(13, 10), pady=(10, 0), sticky="w")
    sel2 = ctk.StringVar(gui)
    cal2 = tkcalendar.DateEntry(calendar_frame, selectmode='day', textvariable=sel2, date_pattern='y-mm-dd', width=12, height=10, year=start_date.year, month=start_date.month, day=day)
    cal2.grid(row=1, column=1, padx=(0, 28), pady=(4, 20), sticky="nsew")
    update_calendar_appearance()

    period_label = ctk.CTkLabel(calendar_frame, corner_radius=0, text='Payroll scope:', height=10, font=('helvetica', 11))
    period_label.grid(row=2, column=0, padx=(30, 10), pady=(10, 0), sticky="nw", columnspan=2)

    period = ctk.CTkLabel(calendar_frame, corner_radius=70, bg_color='white', text='', height=40, width=170, font=('helvetica', 13, 'bold'), text_color='gray')
    period.grid(row=3, column=0, padx=(28, 28), pady=(5, 15), sticky="nsew", columnspan=2)
    update_value('<<DateEntrySelected>>')
    cal.bind('<<DateEntrySelected>>', update_value)
    cal2.bind('<<DateEntrySelected>>', update_value2)

    # RUN BUTTON
    print_label = ctk.CTkLabel(left_frame, corner_radius=0, text='', width=210, height=12,
                               font=('helvetica', 10))
    print_label.grid(row=3, column=0, padx=(20, 10), pady=(10, 0), sticky="nw")
    loading_bar = ctk.CTkProgressBar(left_frame, mode="indeterminate", progress_color='#5bda7f')
    run_btn = ctk.CTkButton(left_frame, text='START', text_color='gray', font=('helvetica', 13, 'bold'),
                            fg_color="transparent", hover_color='#cfcfcf', border_color='#939ba2', border_width=3,
                            corner_radius=30, command=run_main, width=30, height=50)

    run_btn.grid(row=5, column=0, padx=(60, 60), pady=(30, 10), sticky="nsew")

    # Dark mode toggle
    mode_frame = ctk.CTkFrame(left_frame, width=100, height=50)
    mode_frame.grid(row=6, column=0, padx=(70, 70), pady=(65, 20), sticky="nsew")
    mode_var = ctk.StringVar(value=mode)
    dark_toggle = ctk.CTkSwitch(mode_frame, text='', onvalue='light', offvalue='dark', variable=mode_var, command=lambda: toggle_dark_mode(mode_var.get()))
    change_pos(dark_toggle, [.1, .09, 1, .8])
    bulb_on_file = image_path + 'bulb_on.png'
    bulb_off_file = image_path + 'bulb_off.png'
    bulb_on_image = PIL.Image.open(bulb_on_file)
    bulb_off_image = PIL.Image.open(bulb_off_file)
    bulb = ctk.CTkImage(light_image=bulb_on_image, dark_image=bulb_off_image, size=(35, 31))
    bulb_label = ctk.CTkButton(mode_frame, image=bulb, text='', fg_color='transparent', bg_color='transparent', hover=False, corner_radius=0)
    change_pos(bulb_label, [.5, .001, .4, 1])

    # RIGHT FRAME DIMENSIONS
    progress_x = 0.6
    label_height = 0.03
    progress_w = get_length(progress_x, border)
    progress_h = 0.02
    relx = 0.42
    company_y = 0.1
    employee_y = 0.57
    rel_width = get_length(relx, border)
    rel_height = get_length(company_y, border)
    rel_height = (rel_height/2) - border
    rel_height_full = rel_height * 2
    line_thickness = 0.003

    cf_dimensions = {'half': [relx, company_y, rel_width, rel_height], 'full': [relx, company_y, rel_width, rel_height_full]}
    ef_dimensions = {'half': [relx, employee_y, rel_width, rel_height], 'full': [relx, 0.17, rel_width, rel_height_full]}

    hf1_dimensions = [relx, .05, 0.041, .040]
    hf2_dimensions = {'bottom': [relx, .91, 0.041, .040], 'mid': [relx, .52, 0.041, .040], 'top': [relx, .12, 0.041, .040]}
    el_dimensions = {'bottom': [.465, .915, .13, label_height], 'mid': [.465, .525, .13, label_height], 'top': [.465, .125, .13, label_height]}
    cl_dimensions = [.475, .055, .11, label_height]
    ep_dimensions = {'bottom': [progress_x, .92, progress_w, progress_h], 'mid': [progress_x, .53, progress_w, progress_h], 'top': [progress_x, .13, progress_w, progress_h]}
    cp_dimensions = [progress_x, .06, progress_w, .02]
    line2_dimensions = {'bottom': [relx, .96, rel_width, line_thickness], 'mid': [relx, employee_y, rel_width, line_thickness], 'top': [relx, .17, rel_width, line_thickness]}
    line1_dimensions = [relx, company_y, rel_width, 0.003]

    # top right progress bar
    progress_x = 0.6
    label_height = 0.03
    progress_w = get_length(progress_x, border)

    up_arrow_file = image_path + 'up_arrow.png'
    up_arrow_img = PIL.Image.open(up_arrow_file)
    down_arrow_file = image_path + 'down_arrow.png'
    down_arrow_img = PIL.Image.open(down_arrow_file)
    up_arrow = ctk.CTkImage(light_image=up_arrow_img, dark_image=up_arrow_img, size=(13, 9))
    down_arrow = ctk.CTkImage(light_image=down_arrow_img, dark_image=down_arrow_img, size=(13, 9))

    # company frame
    relx = 0.42
    company_y = 0.1
    rel_width = get_length(relx, border)
    rel_height = get_length(company_y, border)
    rel_height = (rel_height/2) - border

    hide_frame_btn1 = ctk.CTkButton(gui, text='', image=up_arrow, fg_color="transparent", hover_color='#cfcfcf', border_color='#939ba2', border_width=2, command=lambda: hide_company_frame(company_frame))
    change_pos(hide_frame_btn1, hf1_dimensions)

    company_label = ctk.CTkLabel(gui, text='File progress:', font=('Helvetica', 12, 'bold'), text_color='gray')
    change_pos(company_label, cl_dimensions)
    company_progress = ctk.CTkProgressBar(gui, progress_color='#5bda7f')
    company_progress.set(0)
    change_pos(company_progress, cp_dimensions)

    line1 = xyframe.CTkXYFrame(gui)
    change_pos(line1, line1_dimensions)
    line1.place_forget()

    company_frame = xyframe.CTkXYFrame(gui)
    change_pos(company_frame, cf_dimensions['half'])
    header = [['Queue', '                    Payroll Excel File                    ', 'Status']]
    table1 = ctktable.CTkTable(company_frame, row=1, column=3, values=header, width=80, height=22, padx=0)
    table1.pack(expand=True, fill="both", padx=0, pady=0)
    for i in range(0, 3):
        table1.edit(row=0, column=i, font=('Helvetica', 12, 'bold'))

    # EMPLOYEE FRAME
    hide_frame_btn2 = ctk.CTkButton(gui, text='', image=up_arrow, fg_color="transparent", hover_color='#cfcfcf', border_color='#939ba2', border_width=2, command=lambda: hide_employee_frame(logs_frame))
    change_pos(hide_frame_btn2, hf2_dimensions['mid'])

    total_label = ctk.CTkLabel(gui, text='Total progress:', font=('Helvetica', 12, 'bold'), text_color='gray')
    total_label.place(relx=.465, rely=.525, relwidth=.13, relheight=label_height, anchor="nw")
    total_progress = ctk.CTkProgressBar(gui, progress_color='#5bda7f')
    total_progress.set(0)
    total_progress.place(relx=progress_x, rely=.53, relwidth=progress_w, relheight=.02, anchor="nw")

    line2 = xyframe.CTkXYFrame(gui)
    line2.place(relx=relx, rely=employee_y, relwidth=rel_width, relheight=0.003, anchor="nw")
    line2.place_forget()

    logs_frame = xyframe.CTkXYFrame(gui)
    change_pos(logs_frame, ef_dimensions['half'])

    gui.mainloop()

except Exception as e:
    err = str(e)
    traceback.print_exc()
    loading_bar.grid_forget()
    tk.messagebox.showerror('Payslip Distributor', 'An unexpected error has occured!')
    t = threading.Thread(target=lambda: p.main(start_date, end_date, request_queue, period_queue, exit_queue, enable_sending, payslip_type))
    gc.collect()
