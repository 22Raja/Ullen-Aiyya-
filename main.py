from external import *
from PIL import Image, ImageTk
from tkinter import ttk
from tkinter import filedialog



# file.close()
# sheet.sheet_view.showGridLines = False

# sheet.sheet_view.zoomScale = 80

# out_work_book.close()

# sheet.sheet_view.showGridLines = False

# sheet.sheet_view.zoomScale = 80

# out_work_book.close()

# It is used to delete a single record.

# To take action on the DATA WE provide in


# this is the main window to create the choose in th tkinter window
def export(event, class_name, institute_name, export_win):
    copy_file = ''

    def process(eve, year_get, month_get, day_get):
        if eve:
            print(year_get, month_get, day_get)
            file = filedialog.asksaveasfilename(initialfile=month_get, defaultextension='.xlsx',
                                                filetypes=["EXCEL file"], title='Ullen Aiyya')
            if file is None:
                return
            connection = os.path.join(last_file, month_get)
            wb = load_workbook(connection)
            print(day_get)
            #sheet = wb[day_get]
            #wb.copy_worksheet(sheet)
            wb.save(connection)

    if event.char == 'e' or event.char == '??':
        if 'Select Institution' in institute_name and 'Select Class' in class_name:
            messagebox.showinfo(title="Remainder", message='Choose Institution and Class')
        elif 'Select Class' in class_name:
            messagebox.showinfo(title="Remainder", message='Choose the Class')
        elif 'Select Institution' in institute_name:
            messagebox.showinfo(title="Remainder", message='Choose the Institution')
        else:
            class_db = os.path.join(section_folder, f'{class_name} {institute_name}.db')
            institute_db = os.path.join(class_folder, f'{institute_name}.db')
            get_institute = os.path.join(excel_folder, institute_name)
            get_class = os.path.join(get_institute, class_name)
            if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(institute_db):
                if os.path.exists(get_class):
                    export_win.destroy()
                    exp = Tk()
                    window(exp)
                    image = Image.open(background_image)
                    imageview = ImageTk.PhotoImage(image)
                    Label(exp, image=imageview).pack()
                    day_var = StringVar()
                    month_var = StringVar()
                    year_var = StringVar()
                    last_file = ''

                    def month_check(eve, year_get):
                        if eve:
                            if year_get != '-- Select YEAR --':
                                nonlocal last_file
                                last_file = os.path.join(get_class, year_get)
                                month_box['values'] = os.listdir(last_file)
                                month_box.config(state='readonly')

                    def day_check(eve, month_get):
                        if eve:
                            if month_get != '-- Select MONTH --':
                                connection = os.path.join(last_file, month_get)
                                while True:
                                    try:
                                        if os.path.exists(connection):
                                            wb = load_workbook(connection)
                                            nonlocal copy_file
                                            copy_file = wb.sheetnames
                                            day_box['values'] = copy_file
                                            wb.save(connection)
                                    except PermissionError:
                                        messagebox.showinfo("Reminder",
                                                            'Kindly close your opened excel file')
                                        continue
                                    except KeyError:
                                        break
                                    else:
                                        break
                                day_box.config(state='readonly')

                    day_radio = Radiobutton(text='DAY')
                    month_radio = Radiobutton(text='MONTH')
                    year_radio = Radiobutton(text='YEAR')
                    day_radio.place(x=30, y=100)
                    month_radio.place(x=30, y=110)
                    year_radio.place(x=30, y=120)

                    year_box = ttk.Combobox(exp, textvariable=year_var, state='readonly', values=os.listdir(get_class),
                                            width=54, cursor="hand2", justify=CENTER,
                                            font=('times', 10))
                    year_box.set('-- Select YEAR --')
                    year_box.bind("<<ComboboxSelected>>", lambda eve: month_check(eve, year_var.get()))
                    year_box.place(x=350, y=242)
                    month_box = ttk.Combobox(exp, textvariable=month_var, state='disable',
                                             width=54, cursor="hand2", justify=CENTER, font=('times', 10))
                    month_box.set('-- Select MONTH --')
                    month_box.bind("<<ComboboxSelected>>", lambda eve: day_check(eve, month_var.get()))
                    month_box.place(x=350, y=282)
                    day_box = ttk.Combobox(exp, textvariable=day_var, state='disable', font=('times', 10), width=54,
                                           justify=CENTER)
                    day_box.set('-- Select DAY --')
                    day_box.place(x=350, y=324)

                    def close(eve):
                        if eve.char == '`' or eve.char == '??':
                            exp.destroy()
                            system()

                    mark = Button(exp, text='BACK', fg='#FFFFFF',
                                  bg="#FF0051", relief=SOLID,
                                  cursor='hand2')
                    mark.bind('<Button>', close)
                    mark.place(x=0, y=23)
                    proceed = Button(exp, text="PROCEED", bg='#67E1E4', fg='#FFFFFF', activebackground='#67E1E4',
                                     activeforeground='#FFFFFF', relief=SOLID, width=10, font='BOLD')
                    proceed.bind('<Button>', lambda eve: process(eve, year_var.get(), month_var.get(), day_var.get()))
                    proceed.place(x=400, y=450)
                    exp.mainloop()
                else:
                    messagebox.showinfo(title="Remainder", message='No Excel there for this class')
            else:
                final_check(class_name, class_db, export_win, institute_db, institute_name)
                system()


# The Entry window

def first_window():
    def class_get(event):
        if event:
            inst_name = institute_var.get()
            institution_db = os.path.join(class_folder, f'{inst_name}.db')
            if os.path.exists(institution_db) and os.path.exists(college_dbpath):
                class_data = []
                total_check()
                connect1 = sqlite3.connect(institution_db)
                pen = connect1.cursor()
                pen.execute("SELECT * FROM classes")
                for d in pen.fetchall():
                    for z in d:
                        class_data.append(z)
                connect1.commit()
                connect1.close()
                if len(class_data) == 0:
                    messagebox.showinfo(title="Remainder", message='No class found to take Attendance')
                    class_box.set(constant_class)
                    class_box.configure(state='disable')
                    mark.place_forget()
                else:
                    class_box.set(constant_class)
                    ttk.Style().configure('TCombobox', foreground='#666666')
                    class_box['values'] = class_data
                    class_box.configure(state='readonly')
                    mark.place(x=635, y=370)
            else:
                board_deleted(institution_db, inst_name)
                main.destroy()
                system()

    institute_data = []
    sections = []
    main = Tk()
    window(main)
    image_view = PhotoImage(file=background_image)
    Label(main, image=image_view).pack()
    institute_var = StringVar()
    class_var = StringVar()
    connect = sqlite3.connect(college_dbpath)
    con = connect.cursor()
    con.execute("SELECT * FROM COLLEGE")
    for s in con.fetchall():
        for x in s:
            institute_data.append(x)
    connect.commit()
    connect.close()
    from institute import institution, excel_color, excel_class, excel_institute
    ttk.Style().configure('TCombobox', foreground=excel_color)
    institute_box = ttk.Combobox(main, textvariable=institute_var, values=institute_data, state='readonly',
                                 width=54, cursor="hand2", justify=CENTER,
                                 font=('times', 10))
    if excel_institute not in institute_data:
        excel_institute = constant
    institute_box.set(excel_institute)
    institute_box.bind("<<ComboboxSelected>>", class_get)
    institute_box.place(x=550, y=269)
    institute_db = os.path.join(class_folder, f'{excel_institute}.db')
    if os.path.exists(institute_db):
        connect5 = sqlite3.connect(institute_db)
        pen3 = connect5.cursor()
        pen3.execute("SELECT * FROM classes")
        for s in pen3.fetchall():
            for x in s:
                sections.append(x)
        connect5.commit()
        connect5.close()

    def color_get(k):
        if k:
            ttk.Style().configure('TCombobox', foreground='black')

    class_box = ttk.Combobox(main, textvariable=class_var, state='readonly', font=('times', 10), width=54,
                             justify=CENTER, values=sections)
    if excel_class not in sections:
        excel_class = constant_class

    class_box.set(excel_class)
    class_box.bind("<<ComboboxSelected>>", color_get)
    class_box.place(x=550, y=306)
    institution.block = main
    entry_institute = Button(main, text="MY INSTITUTION", activebackground='#67E1E4', relief=FLAT, borderwidth=0,
                             background='red',
                             fg='white', activeforeground='#00008b',
                             cursor="hand2")
    entry_institute.bind('<Button>',
                         lambda mouse: institution.my_institute(mouse))
    entry_institute.place(x=186, y=306)
    making = Button(main, text="CREATE INSTITUTION", activebackground='#67E1E4', relief=FLAT, borderwidth=0,
                    background='red', fg='white',
                    cursor="hand2")
    making.bind('<Button>', lambda mouse: institution.create_institute(mouse))
    making.place(x=186, y=269)
    mark = Button(main, text="Attendance", activebackground='red', relief=FLAT, borderwidth=0, background='red',
                  fg='white', width=20, cursor="hand2")
    mark.bind('<Button>',
              lambda mouse: institution.excel(mouse, class_var.get(), institute_var.get()))
    mark.place(x=635, y=370)
    export_excel = Button(main, text="Collect Data", activebackground='red', relief=FLAT, borderwidth=0,
                          background='red',
                          fg='white', width=20, cursor="hand2")
    export_excel.bind('<Button>',
                      lambda mouse: export(mouse, class_var.get(), institute_var.get(), main))
    export_excel.place(x=635, y=450)
    main.bind('<Key>', lambda key: institution.my_institute(key) or institution.create_institute(key) or institution
              .excel(key, class_var.get(), institute_var.get()) or export(key, class_var.get(), institute_var.get(),
                                                                          main))
    main.mainloop()


def outer_institute():

    def name_check(eve):
        if eve:
            inst_name = institution_name.get().strip().upper()
            if inst_name.isnumeric():
                messagebox.showinfo(title="Remainder", message="Institute Name can't only be Numbers")
            else:
                proceed = False
                if "EX: MEENAKSHI SUNDARARAJAN ENGINEERING COLLEGE" == inst_name or len(inst_name) == 0:
                    messagebox.showinfo(title="Remainder", message="Your Entry is Missing")
                else:
                    name = []
                    for _ in range(1):
                        if os.path.exists(college_dbpath):
                            board_class = sqlite3.connect(college_dbpath)
                            pen = board_class.cursor()
                            pen.execute("SELECT * FROM COLLEGE")
                            get_name = pen.fetchall()
                            for _name_ in get_name:
                                college_name_ = _name_[0]
                                if remove_white_space(college_name_) == remove_white_space(inst_name):
                                    proceed = True
                            board_class.commit()
                            board_class.close()
                    else:
                        if proceed:
                            messagebox.showinfo(title="Remainder", message='Institute Name already Exits')
                        else:
                            if messagebox.askyesno('Confirmation',
                                                   message=f'Are you sure with Institute Name -->'
                                                           f'{inst_name.upper()}?'):
                                least.destroy()
                                name.append(inst_name)
                                for _ in range(1):
                                    if os.path.exists(college_dbpath) is False:
                                        connect = sqlite3.connect(college_dbpath)
                                        con = connect.cursor()
                                        con.execute(f'CREATE TABLE COLLEGE (name text)')
                                        connect.commit()
                                        connect.close()
                                else:
                                    if os.path.exists(college_dbpath):
                                        connect1 = sqlite3.connect(college_dbpath)
                                        con1 = connect1.cursor()
                                        con1.execute('INSERT INTO COLLEGE VALUES (?)', name)
                                        connect1.commit()
                                        connect1.close()
                                        class_path = os.path.join(class_folder, f"{inst_name}.db")
                                        if os.path.exists(class_path) is False:
                                            connect2 = sqlite3.connect(class_path)
                                            con2 = connect2.cursor()
                                            con2.execute(f'CREATE TABLE classes (name text)')
                                            connect2.commit()
                                            connect2.close()
    least = Tk()
    window(least)
    image = Image.open(background_image)
    imageview = ImageTk.PhotoImage(image)
    Label(least, image=imageview).pack()
    Label(least, text="Name of your Institution", fg='#FFFFFF', bg="#FF0051", relief=FLAT).place(x=180, y=250)
    institution_name = StringVar()
    college_name = Entry(least, fg='#666666', bg="white", relief=FLAT, textvariable=institution_name, width=70,
                         state='readonly')
    college_name.insert(0, "Ex: Meenakshi Sundararajan Engineering College")
    clicked = college_name.bind('<Button>', lambda eve: click(college_name, eve, clicked))
    college_name.place(x=280, y=280)
    load = Image.open('My Class button.png')
    render = ImageTk.PhotoImage(load)
    pro = Button(least, image=render, bg='#00203F', activebackground='#00203F', relief=FLAT, borderwidth=0,
                 cursor='hand2')
    pro.bind('<Button>', name_check)
    pro.place(x=420.75, y=434.22)
    least.bind('<Return>', name_check)
    least.mainloop()


def system():
    for qs in range(1):
        total_check()
        if os.path.exists(college_dbpath) is False:
            outer_institute()
    else:
        if os.path.exists(college_dbpath):
            coll = sqlite3.connect(college_dbpath)
            bus = coll.cursor()
            bus.execute("SELECT * FROM COLLEGE")
            institute_name = bus.fetchall()
            coll.commit()
            coll.close()
            for q in range(1):
                if len(institute_name) == 0:
                    outer_institute()
                    if os.path.exists(college_dbpath):
                        coll = sqlite3.connect(college_dbpath)
                        bus = coll.cursor()
                        bus.execute("SELECT * FROM COLLEGE")
                        institute_name = bus.fetchall()
                        coll.commit()
                        coll.close()
            else:
                if len(institute_name) != 0:
                    first_window()


if __name__ == '__main__':
    system()


#   from PDFWriter import PDFWriter
#     workbook = load_workbook(f"{str(get)}{str(month[days.month])}.xlsx", guess_types=True, data_only=True)
#    worksheet = workbook.active

#  pw = PDFWriter('fruits2.pdf')
#  pw.setFont('Courier', 12)
#  pw.setHeader('XLSXtoPDF.py - convert XLSX data to PDF')
#  pw.setFooter('Generated using openpyxl and xtopdf')
#  ws_range = worksheet.iter_rows('A1:L1')
#  for row in ws_range:
#   s = ''
#   for cell in row:
#       if cell.value is None:
#           s += ' ' * 11
#       else:
#           s += str(cell.value).rjust(10) + ' '
# pw.writeLine(s)
# pw.savePage()
# pw.close()
# wb.copy_worksheet('RAJAJune28')
# wb.save(f"{str(get)}{str(month[days.month])}.pdf")

#   wb.get_sheet_by_name(f'{str(get)}{str(month[days.month])}{days.day}')

#from tkcalendar import Calendar

# window = Tk()
# window.configure(background = "black")

# style = ttk.Style(window)
# style.theme_use('clam')  # change theme, you can use style.theme_names() to list themes

# cal = Calendar(window, background="black", disabledbackground="black", bordercolor="black",
#               headersbackground="black", normalbackground="black", foreground='white',
#               normalforeground='white', headersforeground='white')
# cal.config(background="black")
# cal.configure()
# cal.pack()
#root = Tk()

# Set geometry
#root.geometry("400x400")

# Add Calendar
#cal = Calendar(root)
#cal.pack(pady=20)


#def grad_date():
#    date.config(text="Selected Date is: " + cal.get_date())


# Add Button and Label
#Button(root, text="Get Date",
#       command=grad_date).pack(pady=20)

#date = Label(root, text="")
#date.pack(pady=20)
#root.mainloop()