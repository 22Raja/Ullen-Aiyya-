from main import *
from tkcalendar import DateEntry

excel_institute = constant
excel_class = constant_class
excel_color = '#666666'

# ==============institute window ================#


class Institute:
    def __init__(self, block, institute_name, institute_db):
        self.block = block
        self.institute = institute_name
        self.institute_db = institute_db

    def institution_handling(self, constant_name, color):
        institution_ = []
        if os.path.exists(college_dbpath):
            total_check()
            coll = sqlite3.connect(college_dbpath)
            bus = coll.cursor()
            bus.execute("SELECT * FROM COLLEGE")
            for s in bus.fetchall():
                for x in s:
                    institution_.append(x)
            coll.commit()
            coll.close()
            if len(institution_) != 0:
                self.institute_window(institution_, constant_name, color)
            else:
                system()
        else:
            system()

    def making_class(self, event, class_name, strength, hour, create_window):

        def label_entry(lx_axis, ly_axis, ex_axis, ey_axis, start, stop):
            for a in range(start, stop):
                roll1 = Label(new3, text=f'{a}-> ', bg='#67E1E4')
                roll1.place(x=lx_axis, y=next(ly_axis) * 30)
                roll_name1 = Entry(new3, relief=SOLID, fg='black', bg="white",
                                   textvariable=student_storage[a])
                roll_name1.place(x=ex_axis, y=next(ey_axis) * 30)
                if a == get_strength:
                    break

        def final(eve):
            if eve:
                special_no, empty, get_hour, get_class, correct_entry = [], [], [], [], []
                for student in student_storage:
                    students_name = student.get().strip().upper()
                    if name_quality_checker(students_name) and len(students_name) >= 1:
                        correct_entry.append(students_name)
                    else:
                        if name_quality_checker(students_name) and len(students_name) < 1:
                            if student_storage.index(student) != 0:
                                empty.append(student_storage.index(student))
                        else:
                            if student_storage.index(student) != 0:
                                special_no.append(student_storage.index(student))
                else:
                    if len(special_no) >= 1 and len(empty) >= 1:
                        messagebox.showinfo(title="Remainder",
                                            message=f" Name can't be Special characters or Numbers at Roll number--> "
                                                    f"{str(special_no)[1:-1]}")
                        messagebox.showinfo(title="Remainder",
                                            message=f"Fill at Roll number --> {str(empty)[1:-1]}")
                    elif len(special_no) >= 1 and len(empty) == 0:
                        messagebox.showinfo(title="Remainder",
                                            message=f" Name can't be Special characters or Numbers at Roll number--> "
                                                    f"{str(special_no)[1:-1]}")
                    elif len(special_no) == 0 and len(empty) >= 1:
                        messagebox.showinfo(title="Remainder",
                                            message=f"Fill at Roll numbers --> {str(empty)[1:-1]}")
                    else:
                        if messagebox.askyesno('Confirmation', message="Would you like To Proceed?"):
                            get_hour.append(total_hour)
                            if os.path.exists(college_dbpath):
                                get_class.append(class_name)
                                for _ in range(1):
                                    if os.path.exists(class_folder) is False:
                                        os.mkdir(class_folder)
                                    if os.path.exists(self.institute_db) is False:
                                        connect1 = sqlite3.connect(self.institute_db)
                                        conn1 = connect1.cursor()
                                        conn1.execute(f'CREATE TABLE classes (name text)')
                                        connect1.commit()
                                        connect1.close()
                                        messagebox.showinfo(title="Remainder",
                                                            message=f"Think u deleted the {self.institute} Data base")
                                else:
                                    student_name = [(number, names) for number, names in enumerate(correct_entry, 1)]
                                    if os.path.exists(section_folder) is False:
                                        os.mkdir(section_folder)
                                    section = os.path.join(section_folder, f'{class_name} {self.institute}.db')
                                    connect4 = sqlite3.connect(section)
                                    pen2 = connect4.cursor()
                                    pen2.execute("CREATE TABLE STUDENT (number INTEGER ,name text)")
                                    pen2.execute("CREATE TABLE HOUR (number INTEGER)")
                                    pen2.execute('INSERT INTO HOUR VALUES (?)', get_hour)
                                    pen2.executemany("INSERT INTO STUDENT VALUES (?,?)", student_name)
                                    connect4.commit()
                                    connect4.close()
                                    college_open = sqlite3.connect(self.institute_db)
                                    connect = college_open.cursor()
                                    connect.execute(f'INSERT INTO classes VALUES (?)', get_class)
                                    college_open.commit()
                                    college_open.close()
                                    messagebox.showinfo('Confirmation',
                                                        message=F"Successfully Created your {class_name} Class")
                                    self.block.destroy()
                                    if os.path.exists(self.institute_db):
                                        self.institution_handling(self.institute, 'black')
                                    else:
                                        self.institution_handling(constant, '#666666')
                            else:
                                messagebox.showerror(title="Remainder",
                                                     message=f"Think u deleted the Board Data base manually")
                                self.block.destroy()
                                self.institution_handling(constant, '#666666')
        if event:
            if class_name.isnumeric():
                messagebox.showinfo(title="Remainder", message="Class Name can't only be Numbers")
            else:
                if os.path.exists(college_dbpath):
                    if 'Select Strength' in strength and 'Select Hour' in hour and (
                            "EX: I ECE - A" in class_name or len(
                            class_name) == 0):
                        messagebox.showinfo(title="Remainder",
                                            message="Give Class Name and Choose Strength,Hours ")
                    elif 'Select Hour' in hour and 'Select Strength' in strength:
                        messagebox.showinfo(title="Remainder",
                                            message='Choose Strength and Hours')
                    elif 'Select Strength' in strength and ("EX: I ECE - A" in class_name or len(
                            class_name) == 0):
                        messagebox.showinfo(title="Remainder",
                                            message='Give Class Name and Strength')
                    elif 'Select Hour' in hour and ("EX: I ECE - A" in class_name or len(
                            class_name) == 0):
                        messagebox.showinfo(title="Remainder",
                                            message='Give Class Name and Hours')
                    elif 'Select Strength' in strength:
                        messagebox.showinfo(title="Remainder", message='Choose Strength')
                    elif 'Select Hour' in hour:
                        messagebox.showinfo(title="Remainder", message='Choose Hours')
                    elif "EX: I ECE - A" in class_name or len(class_name) == 0:
                        messagebox.showinfo(title="Remainder", message="Give Class Name")
                    else:
                        proceed = True
                        for _ in range(1):
                            if os.path.exists(self.institute_db):
                                open_institute = sqlite3.connect(self.institute_db)
                                con = open_institute.cursor()
                                con.execute("SELECT * FROM classes")
                                class_data = con.fetchall()
                                open_institute.commit()
                                open_institute.close()
                                for q in class_data:
                                    for data in q:
                                        if remove_white_space(class_name) == remove_white_space(data):
                                            proceed = False
                        else:
                            if proceed:
                                get_strength = int(strength)
                                total_hour = int(hour)
                                if messagebox.askyesno('Confirmation',
                                                       message=f'Are you sure with Section Name --> '
                                                               f'{class_name}?'):
                                    if os.path.exists(college_dbpath):
                                        create_window.destroy()
                                        new3 = Tk()
                                        window(new3)
                                        self.block = new3
                                        for _ in range(1):
                                            if get_strength >= 51:
                                                main_frame = Frame(new3)
                                                main_frame.pack(fill=BOTH, expand=1)
                                                my_canvas = Canvas(main_frame)
                                                my_canvas.pack(fill=BOTH, expand=1, anchor=S)
                                                my_scrollbar = ttk.Scrollbar(main_frame, orient=HORIZONTAL,
                                                                             cursor='hand2',
                                                                             command=my_canvas.xview)
                                                my_scrollbar.pack(side=BOTTOM, fill=X)
                                                my_canvas.configure(xscrollcommand=my_scrollbar.set)
                                                my_canvas.bind("<Configure>",
                                                               lambda d: my_canvas.config(
                                                                   scrollregion=my_canvas.bbox(ALL)))
                                                second_frame = Frame(my_canvas)
                                                my_canvas.create_window((0, 0), window=second_frame, anchor="nw")
                                                new3 = second_frame
                                        else:
                                            image = Image.open(background_image)
                                            imageview = ImageTk.PhotoImage(image)
                                            Label(new3, image=imageview).pack()
                                            student_storage = [StringVar() for _ in range(get_strength + 1)]
                                            for i in range(get_strength + 1):
                                                def generator(number):
                                                    for integer in range(6, number):
                                                        yield integer

                                                first = generator(16)
                                                second = generator(16)
                                                if i == 0:
                                                    Label(new3, text=f"Total Strength --> {get_strength}",
                                                          fg='#FFFFFF',
                                                          bg="#FF0051").place(
                                                        x=30, y=5)
                                                    Label(new3, text=f"Class Hour --> {total_hour}",
                                                          fg='#FFFFFF',
                                                          bg="#FF0051").place(
                                                        x=900, y=5)
                                                    Label(new3, text=f"Class Name --> {class_name}",
                                                          fg='#FFFFFF',
                                                          bg="#FF0051").place(
                                                        x=450, y=5)
                                                if i == 1:
                                                    label_entry(30, first, 60,
                                                                second, 1, 11)
                                                if i == 11:
                                                    label_entry(200, first, 230,
                                                                second, 11, 21)
                                                if i == 21:
                                                    label_entry(370, first, 400,
                                                                second, 21, 31)
                                                if i == 31:
                                                    label_entry(540, first, 570,
                                                                second, 31, 41)
                                                if i == 41:
                                                    label_entry(710, first, 740,
                                                                second, 41, 51)
                                                if i == 51:
                                                    label_entry(880, first, 910,
                                                                second, 51, 61)
                                                if i == 61:
                                                    label_entry(1050, first, 1080,
                                                                second, 61, 71)
                                                if i == 71:
                                                    label_entry(1220, first, 1250,
                                                                second, 71, 81)
                                                if i == 81:
                                                    label_entry(1390, first, 1420,
                                                                second, 81, 91)
                                                if i == 91:
                                                    label_entry(1560, first, 1590,
                                                                second, 91,
                                                                101)
                                            else:
                                                def close(eve):
                                                    if eve.char == '`' or eve.char == '??':
                                                        if messagebox.askyesno('Confirmation',
                                                                               message=f'Are you sure to go back'):
                                                            self.block.destroy()
                                                            if os.path.exists(self.institute_db):
                                                                self.institution_handling(self.institute, 'black')
                                                            else:
                                                                self.institution_handling(constant, '#666666')
                                                Label(new3, text="STUDENT'S NAME", font='BOLD', bg='#67E1E4',
                                                      fg='#FFFFFF',
                                                      activebackground='#67E1E4',
                                                      activeforeground='#FFFFFF', cursor="hand2",
                                                      relief=FLAT, width=20).place(
                                                        x=400, y=30)

                                                mark = Button(new3, text='BACK', fg='#FFFFFF',
                                                              bg='#67E1E4', relief=FLAT, borderwidth=0,
                                                              cursor='hand2')
                                                mark.bind('<Button>', close)
                                                mark.place(x=0, y=30)

                                                proceed = Button(new3, text="Proceed",
                                                                 bg='#67E1E4',
                                                                 fg='#FFFFFF',
                                                                 activebackground='#67E1E4',
                                                                 activeforeground='#FFFFFF', cursor="hand2",
                                                                 relief=FLAT)
                                                proceed.bind('<Button>',
                                                             lambda mouse: final(mouse))
                                                proceed.place(x=450, y=500)
                                                new3.bind('<Return>',
                                                          lambda key: final(key))
                                                new3.bind('<Key>', close)
                                            new3.mainloop()
                                    else:
                                        messagebox.showerror(title="Remainder",
                                                             message=f"You deleted the Board Data base manually")
                                        self.block.destroy()
                                        self.institution_handling(constant, '#666666')
                            else:
                                messagebox.showinfo(title="Remainder", message='Class Name Already Exits')
                else:
                    messagebox.showerror(title="Remainder",
                                         message=f"You deleted the Board Data base manually")
                    self.block.destroy()
                    self.institution_handling(constant, '#666666')

    def create_class(self, event, institute_name):
        if event.char == 'c' or event.char == '??':
            if 'Select Institution' in institute_name:
                messagebox.showinfo(title="Remainder",
                                    message=f"Select an Institution")
            else:
                institution_db = os.path.join(class_folder, f'{institute_name}.db')
                if os.path.exists(institution_db) and os.path.exists(college_dbpath):
                    self.block.destroy()
                    create_window = Tk()
                    window(create_window)
                    image = Image.open(background_image)
                    imageview = ImageTk.PhotoImage(image)
                    Label(create_window, image=imageview).pack()
                    ttk.Style().configure('TCombobox', foreground='black')
                    strength = StringVar()
                    section_hour = StringVar()
                    class_name_data = StringVar()
                    class_name = Entry(create_window, fg='#666666', bg="white", textvariable=class_name_data,
                                       relief=FLAT,
                                       width=60)
                    class_name.insert(0, "Ex: I ECE - A")
                    clicked = class_name.bind('<Button>', lambda eve: click(class_name, eve, clicked))
                    class_name.configure(state='readonly')
                    class_name.place(x=400, y=250)
                    strength_box = ttk.Combobox(create_window, textvariable=strength, values=list(range(1, 101)),
                                                state='readonly', justify='center',
                                                font=('times', 10))
                    strength_box.set('---- Select Strength ----')
                    strength_box.place(x=200, y=150)
                    hour_box = ttk.Combobox(create_window, textvariable=section_hour, values=list(range(1, 11)),
                                            state='readonly', justify='center',
                                            font=('times', 10))
                    hour_box.set('---- Select Hour ----')
                    hour_box.place(x=700, y=150)
                    self.institute = institute_name
                    self.institute_db = institution_db
                    confirm = Button(create_window, text="Proceed", bg='#FF0051', fg='#FFFFFF',
                                     activebackground='#FF0051',
                                     activeforeground='#FFFFFF', relief=FLAT)
                    confirm.bind('<Button>',
                                 lambda mouse: self.making_class(mouse, class_name_data.get().strip().upper(),
                                                                 strength.get(),
                                                                 section_hour.get(), create_window))
                    confirm.place(x=400, y=450)

                    def close(eve):
                        if eve.char == '`' or eve.char == '??':
                            create_window.destroy()
                            if os.path.exists(institution_db):
                                self.institution_handling(institute_name, 'black')
                            else:
                                self.institution_handling(constant, '#666666')

                    back = Button(create_window, text='BACK', fg='#FFFFFF',
                                  bg="#FF0051", relief=FLAT, borderwidth=0,
                                  cursor='hand2')
                    back.bind('<Button>', close)
                    back.place(x=0, y=23)
                    create_window.bind('<Return>',
                                       lambda mouse: self.making_class(mouse, class_name_data.get().strip().upper(),
                                                                       strength.get(),
                                                                       section_hour.get(), create_window))
                    create_window.bind('<Key>', close)
                    create_window.mainloop()
                else:
                    board_deleted(institution_db, institute_name)
                    self.block.destroy()
                    self.institution_handling(constant, '#666666')

    def process_class(self, event, institute_name):
        if event.char == 'p' or event.char == '??':
            if 'Select Institution' in institute_name:
                messagebox.showinfo(title="Remainder",
                                    message=f"Select an Institution")
            else:
                institution_db = os.path.join(class_folder, f'{institute_name}.db')
                if os.path.exists(institution_db) and os.path.exists(college_dbpath):
                    sections = []
                    connect5 = sqlite3.connect(institution_db)
                    pen3 = connect5.cursor()
                    pen3.execute("SELECT * FROM classes")
                    for s in pen3.fetchall():
                        for x in s:
                            sections.append(x)
                    connect5.commit()
                    connect5.close()
                    if len(sections) != 0:
                        self.block.destroy()
                        from section import class_window
                        class_window.database = institution_db
                        class_window.institution_name = institute_name
                        class_window.section_window(sections, constant_class, '#666666')
                    else:
                        messagebox.showinfo(title="Remainder", message='Create your sections')
                else:
                    board_deleted(institution_db, institute_name)
                    self.block.destroy()
                    self.institution_handling(constant, '#666666')

    def delete_institute(self, event, institute_name):
        if event.char == 'd' or event.char == '??':
            if 'Select Institution' in institute_name:
                messagebox.showinfo(title="Remainder", message=f"Select an Institute ")
            else:
                name = institute_name.lower()
                if messagebox.askyesno("Confirmation", f'Are you sure to Delete {name} Institute?'):
                    institution_db = os.path.join(class_folder, f'{institute_name}.db')
                    if os.path.exists(college_dbpath) and os.path.exists(institution_db):
                        os.remove(institution_db)
                        messagebox.showinfo('Confirmation',
                                            message=f'Successfully Deleted {name} Institute')
                        self.block.destroy()
                        self.institution_handling(constant, '#666666')
                    else:
                        board_deleted(institution_db, institute_name)
                        self.block.destroy()
                        self.institution_handling(constant, '#666666')

    def modify_institute(self, event, institute_name):
        def change_institution(eve):
            if eve:
                inst_name = name_data.get().strip().upper()
                proceed = False
                if len(inst_name) == 0:
                    messagebox.showinfo(title="Remainder", message="Your Entry is Missing")
                elif remove_white_space(inst_name) == remove_white_space(institute_name):
                    messagebox.showinfo(title="Remainder", message="No Changes has been done")
                elif inst_name.isnumeric():
                    messagebox.showinfo(title="Remainder", message=f"Institute Name can't be only Numbers")
                else:
                    if messagebox.askyesno('Confirmation',
                                           message=f'Are you sure with the Institute Name --> '
                                                   f'{inst_name.upper()}?'):
                        _id_ = []
                        if os.path.exists(institution_db) and os.path.exists(college_dbpath):
                            for _ in range(1):
                                for s in range(1):
                                    board_class = sqlite3.connect(college_dbpath)
                                    pen = board_class.cursor()
                                    pen.execute("SELECT * FROM COLLEGE")
                                    for _name_ in pen.fetchall():
                                        for _ in _name_:
                                            if remove_white_space(_) == remove_white_space(inst_name):
                                                proceed = True
                                    board_class.commit()
                                    board_class.close()
                                else:
                                    if proceed:
                                        messagebox.showinfo(title="Remainder",
                                                            message='Institution Name already Exits')
                                    else:
                                        org_folder = os.path.join(excel_folder, institute_name)
                                        if os.path.exists(excel_folder) and os.path.exists(
                                                org_folder):
                                            temp_folder = os.path.join(parent_dir, '_temp_')
                                            try:
                                                os.rename(excel_folder, temp_folder)
                                            except OSError:
                                                messagebox.showinfo("Reminder", 'Kindly close your Folder')
                                                break
                                            else:
                                                os.rename(temp_folder, excel_folder)
                                                for excel_name in os.listdir(excel_folder):
                                                    if remove_white_space(institute_name) == remove_white_space(
                                                            excel_name):
                                                        os.rename(os.path.join(excel_folder, excel_name),
                                                                  os.path.join(excel_folder, inst_name))
                            else:
                                _id_.append(institute_name)
                                connect4 = sqlite3.connect(college_dbpath)
                                pen2 = connect4.cursor()
                                pen2.execute("SELECT rowid FROM COLLEGE WHERE name = ?", _id_)
                                rowid = pen2.fetchone()[0]
                                pen2.execute("UPDATE COLLEGE SET name = ? WHERE rowid = ?",
                                             (inst_name, rowid))
                                connect4.commit()
                                connect4.close()
                                modify_db = os.path.join(class_folder, f'{inst_name}.db')
                                os.rename(institution_db, modify_db)
                                connect5 = sqlite3.connect(modify_db)
                                pen3 = connect5.cursor()
                                pen3.execute("SELECT * FROM classes")
                                for open_db in pen3.fetchall():
                                    for class_name in open_db:
                                        class_db = os.path.join(section_folder, f'{class_name} {institute_name}.db')
                                        os.rename(class_db,
                                                  os.path.join(section_folder, f'{class_name} {inst_name}.db'))
                                else:
                                    connect5.commit()
                                    connect5.close()
                                    messagebox.showinfo("Reminder", 'Successfully updated')
                                    change_window.destroy()
                                    if os.path.exists(modify_db):
                                        self.institution_handling(inst_name, 'black')
                                    else:
                                        self.institution_handling(constant, '#666666')
                        else:
                            board_deleted(institution_db, institute_name)
                            change_window.destroy()
                            self.institution_handling(constant, '#666666')
        if event.char == 'm' or event.char == '??':
            if 'Select Institution' in institute_name:
                messagebox.showinfo(title="Remainder",
                                    message=f"Select an Institution")
            else:
                institution_db = os.path.join(class_folder, f'{institute_name}.db')
                if os.path.exists(institution_db) and os.path.exists(college_dbpath):
                    self.block.destroy()
                    change_window = Tk()
                    window(change_window)
                    image = Image.open(background_image)
                    imageview = ImageTk.PhotoImage(image)
                    Label(change_window, image=imageview).pack()
                    name_data = StringVar()
                    entry = Entry(change_window, fg='black', bg="white", textvariable=name_data, relief=FLAT,
                                  width=60)
                    entry.insert(0, institute_name)
                    entry.place(x=400, y=250)
                    confirm = Button(change_window, text="Proceed", bg='#FF0051', fg='#FFFFFF',
                                     activebackground='#FF0051',
                                     activeforeground='#FFFFFF', relief=FLAT)
                    confirm.bind('<Button>', change_institution)
                    confirm.place(x=400, y=450)

                    def close(eve):
                        if eve.char == '`' or eve.char == '??':
                            change_window.destroy()
                            if os.path.exists(institution_db):
                                self.institution_handling(institute_name, 'black')
                            else:
                                self.institution_handling(constant, '#666666')

                    mark = Button(change_window, text='BACK', fg='#FFFFFF',
                                  bg="#FF0051", relief=FLAT, borderwidth=0,
                                  cursor='hand2')
                    mark.bind('<Button>', close)
                    mark.place(x=0, y=23)
                    change_window.bind('<Return>', change_institution)
                    change_window.bind('<Key>', close)
                    change_window.mainloop()
                else:
                    board_deleted(institution_db, institute_name)
                    self.block.destroy()
                    self.institution_handling(constant, '#666666')

    def clear(self, event, institute_name):
        def _excel_(eve):
            if eve:
                _name_ = excel_variable.get()
                if "SELECT INSTITUTE" in _name_:
                    messagebox.showinfo(title="Remainder",
                                        message=f"Select a folder Name")
                else:
                    if messagebox.askyesno(title="Remainder",
                                           message=f"Are you sure to clear {_name_} "):
                        if os.path.exists(excel_folder):
                            temp = os.path.join(parent_dir, '__M__')
                            try:
                                os.rename(excel_folder, temp)
                            except OSError:
                                messagebox.showinfo("Reminder", 'Kindly close your Folder')
                            else:
                                os.rename(temp, excel_folder)
                                institute_folder = os.path.join(excel_folder, _name_)
                                for _ in range(1):
                                    if os.path.exists(institute_folder):
                                        for get_folder in os.listdir(institute_folder):
                                            cls_folder = os.path.join(institute_folder, get_folder)
                                            for get1_folder in os.listdir(cls_folder):
                                                year_folder = os.path.join(cls_folder, get1_folder)
                                                for excel in os.listdir(year_folder):
                                                    os.remove(os.path.join(year_folder, excel))
                                                else:
                                                    os.rmdir(year_folder)
                                            else:
                                                os.rmdir(cls_folder)
                                        else:
                                            os.rmdir(institute_folder)
                                else:
                                    messagebox.showinfo(title="Remainder",
                                                        message=f"Hey successfully cleared the deleted folder ")
                                    excel_window.destroy()
                                    if os.path.exists(institute_db):
                                        self.institution_handling(institute_name, 'black')
                                    else:
                                        self.institution_handling(constant, '#666666')
                        else:
                            messagebox.showinfo(title="Remainder",
                                                message=f"You deleted the excel folder")
                            excel_window.destroy()
                            if os.path.exists(institute_db):
                                self.institution_handling(institute_name, 'black')
                            else:
                                self.institution_handling(constant, '#666666')

        if event.char == 'x' or event.char == '??':
            if os.path.exists(excel_folder):
                excel_data = list(filter(lambda get: '#Deleted' in get, os.listdir(excel_folder)))
                if len(excel_data) != 0:
                    self.block.destroy()
                    excel_window = Tk()
                    window(excel_window)
                    image = Image.open(background_image)
                    imageview = ImageTk.PhotoImage(image)
                    Label(excel_window, image=imageview).pack()
                    institute_db = os.path.join(class_folder, f'{institute_name}.db')
                    excel_variable = StringVar()
                    strength_box = ttk.Combobox(excel_window, textvariable=excel_variable, values=excel_data,
                                                state='readonly', justify='center', width=60,
                                                font=('times', 10))
                    strength_box.set('SELECT INSTITUTE')
                    strength_box.place(x=350, y=200)
                    proceed = Button(excel_window, text="Proceed", bg='#67E1E4', fg='#FFFFFF',
                                     activebackground='#67E1E4',
                                     activeforeground='#FFFFFF', cursor="hand2", relief=FLAT, width=10, font='BOLD')
                    proceed.bind('<Button>', _excel_)
                    proceed.place(x=430, y=450)

                    def close(eve):
                        if eve.char == '`' or eve.char == '??':
                            excel_window.destroy()
                            if os.path.exists(institute_db):
                                self.institution_handling(institute_name, 'black')
                            else:
                                self.institution_handling(constant, '#666666')
                    mark = Button(excel_window, text='BACK', fg='#FFFFFF',
                                  bg="#FF0051", relief=FLAT, borderwidth=0,
                                  cursor='hand2')
                    mark.bind('<Button>', close)
                    mark.place(x=0, y=23)
                    excel_window.bind('<Key>', close)
                    excel_window.bind('<Return>', _excel_)
                    excel_window.mainloop()

                else:
                    messagebox.showinfo(title="Remainder", message=f"No folder there to clear")
            else:
                messagebox.showinfo(title="Remainder", message=f"No folder there to clear")

    # The Institute window
    def institute_window(self, institution_data, constant_name, color):
        new1 = Tk()
        window(new1)
        image_view = PhotoImage(file=background_image)
        Label(new1, image=image_view).pack()
        self.block = new1
        institution_var = StringVar()
        ttk.Style().configure('TCombobox', foreground=color)
        institution_box = ttk.Combobox(new1, textvariable=institution_var, values=institution_data,
                                       state='readonly', justify='center',
                                       width=60, cursor="hand2",
                                       font=('times', 10))
        institution_box.set(constant_name)
        institution_box.place(x=345, y=20)
        class_button = PhotoImage(file='My Class button.png')
        create_button = PhotoImage(file='Create class button.png')

        def close(eve):
            if eve.char == '`' or eve.char == '??':
                new1.destroy()
                system()

        mark = Button(new1, text='BACK', fg='#FFFFFF',
                      bg="#FF0051", relief=SOLID,
                      cursor='hand2')
        mark.bind('<Button>', close)
        mark.place(x=0, y=23)
        create = Button(new1, image=create_button, activebackground='#67E1E4', relief=FLAT, borderwidth=0,
                        background='#67E1E4',
                        cursor="hand2")
        create.bind('<Button>', lambda mouse: self.create_class(mouse, institution_var.get()))
        create.place(x=545.65, y=280.45)
        process = Button(new1, image=class_button, activebackground='#67E1E4', relief=FLAT, borderwidth=0,
                         background='#67E1E4',
                         cursor="hand2")
        process.bind('<Button>', lambda mouse: self.process_class(mouse, institution_var.get()))
        process.place(x=295.84, y=280.45)
        delete = Button(new1, text="DELETE ", bg='#67E1E4', fg='#FFFFFF', activebackground='#67E1E4',
                        activeforeground='#FFFFFF', cursor="hand2", relief=FLAT)
        delete.bind('<Button>', lambda mouse: self.delete_institute(mouse, institution_var.get()))
        delete.place(x=120, y=150)
        modify = Button(new1, text="MODIFY", bg='#67E1E4', fg='#FFFFFF', activebackground='#67E1E4',
                        activeforeground='#FFFFFF', cursor="hand2", relief=FLAT)
        modify.bind('<Button>', lambda mouse: self.modify_institute(mouse, institution_var.get()))
        modify.place(x=680, y=150)
        clear = Button(new1, text="CLEAR", bg='#67E1E4', fg='#FFFFFF', activebackground='#67E1E4',
                       activeforeground='#FFFFFF', cursor="hand2", relief=FLAT)
        clear.bind('<Button>', lambda mouse: self.clear(mouse, institution_var.get()))

        clear.place(x=680, y=400)
        new1.bind('<Key>',
                  lambda key: close(key) or self.create_class(key,
                                                              institution_var.get()) or self.process_class(
                      key, institution_var.get()) or self.delete_institute(key,
                                                                           institution_var.get()) or self.
                  modify_institute(
                      key, institution_var.get()) or self.clear(key, institution_var.get()))

        new1.mainloop()

    # Starting point of the institute window
    def my_institute(self, event):
        if event.char == 'i' or event.char == '??':
            if os.path.exists(college_dbpath):
                total_check()
                institute_data = []
                coll = sqlite3.connect(college_dbpath)
                bus = coll.cursor()
                bus.execute("SELECT * FROM COLLEGE")
                for s in bus.fetchall():
                    for x in s:
                        institute_data.append(x)
                coll.commit()
                coll.close()
                if len(institute_data) != 0:
                    self.block.destroy()
                    self.institute_window(institute_data, constant, '#666666')
                else:
                    messagebox.showerror(title="Remainder",
                                         message=f'You Deleted a Data Base ')
                    self.block.destroy()
                    system()
            else:
                messagebox.showerror(title="Remainder",
                                     message=f'You Deleted Board Data Base ')
                self.block.destroy()
                system()

    # Making institute
    def create_institute(self, event):

        def name_check(han):
            if han:
                inst_name = institution_name.get().strip().upper()
                if inst_name.isnumeric():
                    messagebox.showinfo(title="Remainder", message=f"Institution Name can't only be Numbers")
                else:
                    proceed = False
                    if "EX: MEENAKSHI SUNDARARAJAN ENGINEERING COLLEGE" == inst_name or len(inst_name) == 0:
                        messagebox.showinfo(title="Remainder", message="Your Entry is Missing")
                    else:
                        name = []
                        for g in range(1):
                            total_check()
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
                                                       message=f'Are you sure with Institute Name --> '
                                                               f'{inst_name.lower()}?'):
                                    create_institute.destroy()
                                    name.append(inst_name)
                                    for _ in range(1):
                                        if os.path.exists(college_dbpath) is False:
                                            connect2 = sqlite3.connect(college_dbpath)
                                            con1 = connect2.cursor()
                                            con1.execute(f'CREATE TABLE COLLEGE (name text)')
                                            connect2.commit()
                                            connect2.close()
                                    else:
                                        if os.path.exists(college_dbpath):
                                            connect4 = sqlite3.connect(college_dbpath)
                                            pen2 = connect4.cursor()
                                            pen2.execute('INSERT INTO COLLEGE VALUES (?)', name)
                                            connect4.commit()
                                            connect4.close()
                                            class_path = os.path.join(class_folder, f"{inst_name}.db")
                                            if os.path.exists(class_path) is False:
                                                connect2 = sqlite3.connect(class_path)
                                                con1 = connect2.cursor()
                                                con1.execute(f'CREATE TABLE classes (name text)')
                                                connect2.commit()
                                                connect2.close()
                                        system()

        if event.char == 'c' or event.char == '??':
            self.block.destroy()
            create_institute = Tk()
            window(create_institute)
            image = Image.open(background_image)
            imageview = ImageTk.PhotoImage(image)
            Label(create_institute, image=imageview).pack()
            Label(create_institute, text="Name of your Institution", fg='#FFFFFF', bg="#FF0051", relief=FLAT).place(x=180,y=250)
            institution_name = StringVar()
            college_name = Entry(create_institute, fg='#666666', bg="white", relief=FLAT, textvariable=institution_name,
                                 width=70)
            college_name.insert(0, "Ex: Meenakshi Sundararajan Engineering College")
            clicked = college_name.bind('<Button>', lambda eve: click(college_name, eve, clicked))
            college_name.config(state='readonly')
            college_name.place(x=280, y=280)
            load = Image.open('My Class button.png')
            render = ImageTk.PhotoImage(load)

            def close(eve):
                if eve.char == '`' or eve.char == '??':
                    create_institute.destroy()
                    system()

            mark = Button(create_institute, text='BACK', fg='#FFFFFF',
                          bg="#FF0051", relief=FLAT, borderwidth=0,
                          cursor='hand2')
            mark.bind('<Button>', close)
            mark.place(x=0, y=23)
            entry = Button(create_institute, image=render, bg='#00203F', activebackground='#00203F', relief=FLAT,
                           borderwidth=0,
                           cursor='hand2')
            entry.bind('<Button>', name_check)
            entry.place(x=420.75, y=434.22)
            create_institute.bind('<Return>', name_check)
            create_institute.bind('<Key>', close)
            create_institute.mainloop()

    def excel(self, event, class_name, institute_name):
        if event.char == 'm' or event.char == '??':
            if 'Select Institution' in institute_name and 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Choose Institution and Class')
            elif 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Choose the Class')
            elif 'Select Institution' in institute_name:
                messagebox.showinfo(title="Remainder", message='Choose the Institution')
            else:
                class_db = os.path.join(section_folder, f'{class_name} {institute_name}.db')
                institute_db = os.path.join(class_folder, f'{institute_name}.db')
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(institute_db):
                    db = sqlite3.connect(class_db)
                    c5 = db.cursor()
                    c5.execute("SELECT rowid FROM STUDENT WHERE name =  'DELETED'")
                    deleted = list(map(lambda get: get[0], c5.fetchall()))
                    c5.execute("SELECT rowid FROM STUDENT WHERE name =  'NONE'")
                    left_institute = list(map(lambda get: get[0], c5.fetchall()))
                    c5.execute("SELECT rowid,name FROM STUDENT")
                    student_name = list(map(lambda x: x[1], c5.fetchall()))
                    my_class = int(c5.execute('SELECT COUNT(*) FROM STUDENT').fetchone()[0])
                    db.commit()
                    db.close()
                    duty = student_name.count('DELETED') + student_name.count('NONE')
                    if my_class == 0:
                        messagebox.showinfo(title="Remainder",
                                            message='Only Add names for this Class')
                    elif duty == my_class:
                        messagebox.showinfo(title="Remainder",
                                            message='Only Add and Modify names for this Class')
                    else:
                        class_var = StringVar()
                        hour_var = StringVar()
                        hour = ''
                        hour_status = ''
                        late_people = []
                        absent_people = []
                        od_people = []
                        self.block.destroy()
                        global excel_institute, excel_class
                        excel_institute = institute_name
                        excel_class = constant_class
                        get_date = days.day
                        total_hour = 0

                        def fin_nal():

                            def sheet_decision(roll, sheet_color, status):
                                for _ in range(1):
                                    if hour_status == 1:
                                        sheet[f'{hour}{roll}'].value = status
                                    elif hour_status == 2:
                                        sheet[f'{hour}{roll}'].value = f'EXAM/{status}'
                                    elif hour_status == 3:
                                        sheet[f'{hour}{roll}'].value = f'LAB/{status}'
                                sheet[f'{hour}{roll}'].fill = sheet_color
                                sheet[f'{hour}{roll}'].alignment = Alignment(horizontal='center',
                                                                             wrap_text=True)

                            def excel_fill(sheet_color, status):
                                for roll, name in enumerate(student_name, 2):
                                    if name == 'NONE':
                                        sheet[f'{hour}{roll}'].value = 'NONE'
                                        sheet[f'{hour}{roll}'].fill = yellow

                                    elif name == 'DELETED':
                                        sheet[f'{hour}{roll}'].value = 'DELETED'
                                        sheet[f'{hour}{roll}'].fill = orange

                                    else:
                                        for o in range(1):
                                            if hour_status == 1:
                                                sheet[f'{hour}{roll}'].value = status
                                            elif hour_status == 2:
                                                sheet[f'{hour}{roll}'].value = f'EXAM/{status}'
                                            elif hour_status == 3:
                                                sheet[f'{hour}{roll}'].value = f'LAB/{status}'
                                        else:
                                            sheet[f'{hour}{roll}'].fill = sheet_color
                                    sheet[f'{hour}{roll}'].alignment = Alignment(horizontal='center',
                                                                                 wrap_text=True)

                            while True:
                                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                        institute_db):
                                    connection = os.path.join(excel_maker(institute_name, class_name),
                                                              f"{class_name} {month[days.month]}.xlsx")
                                    if os.path.exists(connection) is False:
                                        excel_sheet(class_db, connection, class_name, get_date)
                                    else:
                                        try:
                                            wb = load_workbook(connection)
                                            sheet = wb[f'{class_name} {get_date} {month[days.month]}']
                                            sheet['O16'].value = total_hour
                                            sheet_data(sheet, student_name, hour_data(class_db))
                                            for i in range(1):
                                                if 120 in od_people:
                                                    excel_fill(super_green, 'On Duty')
                                                elif 1000 in absent_people:
                                                    excel_fill(red, 'Absent')
                                                elif 140 in late_people:
                                                    excel_fill(white, 'Late')
                                                else:
                                                    excel_fill(white, 'Present')
                                                    for roll_no in od_people:
                                                        sheet_decision(roll_no + 1, super_green, 'On Duty')
                                                    for roll_no in absent_people:
                                                        sheet_decision(roll_no + 1, red, 'Absent')
                                                    for roll_no in late_people:
                                                        sheet_decision(roll_no + 1, white, 'Late')
                                            else:
                                                for clear in od_people, absent_people, late_people:
                                                    clear.clear()
                                            wb.save(connection)
                                            messagebox.showinfo('confirmation',
                                                                message="Attendance updated Successfully")
                                        except PermissionError:
                                            messagebox.showinfo("Reminder", 'Kindly close your opened excel file')
                                        except KeyError:
                                            try:
                                                excel_sheet(class_db, connection, class_name, get_date)
                                            except PermissionError:
                                                messagebox.showinfo("Reminder", 'Kindly close your opened excel file')
                                        else:
                                            self.block.destroy()
                                            attendance('od time', 'Everyone On Duty', 'NO On Duty', 120, [])
                                            break
                                else:
                                    final_check(class_name, class_db, self.block, institute_db, institute_name)
                                    system()

                        def late_time(eve, no_one, num):
                            if eve:
                                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                        institute_db):
                                    nonlocal late_people
                                    late_people = []
                                    for i in num:
                                        roll = i.get()
                                        if roll != 0:
                                            late_people.append(roll)
                                    if no_one == 1000:
                                        late_people.append(no_one)
                                    count_late = len(late_people)
                                    if count_late == 0:
                                        messagebox.showinfo(title="Remainder", message="Select a value")
                                    elif 140 in late_people and 1000 in late_people and count_late >= 2:
                                        messagebox.showinfo(title="Remainder",
                                                            message='IF Everyone_Late click -> Everyone_Late')
                                        messagebox.showinfo(title="Remainder",
                                                            message='IF No one_Late -> NO Behind_Late')
                                    elif 140 in late_people and 1000 not in late_people and count_late >= 2:
                                        messagebox.showinfo(title="Remainder",
                                                            message='IF Everyone_Late click -> Everyone_Late')
                                    elif 1000 in late_people and 140 not in late_people and count_late >= 2:
                                        messagebox.showinfo(title="Remainder",
                                                            message='IF No one_Late -> NO Behind_Late')
                                    else:
                                        if messagebox.askyesno('Confirmation',
                                                               message="Would you like To Proceed?"):
                                            for i in range(1):
                                                if 130 in absent_people:
                                                    absent_people.remove(130)
                                                if 1000 in late_people:
                                                    late_people.remove(1000)
                                                if 140 not in late_people:
                                                    if my_class == count_late + duty:
                                                        late_people.clear()
                                                        late_people.append(140)
                                            fin_nal()
                                else:
                                    final_check(class_name, class_db, self.block, institute_db, institute_name)
                                    system()

                        def normal_time(eve, no_one, num, display):
                            if eve:
                                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                        institute_db):
                                    h_time = class_var.get()
                                    h_sit = hour_var.get()

                                    def condition():
                                        if "HOUR" in h_time and "CLASS" in h_sit:
                                            messagebox.showinfo(title="Remainder",
                                                                message="Choose the Hour and Class")

                                        elif "HOUR" in h_time and "CLASS" not in h_sit:
                                            messagebox.showinfo(title="Remainder",
                                                                message="Choose the Hour ")
                                        elif "HOUR" not in h_time and "CLASS" in h_sit:
                                            messagebox.showinfo(title="Remainder",
                                                                message="Choose the Class ")
                                        else:
                                            nonlocal hour, hour_status
                                            hour = table[time.index(h_time)]
                                            hour_status = class_situation.index(h_sit) + 1
                                            return 'run'

                                    nonlocal absent_people
                                    absent_people = []
                                    if 120 not in display:
                                        for roll_no in num:
                                            roll = roll_no.get()
                                            if roll != 0:
                                                absent_people.append(roll)
                                        if no_one == 1000:
                                            absent_people.append(no_one)
                                        count_absent = len(absent_people)
                                        if count_absent == 0:
                                            messagebox.showinfo(title="Remainder",
                                                                message="Select a value")
                                        elif 130 in absent_people and 1000 in absent_people and count_absent >= 2:
                                            messagebox.showinfo(title="Remainder",
                                                                message="IF All present click-> Everyone present")
                                            messagebox.showinfo(title="Remainder",
                                                                message=f'IF No one Present-> No one Present')
                                        elif 130 in absent_people and 1000 not in absent_people and count_absent >= 2:
                                            messagebox.showinfo(title="Remainder",
                                                                message="IF All present click-> Everyone present")
                                        elif 1000 in absent_people and 130 not in absent_people and count_absent >= 2:
                                            messagebox.showinfo(title="Remainder",
                                                                message=f'IF No one Present-> No one Present')
                                        else:
                                            if condition() == 'run':
                                                if messagebox.askyesno('Confirmation',
                                                                       message="Would you like To Proceed?"):
                                                    # 1000 every one absent
                                                    # 130 every one present
                                                    absent = False
                                                    for _ in range(1):
                                                        if 1000 in display:
                                                            display.remove(1000)
                                                        if 130 not in absent_people and 1000 not in absent_people:
                                                            if my_class == count_absent + duty:
                                                                absent_people.clear()
                                                                absent_people.append(1000)
                                                            elif my_class == count_absent + duty + len(display):
                                                                absent = True
                                                    else:
                                                        if absent or 1000 in absent_people:
                                                            fin_nal()
                                                        else:
                                                            self.block.destroy()
                                                            attendance('late', 'Everyone Late', 'No one Late', 140,
                                                                       absent_people + display)
                                    else:
                                        if condition() == 'run':
                                            fin_nal()
                                else:
                                    final_check(class_name, class_db, self.block, institute_db, institute_name)
                                    system()

                        def od_time(eve, no_one, num):
                            if eve:
                                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                        institute_db):
                                    nonlocal od_people
                                    od_people = []
                                    for roll_no in num:
                                        roll = roll_no.get()
                                        if roll != 0:
                                            od_people.append(roll)
                                    if no_one == 1000:
                                        od_people.append(1000)
                                    if len(od_people) == 0:
                                        messagebox.showinfo(title="Remainder", message="Select a value")

                                    elif 120 in od_people and 1000 in od_people and len(od_people) >= 2:
                                        messagebox.showinfo(title="Remainder",
                                                            message=f'IF Everyone On Duty click-> Everyone On Duty ')
                                        messagebox.showinfo(title="Remainder",
                                                            message=f'IF NO On Duty click-> NO On Duty')
                                    elif 120 in od_people and 1000 not in od_people and len(od_people) >= 2:
                                        messagebox.showinfo(title="Remainder",
                                                            message=f'IF Everyone On Duty click-> Everyone On Duty ')
                                    elif 1000 in od_people and 120 not in od_people and len(od_people) >= 2:
                                        messagebox.showinfo(title="Remainder",
                                                            message=f'IF NO On Duty click-> NO On Duty')
                                    else:
                                        if messagebox.askyesno('Confirmation',
                                                               message="Would you like To Proceed?"):
                                            for _ in range(1):
                                                self.block.destroy()
                                                if 1000 in od_people:
                                                    od_people.remove(1000)
                                                else:
                                                    if 120 not in od_people:
                                                        if my_class == len(od_people) + duty:
                                                            od_people.clear()
                                                            od_people.append(120)
                                            else:
                                                if 120 not in od_people:
                                                    attendance('period', 'Everyone Present', 'No one Present',
                                                               130, od_people)
                                                else:
                                                    attendance('period', 'Surprises that Everyone is ON DUTY', ' ',
                                                               130, od_people)
                                else:
                                    final_check(class_name, class_db, self.block, institute_db, institute_name)
                                    system()

                        #  time_period, status, case, display, value
                        #  'od time', 'Everyone On Duty', 'NO On Duty', [], 120

                        def attendance(time_period, for_all, no_one, value, display):

                            def button_verify(show_of, cx_axis, cy_axis, start, stop,
                                              number):
                                for a in range(start, stop):
                                    check_button = Checkbutton(show_of, text=f'{a}', variable=num[a], onvalue=a,
                                                               fg='black',
                                                               bg='#67E1E4')
                                    for roll_no in [display, left_institute, deleted]:
                                        if a in roll_no:
                                            check_button.config(state='disabled')
                                    check_button.place(x=cx_axis, y=next(cy_axis) * 30)
                                    if a == number:
                                        break
                            master = Tk()
                            window(master)
                            image = Image.open(background_image)
                            imageview = ImageTk.PhotoImage(image)
                            Label(master, image=imageview).pack()
                            self.block = master
                            nonlocal class_var, hour_var
                            class_var = StringVar()
                            hour_var = StringVar()
                            no_people = IntVar()
                            num = [IntVar() for _ in range(my_class + 1)]
                            for i in range(my_class + 1):
                                def generator(number):
                                    for integer in range(6, number):
                                        yield integer
                                first = generator(number=16)
                                if 120 not in display:
                                    if i == 0:
                                        Checkbutton(master, text=f'{for_all}', variable=num[0], onvalue=value,
                                                    fg='black',
                                                    bg="#e6e6e6").place(x=120, y=111)
                                        Checkbutton(master, text=f'{no_one}', fg='black', variable=no_people,
                                                    onvalue=1000,
                                                    bg="#e6e6e6").place(x=800, y=110)
                                    if i == 1:
                                        button_verify(master, 130, first, 1, 11, my_class)
                                    if i == 11:
                                        button_verify(master, 210, first, 11, 21, my_class)
                                    if i == 21:
                                        button_verify(master, 290, first, 21, 31, my_class)
                                    if i == 31:
                                        button_verify(master, 370, first, 31, 41, my_class)
                                    if i == 41:
                                        button_verify(master, 450, first, 41, 51, my_class)
                                    if i == 51:
                                        button_verify(master, 530, first, 51, 61, my_class)
                                    if i == 61:
                                        button_verify(master, 610, first, 61, 71, my_class)
                                    if i == 71:
                                        button_verify(master, 690, first, 71, 81, my_class)
                                    if i == 81:
                                        button_verify(master, 770, first, 81, 91, my_class)
                                    if i == 91:
                                        button_verify(master, 850, first, 91, 101, my_class)
                                else:
                                    Label(master, bg='#67E1E4', fg='#FFFFFF', text=for_all).place(
                                        x=340, y=200)

                            else:
                                title = Label(master, relief=FLAT, fg='black',
                                              bg='#67E1E4', width=50)
                                title.place(x=340, y=10)
                                for _ in range(1):
                                    def close(eve):
                                        if eve.char == '`' or eve.char == '??':
                                            global excel_institute, excel_class, excel_color
                                            excel_institute = institute_name
                                            excel_class = class_name
                                            excel_color = 'black'
                                            master.destroy()
                                            system()

                                    mark = Button(master, text='BACK', fg='#FFFFFF',
                                                  bg="#FF0051", relief=FLAT, borderwidth=0,
                                                  cursor='hand2')
                                    mark.bind('<Button>', close)
                                    mark.place(x=0, y=23)
                                    master.bind('<Key>', close)
                                    left_count = len(left_institute)
                                    delete_count = len(deleted)
                                    od_count = len(od_people)
                                    absent_count = len(absent_people)
                                    if left_count != 0 and left_count <= 10:
                                        Label(master, bg='#67E1E4',
                                              text=f'LEFT INSTITUTION --> {str(left_institute)[1:-1]}').place(
                                            x=730, y=10)
                                    elif left_count >= 11:
                                        messagebox.showinfo(title="Remainder",
                                                            message=f'LEFT INSTITUTION --> {str(left_institute)[1:-1]}')

                                    if delete_count != 0 and delete_count <= 10:
                                        Label(master, bg='#67E1E4', text=f'DELETED --> {str(deleted)[1:-1]}').place(
                                            x=100,
                                            y=10)
                                    elif delete_count >= 11:
                                        messagebox.showinfo(title="Remainder",
                                                            message=f'DELETED --> {str(deleted)[1:-1]}')
                                    if 120 not in od_people and 1000 not in od_people:
                                        if od_count != 0 and od_count <= 10:
                                            Label(master, bg='#67E1E4',
                                                  text=f'ON DUTY --> {str(od_people)[1:-1]}').place(
                                                x=100,
                                                y=500)
                                        elif od_count >= 11:
                                            messagebox.showinfo(title="Remainder",
                                                                message=f'ON DUTY --> {str(od_people)[1:-1]}')
                                    if 130 not in absent_people:
                                        if absent_count != 0 and absent_count <= 10:
                                            Label(master, bg='#67E1E4',
                                                  text=f'ABSENTEES --> {str(absent_people)[1:-1]}').place(
                                                x=730, y=500)
                                        elif absent_count >= 11:
                                            messagebox.showinfo(title="Remainder",
                                                                message=f'ABSENTEES --> {str(absent_people)[1:-1]}')
                                else:
                                    if time_period == 'od time':
                                        title['text'] = 'ON DUTY ATTENDANCE'

                                        def color_change(eve):
                                            if eve:
                                                nonlocal get_date
                                                get_date = date_pick.get_date().day

                                        date_pick = DateEntry(master, width=12, background='#67E1E4', state='readonly',
                                                              cursor="hand2", selectmode='day',
                                                              foreground='white', date_pattern='d/m/yy',
                                                              justify='center', mindate=date(year, days.month, 1),
                                                              maxdate=date(year, days.month, days.day),
                                                              normalforeground='#67E1E4', selectbackground='#67E1E4',
                                                              showweeknumbers=False,
                                                              headersforeground='white', headersbackground='#67E1E4',
                                                              bordercolor='#67E1E4')
                                        date_pick.bind("<<DateEntrySelected>>", lambda sel: color_change(sel))
                                        date_pick.place(x=125, y=60)
                                        # convey = Message(master, text="Hey!? The selected sheet ?", relief=FLAT)
                                        # convey.place(x=215, y=81)
                                        pro = Button(master, text="Proceed", fg='white', bg="#FF0051",
                                                     width=20, relief=FLAT, cursor='hand2')
                                        pro.bind('<Button>',
                                                 lambda mouse: od_time(mouse, no_people.get(), num))
                                        pro.place(x=410, y=480)
                                        master.bind('<Return>',
                                                    lambda enter: od_time(enter, no_people.get(), num))
                                    elif time_period == 'period':
                                        class_hour = ttk.Combobox(master, textvariable=class_var,
                                                                  state='readonly', cursor="hand2",
                                                                  justify=CENTER,
                                                                  font=('times', 10))
                                        class_hour.set("HOUR")
                                        class_hour.place(x=100, y=50)
                                        situation = ttk.Combobox(master, textvariable=hour_var,
                                                                 values=class_situation, state='readonly',
                                                                 cursor="hand2",
                                                                 justify=CENTER,
                                                                 font=('times', 10))
                                        situation.set("CLASS")
                                        situation.place(x=850, y=50)
                                        title['text'] = 'ABSENTEES ATTENDANCE'
                                        proceed = Button(master, text="Proceed", fg='white', bg="#FF0051",
                                                         width=20, relief=FLAT)
                                        proceed.bind('<Button>',
                                                     lambda mouse: normal_time(mouse, no_people.get(), num, display))
                                        proceed.place(x=410, y=480)
                                        master.bind('<Return>',
                                                    lambda enter: normal_time(enter, no_people.get(), num, display))
                                        tick = '\u2705'
                                        nonlocal total_hour
                                        for no, _ in enumerate(time):
                                            if tick in _:
                                                time[no] = _.replace(tick, '')

                                        connection = os.path.join(
                                            excel_file_check(class_name, institute_name),
                                            f"{class_name} {month[days.month]}.xlsx")

                                        if os.path.exists(connection):
                                            try:
                                                excel = load_workbook(connection)
                                                fill = excel[f'{class_name} {get_date} {month[days.month]}']
                                                get_hour = fill['O16'].value
                                                if get_hour not in [None]:
                                                    total_hour = get_hour
                                                for roll, column in enumerate(table[:total_hour]):
                                                    if fill[f'{column}{2}'].value not in [None]:
                                                        if tick not in time[roll]:
                                                            time[roll] = time[roll] + ' ' + tick
                                            except KeyError:
                                                pass

                                        if total_hour == 0:
                                            period = sqlite3.connect(class_db)
                                            led = period.cursor()
                                            led.execute("SELECT number FROM  HOUR")
                                            total_hour = led.fetchone()[0]
                                            period.commit()
                                            period.close()

                                        class_hour['values'] = time[:total_hour]
                                    elif time_period == 'late':
                                        title['text'] = 'LATE_COMER ATTENDANCE'
                                        proceed = Button(master, text="Proceed", fg='white', bg="#FF0051", width=20,
                                                         relief=FLAT)
                                        proceed.bind('<Button>',
                                                     lambda mouse: late_time(mouse, no_people.get(), num))
                                        proceed.place(x=410, y=480)
                                        master.bind('<Return>',
                                                    lambda enter: late_time(enter, no_people.get(), num))
                                    master.mainloop()
                        attendance('od time', 'Everyone On Duty', 'NO On Duty', 120, [])
                else:
                    final_check(class_name, class_db, self.block, institute_db, institute_name)
                    system()


institution = Institute('window', 'institute', 'institute_db')


# sheet.sheet_view.showGridLines = False

# sheet.sheet_view.zoomScale = 80

# out_work_book.close()

# sheet.sheet_view.showGridLines = False

# sheet.sheet_view.zoomScale = 80

# out_work_book.close()

# It is used to delete a single record.

# To take action on the DATA WE provide in

# this is the main window to create the choose in th tkinter window
