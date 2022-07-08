from institute import *

# ============== section window ================#


class Section:

    def __init__(self, win, database, institution_name):
        self.win = win
        self.database = database
        self.institution_name = institution_name
        self.class_strength = total_strength

    def close(self, eve, back, class_name):
        if eve.char == '`' or eve.char == '??':
            back.destroy()
            class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')
            self.next_window(class_db, class_name)

    def back(self, back, class_name):
        mark = Button(back, text='BACK', bg='#67E1E4', fg='#FFFFFF', relief=FLAT, borderwidth=0,
                      cursor='hand2')
        mark.bind('<Button>', lambda mouse: self.close(mouse, back, class_name))
        mark.place(x=3, y=25)

    def next_window(self, class_db, class_name):
        if os.path.exists(class_db):
            self.section_handling(class_name, 'black')
        else:
            self.section_handling(constant_class,  '#666666')

    def section_handling(self, class_name, color):
        if os.path.exists(college_dbpath):
            total_check()
            if os.path.exists(self.database):
                sections = []
                connect5 = sqlite3.connect(self.database)
                pen3 = connect5.cursor()
                pen3.execute("SELECT * FROM classes")
                for s in pen3.fetchall():
                    for x in s:
                        sections.append(x)
                connect5.commit()
                connect5.close()
                if len(sections) != 0:
                    self.section_window(sections, class_name, color)
                else:
                    institution.institution_handling(self.institution_name, 'black')
            else:
                institution.institution_handling(constant, '#666666')
        else:
            system()

    def delete_total(self, event, class_name):
        if event.char == '' or event.char == '??':
            if 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Hey please select the class')
            else:
                if messagebox.askyesno("Confirmation", f'Are you sure to Delete Your Class --> {class_name}?'):
                    class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')

                    if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(self.database):
                        os.remove(class_db)
                        messagebox.showinfo('Confirmation',
                                            message=f'Successfully Deleted Class--> ' f'{class_name}')
                        self.win.destroy()
                        self.section_handling(constant_class, '#666666')
                    else:
                        final_check(class_name, class_db, self.win, self.database, self.institution_name)
                        self.section_handling(constant_class, '#666666')

    def delete_name(self, event, class_name):

        def delete_window(eve):
            if eve:
                if os.path.exists(class_db) and os.path.exists(
                        college_dbpath) and os.path.exists(self.database):
                    _id_ = []
                    student_name = delete_variable.get()
                    if len(student_name) == 0:
                        messagebox.showinfo(title="Remainder", message='Select a Student')
                    else:
                        if messagebox.askyesno('Confirmation',
                                               message=f'Are you sure to Delete {student_name}'):
                            if os.path.exists(class_db) and os.path.exists(
                                    college_dbpath) and os.path.exists(self.database):
                                _id_.append(student_name)
                                connect6 = sqlite3.connect(class_db)
                                pen4 = connect6.cursor()
                                pen4.execute("SELECT rowid FROM Student WHERE name = ?", _id_)
                                row_id = pen4.fetchone()[0]
                                for roll, name in enumerate(delete_name[row_id-1:], row_id):
                                    pen4.execute(
                                        "DELETE FROM Student WHERE rowid= " + str(roll))
                                else:
                                    for roll, name in enumerate(delete_name[row_id:], row_id):
                                        pen4.execute("INSERT INTO STUDENT  VALUES (?,?) ", (roll, name))
                                connect6.commit()
                                connect6.close()
                                messagebox.showinfo('Confirmation',
                                                    message=f"Successfully Deleted {student_name}")
                                root1.destroy()
                                self.next_window(class_db, class_name)
                            else:
                                final_check(class_name, class_db,  root1, self.database, self.institution_name)
                                self.section_handling(constant_class, '#666666')
                else:
                    final_check(class_name, class_db, root1, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')
        if event.char == 'd' or event.char == '??':
            if 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Hey please select the class')
            else:
                class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                        self.database):
                    my_class = self.class_strength(class_db)
                    if my_class == 0:
                        messagebox.showinfo(title="Remainder", message='You can only Add')
                    else:
                        window_proceed = True
                        show_status = True
                        for _ in range(1):
                            while True:
                                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                        self.database):
                                    connection = os.path.join(
                                        excel_file_check(class_name, self.institution_name),
                                        f"{class_name} {month[days.month]}.xlsx")
                                    try:
                                        if os.path.exists(connection):
                                            wb1 = load_workbook(connection)
                                            sheet1 = wb1[f'{class_name} {days.day} {month[days.month]}']
                                            present = list(
                                                filter(lambda a: sheet1[f'{a}{2}'].value not in [None],
                                                       hour_data(class_db)))
                                            if len(present) >= 1:
                                                window_proceed = False
                                                show_status = False
                                                messagebox.showinfo(title="Remainder",
                                                                    message=f"Name can't Delete During Attendance")
                                            wb1.save(connection)
                                    except PermissionError:
                                        if show_status:
                                            messagebox.showinfo("Reminder",
                                                                'Kindly close your opened excel file')
                                            window_proceed = False
                                            break
                                        else:
                                            break
                                    except KeyError:
                                        break
                                    else:
                                        break
                                else:
                                    final_check(class_name, class_db, self.win, self.database, self.institution_name)
                                    self.section_handling(constant_class, '#666666')
                                    break
                        else:
                            if window_proceed:
                                self.win.destroy()
                                root1 = Tk()
                                window(root1)
                                image = Image.open(background_image)
                                imageview = ImageTk.PhotoImage(image)
                                Label(root1, image=imageview).pack()
                                delete_variable = StringVar()
                                delete_name = students_names(class_db)
                                Label(root1, text=f"Class Name = {class_name}", bg='#67E1E4', fg='#FFFFFF',
                                      ).place(
                                    x=30, y=5)

                                def label_place():
                                    my_label.config(text=f"Selected Name ---> {delete_variable.get()} ")

                                delete_menu = Menubutton(root1, text="Roll-> NAME", relief=FLAT, fg='#666666',
                                                         bg="white",
                                                         width=60)
                                delete_menu.menu = Menu(delete_menu, tearoff=0)
                                delete_menu["menu"] = delete_menu.menu
                                for i, name_data in enumerate(delete_name, 1):
                                    delete_menu.menu.add_radiobutton(label=f'  {i} ---> {name_data}',
                                                                     variable=delete_variable,
                                                                     value=name_data,
                                                                     command=label_place)

                                delete_menu.place(x=350, y=200)
                                my_label = Label(root1, bg='#67E1E4', text='Selected Name ---> !!')
                                my_label.place(x=700, y=10)
                                proceed = Button(root1, text="Proceed", activebackground='#FF0051',
                                                 activeforeground='#F95289',
                                                 padx=40,
                                                 pady=10, relief=FLAT)
                                proceed.bind('<Button>', delete_window)
                                proceed.place(x=400, y=450)

                                root1.bind('<Key>', lambda key: self.close(key, root1, class_name))
                                self.back(root1, class_name)
                                root1.bind('<Return>', delete_window)
                                root1.mainloop()
                else:
                    final_check(class_name, class_db,  self.win, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

    def modify(self, event, class_name):

        def process(eve):
            if eve:
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(self.database):
                    modify_name = entry_variable.get().strip().upper()
                    old_roll = modify_variable.get()
                    if old_roll == 0 and ("EX: BABAI SUNDARARAJAN" in modify_name or len(modify_name) == 0):
                        messagebox.showinfo(title="Remainder",
                                            message='Give roll number and Name to process')

                    elif old_roll >= 1 and ("EX: BABAI SUNDARARAJAN" in modify_name or len(modify_name) == 0):
                        messagebox.showinfo(title="Remainder",
                                            message='Give Name to process')

                    elif old_roll == 0 and len(modify_name) >= 1:
                        messagebox.showinfo(title="Remainder", message='Select the roll number ')
                    else:
                        if name_quality_checker(modify_name):
                            check_delete = remove_white_space(modify_name)
                            old_name = modify_names[old_roll - 1]
                            if check_delete != remove_white_space(old_name):
                                if check_delete == 'NONE':
                                    modify_name = 'NONE'
                                if check_delete == 'DELETED':
                                    modify_name = 'DELETED'
                                if messagebox.askyesno("Confirmation",
                                                       f'Are you sure to Modify {old_name} --> {modify_name}'):
                                    for _ in range(1):
                                        while True:
                                            if os.path.exists(class_db) and os.path.exists(
                                                    college_dbpath) and os.path.exists(self.database):
                                                connection = os.path.join(
                                                    excel_file_check(class_name, self.institution_name),
                                                    f"{class_name} {month[days.month]}.xlsx")
                                                try:
                                                    if os.path.exists(connection):
                                                        wb1 = load_workbook(connection)
                                                        sheet1 = wb1[f'{class_name} {days.day} {month[days.month]}']
                                                        modify_names[old_roll - 1] = modify_name
                                                        sheet_data(sheet1, modify_names, hour_data(class_db))
                                                        wb1.save(connection)
                                                except PermissionError:
                                                    messagebox.showinfo("Reminder",
                                                                        'Kindly close your opened excel file')
                                                except KeyError:
                                                    break
                                                else:
                                                    break
                                            else:
                                                break
                                    else:
                                        _id_ = []
                                        if os.path.exists(class_db) and os.path.exists(
                                                college_dbpath) and os.path.exists(self.database):
                                            _id_.append(old_name)
                                            p1 = sqlite3.connect(class_db)
                                            cur = p1.cursor()
                                            cur.execute("UPDATE Student SET name = ? WHERE rowid = ?",
                                                        (modify_name, old_roll))
                                            p1.commit()
                                            p1.close()
                                            messagebox.showinfo('Confirmation',
                                                                message=f"Hey Successfully Modified "
                                                                        f"{old_name} -->"
                                                                        f" {modify_name}")
                                            root2.destroy()
                                            self.next_window(class_db, class_name)
                                        else:
                                            final_check(class_name, class_db, root2, self.database,
                                                        self.institution_name)
                                            self.section_handling(constant_class, '#666666')
                            else:
                                messagebox.showinfo(title="Remainder",
                                                    message=f"Same Name can't Modify {old_name} -->"
                                                            f" {modify_name}")
                        else:
                            messagebox.showinfo(title="Remainder",
                                                message=f"Name can't be Special characters and Numbers ")
                else:
                    final_check(class_name, class_db, root2, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

        if event.char == 'm' or event.char == '??':
            if 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Hey please select the class')
            else:
                class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                        self.database):
                    my_class = self.class_strength(class_db)
                    if my_class == 0:
                        messagebox.showinfo(title="Remainder", message='You can only Add')
                    else:
                        self.win.destroy()
                        root2 = Tk()
                        window(root2)
                        image = Image.open(background_image)
                        imageview = ImageTk.PhotoImage(image)
                        Label(root2, image=imageview).pack()
                        modify_names = students_names(class_db)
                        entry_variable = StringVar()
                        modify_variable = IntVar()
                        Label(root2, text=f"Class Name = {class_name}", bg='#67E1E4', fg='#FFFFFF').place(
                            x=30, y=5)

                        def label_place():
                            my_label.config(text=f"Selected Name ---> {modify_names[modify_variable.get() - 1]} ")

                        student_name = Entry(root2, fg='#666666', bg="white", textvariable=entry_variable, width=60,
                                             relief=FLAT)
                        student_name.insert(0, "Ex: BABAI Sundararajan")
                        clicked = student_name.bind('<Button>', lambda eve: click(student_name, eve, clicked))
                        student_name.config(state='readonly')
                        student_name.place(x=350, y=300)
                        modify_menu = Menubutton(root2, text="Roll-> NAME", relief=FLAT, fg='#666666', bg="white",
                                                 width=60)
                        modify_menu.menu = Menu(modify_menu, tearoff=0)
                        modify_menu["menu"] = modify_menu.menu
                        for i, name_data in enumerate(modify_names, 1):
                            modify_menu.menu.add_radiobutton(label=f'  {i} ---> {name_data}', variable=modify_variable,
                                                             value=i,
                                                             command=label_place)
                        modify_menu.place(x=350, y=200)
                        my_label = Label(root2, bg='#67E1E4', text='Selected Name ---> !!')
                        my_label.place(x=700, y=10)
                        proceed = Button(root2, text="Proceed", activebackground='#FF0051', activeforeground='#F95289',
                                         padx=40,
                                         pady=10, relief=FLAT)
                        proceed.bind('<Button>', process)
                        proceed.place(x=400, y=450)
                        root2.bind('<Key>', lambda key: self.close(key, root2, class_name))
                        self.back(root2, class_name)
                        root2.bind('<Return>', process)
                        root2.mainloop()
                else:
                    final_check(class_name, class_db, self.win, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

    def add(self, event, class_name):
        def add_process(eve):
            if eve:
                if os.path.exists(class_db) and os.path.exists(
                                                    college_dbpath) and os.path.exists(
                                                    self.database):
                    add_name = add_variable.get().strip().upper()
                    if "Ex: Keerthi Raja" in add_name or len(add_name) == 0:
                        messagebox.showinfo(title="Remainder", message="Your Entry is Missing")
                    else:
                        if name_quality_checker(add_name):
                            roll_no = my_class + 1
                            none_delete = remove_white_space(add_name)
                            if none_delete == 'NONE':
                                add_name = 'NONE'
                            if none_delete == 'DELETED':
                                add_name = 'DELETED'
                            if messagebox.askyesno('Confirmation',
                                                   message=f'Are you sure to Add {add_name} '
                                                           f'at Roll number --> {roll_no} '):
                                add_student.append(add_name)
                                for _ in range(1):
                                    while True:
                                        if os.path.exists(class_db) and os.path.exists(
                                                college_dbpath) and os.path.exists(self.database):
                                            connection = os.path.join(
                                                excel_file_check(class_name, self.institution_name),
                                                f"{class_name} {month[days.month]}.xlsx")
                                            try:
                                                if os.path.exists(connection):
                                                    wb = load_workbook(connection)
                                                    sheet1 = wb[f'{class_name} {days.day} {month[days.month]}']
                                                    sheet_data(sheet1, add_student, hour_data(class_db))
                                                    wb.save(connection)
                                            except PermissionError:
                                                messagebox.showinfo("Reminder",
                                                                    'Kindly close your opened excel file')
                                            except KeyError:
                                                break
                                            else:
                                                break
                                        else:
                                            break
                                else:
                                    if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                            self.database):
                                        p1 = sqlite3.connect(class_db)
                                        cur1 = p1.cursor()
                                        cur1.execute("INSERT INTO STUDENT  VALUES (?,?) ", (roll_no, add_name))
                                        p1.commit()
                                        p1.close()
                                        messagebox.showinfo('confirmation',
                                                            message=f'Successfully Added {add_name} '
                                                                    f'at Roll number --> {roll_no}')
                                        root3.destroy()
                                        self.next_window(class_db, class_name)
                                    else:
                                        final_check(class_name, class_db, root3, self.database,
                                                    self.institution_name)
                                        self.section_handling(constant_class, '#666666')
                        else:
                            messagebox.showinfo(title="Remainder",
                                                message=f"Name can't be Special characters and Numbers ")
                else:
                    final_check(class_name, class_db, root3, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

        if event.char == 'a' or event.char == '??':
            if 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Hey please select the class')
            else:
                class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                        self.database):
                    my_class = self.class_strength(class_db)
                    if my_class >= 100:
                        messagebox.showinfo(title="Remainder",
                                            message=f"Hey can't insert more than 100 Names")
                    else:
                        self.win.destroy()
                        root3 = Tk()
                        window(root3)
                        image = Image.open(background_image)
                        imageview = ImageTk.PhotoImage(image)
                        Label(root3, image=imageview).pack()
                        add_student = students_names(class_db)
                        Label(root3, text=f"Class Name = {class_name}", bg='#67E1E4', fg='#FFFFFF').place(x=30, y=5)
                        add_variable = StringVar()
                        Label(root3, text="Good Name --> ", bg='#67E1E4', fg='#FFFFFF', relief=FLAT).place(x=240, y=300)
                        Label(root3, text=f"Total Strength = {my_class}", bg='#67E1E4', fg='#FFFFFF',
                              relief=FLAT).place(x=730, y=10)
                        student_name = Entry(root3, fg='#666666', bg="white", textvariable=add_variable, width=60,
                                             relief=FLAT)
                        student_name.insert(0, "Ex: Keerthi Raja")
                        clicked = student_name.bind('<Button>', lambda eve: click(student_name, eve, clicked))
                        student_name.config(state='readonly')
                        student_name.place(x=350, y=300)
                        proceed = Button(root3, text="PROCEED", bg='#67E1E4', fg='#FFFFFF', activebackground='#67E1E4',
                                         activeforeground='#FFFFFF', relief=SOLID, width=10, font='BOLD')
                        proceed.bind('<Button>', add_process)
                        proceed.place(x=400, y=450)
                        root3.bind('<Key>', lambda key: self.close(key, root3, class_name))
                        self.back(root3, class_name)
                        root3.bind('<Return>', add_process)
                        root3.mainloop()
                else:
                    final_check(class_name, class_db, self.win, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

    def insert(self, event, class_name):

        def process(eve):
            if eve:
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                        self.database):
                    insert_name = entry_variable.get().strip().upper()
                    roll_insert = insert_variable.get()
                    if roll_insert == 0 and ("Ex: Keerthi Raja" in insert_name or len(insert_name) == 0):
                        messagebox.showinfo(title="Remainder",
                                            message='Give roll number and Name to process')

                    elif roll_insert >= 1 and ("Ex: Keerthi Raja" in insert_name or len(insert_name) == 0):
                        messagebox.showinfo(title="Remainder",
                                            message='Give Name to process')

                    elif roll_insert == 0 and len(insert_name) >= 1:
                        messagebox.showinfo(title="Remainder", message='Select the roll number ')
                    else:
                        if name_quality_checker(insert_name):
                            if messagebox.askyesno("Confirmation",
                                                   f'Are you sure to insert {insert_name} --> {roll_insert}'):
                                none_delete = remove_white_space(insert_name)
                                if none_delete == 'NONE':
                                    insert_name = 'NONE'
                                if none_delete == 'DELETED':
                                    insert_name = 'DELETED'
                                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                        self.database):
                                    student_insert_name.insert(roll_insert - 1, insert_name)
                                    insert = sqlite3.connect(class_db)
                                    rubber = insert.cursor()
                                    for roll, name in enumerate(student_insert_name[roll_insert - 1:], roll_insert):
                                        if roll > my_class:
                                            rubber.execute("INSERT INTO STUDENT  VALUES (?,?) ",
                                                           (roll, name))
                                        else:
                                            rubber.execute("UPDATE Student SET name = ? WHERE rowid = ?",
                                                           (name, roll))
                                    else:
                                        insert.commit()
                                        insert.close()
                                        messagebox.showinfo('Confirmation',
                                                            message=f"Successfully Inserted {insert_name} --> "
                                                                    f"{roll_insert}")
                                        root4.destroy()
                                        self.next_window(class_db, class_name)
                                else:
                                    final_check(class_name, class_db, root4, self.database, self.institution_name)
                                    self.section_handling(constant_class, '#666666')
                        else:
                            messagebox.showinfo(title="Remainder",
                                                message=f"Name can't be Special characters and Numbers ")
                else:
                    final_check(class_name, class_db, root4, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

        if event.char == 'i' or event.char == '??':
            if 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Select your class')
            else:
                class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                        self.database):
                    my_class = self.class_strength(class_db)
                    if my_class == 0:
                        messagebox.showinfo(title="Remainder",
                                            message=f"You can only add Names")
                    elif my_class >= 100:
                        messagebox.showinfo(title="Remainder",
                                            message=f"Hey can't insert more than 100 Names")
                    else:
                        window_proceed = True
                        show_status = True
                        for _ in range(1):
                            while True:
                                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                        self.database):
                                    connection = os.path.join(
                                        excel_file_check(class_name, self.institution_name),
                                        f"{class_name} {month[days.month]}.xlsx")
                                    try:
                                        if os.path.exists(connection):
                                            wb1 = load_workbook(connection)
                                            sheet1 = wb1[f'{class_name} {days.day} {month[days.month]}']
                                            present = list(
                                                filter(lambda a: sheet1[f'{a}{2}'].value not in [None],
                                                       hour_data(class_db)))
                                            if len(present) >= 1:
                                                window_proceed = False
                                                show_status = False
                                                messagebox.showinfo("Reminder",
                                                                    'Name can\'t insert During Attendance')
                                            wb1.save(connection)
                                    except PermissionError:
                                        if show_status:
                                            messagebox.showinfo("Reminder", 'Kindly close your opened excel file')
                                            window_proceed = False
                                            break
                                        else:
                                            break
                                    except KeyError:
                                        break
                                    else:
                                        break
                                else:
                                    final_check(class_name, class_db,  self.win, self.database, self.institution_name)
                                    self.section_handling(constant_class, '#666666')
                                    break
                        else:
                            if window_proceed:
                                self.win.destroy()
                                root4 = Tk()
                                window(root4)
                                image = Image.open(background_image)
                                imageview = ImageTk.PhotoImage(image)
                                Label(root4, image=imageview).pack()
                                Label(root4, text=f"Class Name = {class_name}", bg='#67E1E4', fg='#FFFFFF').place(
                                    x=30, y=5)
                                student_insert_name = students_names(class_db)
                                entry_variable = StringVar()
                                insert_variable = IntVar()

                                def label_place():
                                    my_label.config(
                                        text=f"Selected Name ---> {student_insert_name[insert_variable.get() - 1]} ")

                                student_name = Entry(root4, fg='#666666', bg="white", textvariable=entry_variable,
                                                     width=60,
                                                     relief=FLAT)
                                student_name.insert(0, "Ex: Keerthi Raja")
                                clicked = student_name.bind('<Button>', lambda eve: click(student_name, eve, clicked))

                                student_name.place(x=350, y=300)
                                insert_menu = Menubutton(root4, text="Roll-> NAME", relief=FLAT, fg='#666666',
                                                         bg="white",
                                                         width=60)
                                insert_menu.menu = Menu(insert_menu, tearoff=0)
                                insert_menu["menu"] = insert_menu.menu
                                for i, name_data in enumerate(student_insert_name, 1):
                                    insert_menu.menu.add_radiobutton(label=f'  {i} ---> {name_data}',
                                                                     variable=insert_variable,
                                                                     value=i,
                                                                     command=label_place)

                                insert_menu.place(x=350, y=200)
                                my_label = Label(root4, bg='#67E1E4', text='Selected Name ---> !!')
                                my_label.place(x=700, y=10)
                                proceed = Button(root4, text="PROCEED", bg='#67E1E4', fg='#FFFFFF',
                                                 activebackground='#67E1E4',
                                                 activeforeground='#FFFFFF', relief=SOLID, width=10, font='BOLD')
                                proceed.bind('<Button>', process)
                                proceed.place(x=400, y=450)
                                root4.bind('<Key>', lambda key: self.close(key, root4, class_name))
                                self.back(root4, class_name)
                                root4.bind('<Return>', process)
                                root4.mainloop()
                else:
                    final_check(class_name, class_db, self.win, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

    def class_hour(self, event, class_name):

        def sin(eve):
            if eve:
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                        self.database):
                    new_hour = hour_variable.get()
                    if new_hour == 0:
                        messagebox.showinfo("Remainder", 'Choose your Hour')
                    else:
                        if old_hour != new_hour:
                            if messagebox.askyesno('Confirmation',
                                                   message=f'Are you sure to Modify {old_hour} --> {new_hour}'):
                                for _ in range(1):
                                    while True:
                                        if os.path.exists(class_db) and os.path.exists(
                                                college_dbpath) and os.path.exists(self.database):
                                            try:
                                                if os.path.exists(connection):
                                                    wb1 = load_workbook(connection)
                                                    sheet1 = wb1[f'{class_name} {days.day} {month[days.month]}']
                                                    for fg in table:
                                                        sheet1[f'{fg}{1}'].value = ''
                                                        sheet1[f'{fg}{1}'].fill = white
                                                        sheet1[f'{fg}{1}'].alignment = Alignment(vertical='center',
                                                                                                 horizontal='center')
                                                        sheet1[f'{fg}{1}'].border = no_border
                                                    else:
                                                        for roll1, d in enumerate(table[:new_hour]):
                                                            sheet1[f'{d}{1}'].value = time[roll1]
                                                            sheet1[f'{d}{1}'].alignment = Alignment(vertical='center',
                                                                                                    horizontal='center')
                                                            sheet1[f'{d}{1}'].border = style
                                                            sheet1[f'{d}{1}'].font = Font(color='FFFFFF', bold=True)
                                                            sheet1[f'{d}{1}'].fill = green
                                                        else:
                                                            for fg1 in table[new_hour:]:
                                                                for index in range(2, 102):
                                                                    if sheet1[f'{fg1}{index}'].value not in [None]:
                                                                        sheet1[f'{fg1}{index}'].value = ''
                                                                        sheet1[f'{fg1}{index}'].fill = white
                                                            else:
                                                                wb1.save(connection)
                                            except PermissionError:
                                                messagebox.showinfo("Reminder", 'Kindly close your opened excel file')
                                            except KeyError:
                                                break
                                            else:
                                                break
                                        else:
                                            break
                                else:
                                    if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                            self.database):
                                        change_hour = sqlite3.connect(class_db)
                                        rubber = change_hour.cursor()
                                        rubber.execute("UPDATE HOUR SET number = ? WHERE rowid = ?", (new_hour, 1))
                                        change_hour.commit()
                                        change_hour.close()
                                        messagebox.showinfo('confirmation',
                                                            message=f'Successfully Modified {old_hour} --> {new_hour}')
                                        root5.destroy()
                                        self.next_window(class_db, class_name)
                                    else:
                                        final_check(class_name, class_db, root5, self.database, self.institution_name)
                                        self.section_handling(constant_class, '#666666')
                        else:
                            messagebox.showinfo(title="Remainder",
                                                message=f"Hey you can't Modify {old_hour} -->"
                                                        f" {new_hour}")
                else:
                    final_check(class_name, class_db, root5, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

        if event.char == 'h' or event.char == '??':
            if 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Select your class')
            else:
                class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                        self.database):
                    connection = os.path.join(
                        excel_file_check(class_name, self.institution_name),
                        f"{class_name} {month[days.month]}.xlsx")
                    self.win.destroy()
                    root5 = Tk()
                    window(root5)
                    image = Image.open(background_image)
                    imageview = ImageTk.PhotoImage(image)
                    Label(root5, image=imageview).pack()
                    hour_variable = IntVar()
                    period3 = sqlite3.connect(class_db)
                    led = period3.cursor()
                    led.execute("SELECT number FROM  HOUR")
                    old_hour = led.fetchall()[0][0]
                    period3.commit()
                    period3.close()
                    Label(root5, text=f'Present Class Hour -->{old_hour}', fg='#d902ee', bg="#ffcccc").place(x=100,
                                                                                                             y=10)

                    def label_place():
                        my_label.config(
                            text=f"Selected Hour ---> {hour_variable.get()} ")
                    hour_menu = Menubutton(root5, text="Class Hour", relief=FLAT, fg='#666666', bg="white",
                                           width=60)
                    hour_menu.menu = Menu(hour_menu, tearoff=0)
                    hour_menu["menu"] = hour_menu.menu
                    for i in range(1, 11):
                        hour_menu.menu.add_radiobutton(label=f'{i}', variable=hour_variable, value=i,
                                                       command=label_place)
                    hour_menu.place(x=350, y=200)
                    my_label = Label(root5, bg='#67E1E4', text='Selected Hour ---> !!')
                    my_label.place(x=700, y=10)
                    proceed = Button(root5, text="PROCEED", bg='#67E1E4', fg='#FFFFFF',
                                     activebackground='#67E1E4',
                                     activeforeground='#FFFFFF', relief=SOLID, width=10, font='BOLD')
                    proceed.bind('<Button>', sin)
                    proceed.place(x=400, y=450)
                    root5.bind('<Key>', lambda key: self.close(key, root5, class_name))
                    self.back(root5, class_name)
                    root5.bind('<Return>', sin)
                    root5.mainloop()
                else:
                    final_check(class_name, class_db, self.win, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

    def excel_clear(self, event, class_name):

        def _excel_(eve):
            if eve:
                _name_ = excel_variable.get()
                if "SELECT CLASS" in _name_:
                    messagebox.showinfo(title="Remainder",
                                        message=f"SELECT CLASS")
                else:
                    if messagebox.askyesno(title="Remainder",
                                           message=f"Are you sure to clear {_name_} "):
                        class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')
                        if os.path.exists(excel_folder):
                            temp = os.path.join(parent_dir, '__M__')
                            try:
                                os.rename(excel_folder, temp)
                            except OSError:
                                messagebox.showinfo("Reminder", 'Kindly close your Folder')
                            else:
                                os.rename(temp, excel_folder)
                                for institute in os.listdir(excel_folder):
                                    institute_folder = os.path.join(excel_folder, institute)
                                    for _class_ in os.listdir(institute_folder):
                                        if _name_ == _class_:
                                            cls_folder = os.path.join(institute_folder, _class_)
                                            for _year_ in os.listdir(cls_folder):
                                                year_folder = os.path.join(cls_folder, _year_)
                                                for excel in os.listdir(year_folder):
                                                    os.remove(os.path.join(year_folder, excel))
                                                else:
                                                    os.rmdir(year_folder)
                                            else:
                                                os.rmdir(cls_folder)
                                else:
                                    messagebox.showinfo(title="Remainder",
                                                        message=f"Successfully cleared the deleted folder ")
                                    excel_window.destroy()
                                    self.next_window(class_db, class_name)
                        else:
                            messagebox.showinfo(title="Remainder",
                                                message=f"You deleted the excel folder")
                            excel_window.destroy()
                            self.next_window(class_db, class_name)

        if event.char == 'x' or event.char == '??':
            if os.path.exists(excel_folder):
                excel_data = []
                for get_excel in os.listdir(excel_folder):
                    for _folder_ in os.listdir(os.path.join(excel_folder, get_excel)):
                        if '#Deleted' in _folder_:
                            excel_data.append(_folder_)
                if len(excel_data) != 0:
                    self.win.destroy()
                    excel_window = Tk()
                    window(excel_window)
                    image = Image.open(background_image)
                    imageview = ImageTk.PhotoImage(image)
                    Label(excel_window, image=imageview).pack()
                    excel_variable = StringVar()
                    strength_box = ttk.Combobox(excel_window, textvariable=excel_variable, values=excel_data,
                                                state='readonly', justify='center', width=60,
                                                font=('times', 10))
                    strength_box.set('---- SELECT CLASS ----')
                    strength_box.place(x=350, y=200)
                    proceed = Button(excel_window, text="PROCEED", bg='#67E1E4', fg='#FFFFFF',
                                     activebackground='#67E1E4',
                                     activeforeground='#FFFFFF', relief=SOLID, width=10, font='BOLD')
                    proceed.bind('<Button>', _excel_)
                    proceed.place(x=400, y=450)
                    excel_window.bind('<Key>', lambda key: self.close(key, excel_window, class_name))
                    self.back(excel_window, class_name)
                    excel_window.bind('<Return>', _excel_)
                    excel_window.mainloop()
                else:
                    messagebox.showinfo(title="Remainder", message=f"Nothing is there to clear")
            else:
                messagebox.showinfo(title="Remainder", message=f"You need to take Attendance")

    def modify_class(self, event, class_name):
        def change_class(eve):
            if eve:
                new_name = name_data.get().strip().upper()
                proceed = False
                if len(new_name) == 0:
                    messagebox.showinfo(title="Remainder", message="Your Entry is Missing")
                elif remove_white_space(new_name) == remove_white_space(class_name):
                    messagebox.showinfo(title="Remainder", message="No Changes has been done")
                elif new_name.isnumeric():
                    messagebox.showinfo(title="Remainder", message=f"Class Name can't be only Numbers")
                else:
                    if messagebox.askyesno('Confirmation',
                                           message=f'Are you sure with the Class Name --> '
                                                   f'{new_name.upper()}?'):
                        if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(
                                self.database):
                            _id_ = []
                            for _ in range(1):
                                for s in range(1):
                                    total_check()
                                    connect5 = sqlite3.connect(self.database)
                                    pen3 = connect5.cursor()
                                    pen3.execute("SELECT * FROM classes")
                                    for _name_ in pen3.fetchall():
                                        for _ in _name_:
                                            if remove_white_space(_) == remove_white_space(new_name):
                                                proceed = True
                                    connect5.commit()
                                    connect5.close()
                                else:
                                    if proceed:
                                        messagebox.showinfo(title="Remainder", message='Class Name already Exits')
                                    else:
                                        org_folder = os.path.join(excel_folder, self.institution_name)
                                        cla_folder = os.path.join(org_folder, class_name)
                                        if os.path.exists(excel_folder) and os.path.exists(
                                                org_folder) and os.path.exists(cla_folder):
                                            temp_folder = os.path.join(excel_folder, '_temp_')
                                            try:
                                                os.rename(org_folder, temp_folder)
                                            except OSError:
                                                messagebox.showinfo("Reminder", 'Kindly close your Folder')
                                                break
                                            else:
                                                os.rename(temp_folder, org_folder)
                                                for folder_institute in os.listdir(excel_folder):
                                                    if remove_white_space(self.institution_name) == remove_white_space(
                                                            folder_institute):
                                                        original_ins = os.path.join(excel_folder, folder_institute)
                                                        for class_fold in os.listdir(original_ins):
                                                            if remove_white_space(class_name) == remove_white_space(
                                                                    class_fold):
                                                                original_class = os.path.join(original_ins, class_fold)
                                                                for class_year in os.listdir(original_class):
                                                                    original_year = os.path.join(original_class,
                                                                                                 class_year)
                                                                    for excel in os.listdir(original_year):
                                                                        os.rename(os.path.join(original_year, excel),
                                                                                  os.path.join(original_year,
                                                                                               excel.replace(class_name,
                                                                                                             new_name)))
                                                                    else:
                                                                        remake_year = os.path.join(original_class,
                                                                                                   class_year.replace(
                                                                                                       class_name,
                                                                                                       new_name))
                                                                        os.rename(original_year, remake_year)
                                                                else:
                                                                    remake_class = os.path.join(original_ins,
                                                                                                class_fold.replace(
                                                                                                    class_name,
                                                                                                    new_name))
                                                                    os.rename(original_class, remake_class)
                            else:
                                _id_.append(class_name)
                                open_class = sqlite3.connect(self.database)
                                cursor = open_class.cursor()
                                cursor.execute("SELECT rowid FROM classes WHERE name = ?", _id_)
                                for w in cursor.fetchall():
                                    for rowid in w:
                                        cursor.execute("UPDATE classes SET name = ? WHERE rowid = ?",
                                                       (new_name, rowid))
                                open_class.commit()
                                open_class.close()
                                modify_db = os.path.join(section_folder, f'{new_name} {self.institution_name}.db')
                                os.rename(class_db, modify_db)
                                messagebox.showinfo("Reminder", 'Successfully updated your Class')
                                change_window.destroy()
                                self.next_window(modify_db, new_name)
                        else:
                            final_check(class_name, class_db, change_window, self.database, self.institution_name)
                            self.section_handling(constant_class, '#666666')

        if event.char == 'm' or event.char == '??':
            if 'Select Class' in class_name:
                messagebox.showinfo(title="Remainder", message='Hey please select the class')
            else:
                class_db = os.path.join(section_folder, f'{class_name} {self.institution_name}.db')
                if os.path.exists(class_db) and os.path.exists(college_dbpath) and os.path.exists(self.database):
                    self.win.destroy()
                    change_window = Tk()
                    window(change_window)
                    image = Image.open(background_image)
                    imageview = ImageTk.PhotoImage(image)
                    Label(change_window, image=imageview).pack()
                    name_data = StringVar()
                    entry = Entry(change_window, fg='black', bg="white", textvariable=name_data, relief=FLAT,
                                  width=60)
                    entry.insert(0, class_name)
                    entry.place(x=350, y=300)
                    confirm = Button(change_window, text="PROCEED", bg='#67E1E4', fg='#FFFFFF',
                                     activebackground='#67E1E4',
                                     activeforeground='#FFFFFF', relief=SOLID, width=10, font='BOLD')

                    confirm.bind('<Button>', change_class)
                    confirm.place(x=400, y=450)
                    change_window.bind('<Key>', lambda key: self.close(key, change_window, class_name))
                    self.back(change_window, class_name)
                    change_window.bind('<Return>', change_class)
                    change_window.mainloop()
                else:
                    final_check(class_name, class_db, self.win, self.database, self.institution_name)
                    self.section_handling(constant_class, '#666666')

    def section_window(self, sections, class_name, color):
        new_window = Tk()
        window(new_window)
        image = Image.open(background_image)
        imageview = ImageTk.PhotoImage(image)
        Label(new_window, image=imageview).pack()
        self.win = new_window
        ttk.Style().configure('TCombobox', foreground=color)
        section_name = StringVar()
        section_box = ttk.Combobox(new_window, textvariable=section_name, state='readonly', font=('times', 10),
                                   width=54,
                                   justify='center',
                                   values=sections)
        section_box.set(class_name)
        section_box.place(x=345, y=20)

        def close(eve):
            if eve.char == '`' or eve.char == '??':
                self.win.destroy()
                if os.path.exists(self.database):
                    institution.institution_handling(self.institution_name, 'black')
                else:
                    institution.institution_handling(constant, '#666666')

        mark = Button(new_window, text='BACK', fg='#FFFFFF',
                      bg="#FF0051", relief=FLAT,
                      cursor='hand2')
        mark.bind('<Button>', close)
        mark.place(x=0, y=23)
        add = Button(new_window, text="ADD", bg='#FF0051', fg='#FFFFFF', activebackground='#FF0051',
                     activeforeground='#FFFFFF', padx=80, pady=10, font='bold', relief=FLAT)
        add.bind('<Button>', lambda mouse: self.add(mouse, section_name.get()))
        add.place(x=680, y=300)
        insert = Button(new_window, text="INSERT", bg='#FF0051', fg='#FFFFFF', activebackground='#FF0051',
                        activeforeground='#FFFFFF', padx=40, pady=10, font='bold', relief=FLAT)
        insert.bind('<Button>', lambda mouse: self.insert(mouse, section_name.get()))
        insert.place(x=400, y=150)
        delete = Button(new_window, text="DELETE", bg='#FF0051', fg='#FFFFFF', activebackground='#FF0051',
                        activeforeground='#FFFFFF', font='bold', padx=67, pady=10, relief=FLAT)
        delete.bind('<Button>', lambda mouse: self.delete_name(mouse, section_name.get()))
        delete.place(x=120, y=150)
        modify = Button(new_window, text="MODIFY", bg='#FF0051', fg='#FFFFFF', activebackground='#FF0051',
                        activeforeground='#FFFFFF', padx=68, pady=10, font='bold', relief=FLAT)
        modify.bind('<Button>', lambda mouse: self.modify(mouse, section_name.get()))
        modify.place(x=680, y=150)
        hour = Button(new_window, text="HOUR", bg='#FF0051', fg='#FFFFFF', activebackground='#FF0051',
                      activeforeground='#FFFFFF', padx=40, pady=10, font='bold', relief=FLAT)
        hour.bind('<Button>', lambda mouse: self.class_hour(mouse, section_name.get()))
        hour.place(x=680, y=450)
        class_clear = Button(new_window, text="CLEAR", bg='#FF0051', fg='#FFFFFF', activebackground='#FF0051',
                             activeforeground='#FFFFFF', padx=40, pady=10, font='bold', relief=FLAT)
        class_clear.bind('<Button>', lambda mouse: self.excel_clear(mouse, section_name.get()))
        class_clear.place(x=120, y=450)
        delete_class = Button(new_window, text="DELETE CLASS", bg='#FF0051', fg='#FFFFFF', activebackground='#FF0051',
                              activeforeground='#FFFFFF', padx=39, pady=10, font='bold', relief=FLAT)
        delete_class.bind('<Button>', lambda mouse: self.delete_total(mouse, section_name.get()))
        delete_class.place(x=120, y=300)
        change_class = Button(new_window, text="MODIFY CLASS", bg='#FF0051', fg='#FFFFFF', activebackground='#FF0051',
                              activeforeground='#FFFFFF', padx=40, pady=10, font='bold', relief=FLAT)
        change_class.bind('<Button>', lambda mouse: self.modify_class(mouse, section_name.get()))
        change_class.place(x=400, y=300)
        new_window.bind('<Key>', lambda key: self.add(key, section_name.get()) or self
                        .insert(key, section_name.get()) or self.delete_name(
            key, section_name.get()) or self.modify(key, section_name.get()) or self
                        .class_hour(key, section_name.get()) or self.excel_clear(
            key, section_name.get()) or self.delete_total(key, section_name.get()) or self
                        .modify_class(key, section_name.get()) or close(key))
        new_window.mainloop()


class_window = Section('window', 'institute_db', 'institute')
