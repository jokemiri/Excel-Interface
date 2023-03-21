import tkinter as tk
from tkinter import ttk
import openpyxl
import pathlib

root = tk.Tk()

def load_data():
    path = ""
    xl_workbook = openpyxl.load_workbook(path)
    sheet = xl_workbook.active

    list_values = list(sheet.values)
    print(list_values)
    for column in list_values[0]:
        column.heading(column, text=column)

    for value_tuple in list_values[1:]:
        tree_view.insert('', tk.END, values=value_tuple)
# file = pathlib.Path('Master-Database.xlsx')
# if file.exists():
#     pass
# else:
#     file= workbook.write('M')
#     sheet=file.active
#     sheet['A1'] = "Name"
#     sheet['B1'] = "Age"
#     sheet['C1'] = "Nationality"
#     sheet['D1'] = "Staff"


#     file.save('Master-Database.xlsx')


style = ttk.Style(root)
root.tk.call("source", "Tkinter-Excel-Interface/forest-light.tcl")
root.tk.call("source", "Tkinter-Excel-Interface/forest-dark.tcl")
style.theme_use("forest-dark")

country_list = open('Tkinter-Excel-Interface/country_list.txt', encoding="utf-8")
# country_list.close()

form_frame = ttk.Frame(root)
form_frame.pack()

widget_frame = ttk.LabelFrame(form_frame, text='Insert Row')
widget_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(widget_frame)
name_entry.insert(0, 'Name')
name_entry.bind("<FocusIn>", lambda clear_name: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, sticky='ew', padx=5, pady=(0, 5))

age_spinbox = ttk.Spinbox(widget_frame, from_=18, to=100)
age_spinbox.grid(row=1, column=0, sticky='ew', padx=5, pady=5)
age_spinbox.insert(0, 'Age')

nationality = ttk.Combobox(widget_frame, values=country_list.readlines())
nationality.insert(0, 'Nationality')
# nationality.bind("<FocusIn>", lambda clear: nationality.delete('0', 'end'))
nationality.grid(row=2, column=0, sticky='ew', padx=5, pady=5)


def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    nation = nationality.get()
    staff_check = "Yes" if staff.get() else "No"
    
    print(name, age, nation, staff_check)
    
    #insert row into Excel Sheet
    path = "Master-Database.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, nation, staff]
    sheet.append(row_values)
    workbook.save(path)

    #Insert into treeview
    tree_view.insert('', tk.END, values=row_values)

    #clear values
    name_entry.delete(0, 'end')
    name_entry.insert(0, 'Name')
    age_spinbox.delete(0, 'end')
    age_spinbox.insert(0, 'Age')
    nationality.delete(0, 'end')
    nationality.set(nationality[0])
    staff.set("!selected")



def toggle_LD():
    if toggle_switch.instate(['selected']):
        style.theme_use('forest-light')
    else:
        style.theme_use('forest-dark')

staff = tk.BooleanVar
staff_check = ttk.Checkbutton(widget_frame, text="Staff", variable=staff)
staff_check.grid(row=3, column=0, sticky='nsew')

save_button = ttk.Button(widget_frame, text='Save', command=insert_row)
save_button.grid(row=4, column=0, sticky='nsew', padx=5, pady=5)

separator = ttk.Separator(widget_frame)
separator.grid(row=5, column=0, sticky='ew', padx=(20, 10), pady=20)

toggle_switch = ttk.Checkbutton(widget_frame, text='Toggle Light/Dark', style='Switch', command=toggle_LD)
toggle_switch.grid(row=6, column=0, sticky='nsew', padx=5, pady=10)

tree_frame = ttk.Frame(form_frame)
tree_frame.grid(row=0, column=1, pady=10)

tree_scroll = ttk.Scrollbar(tree_frame)
tree_scroll.pack(side='right', fill='y')

column = ["Name", "Age", "Nationality", "Staff"]
tree_view = ttk.Treeview(tree_frame, show='headings', columns=column, height=13)
tree_view.column("Name", width=100)
tree_view.column("Age", width=50)
tree_view.column("Nationality", width=100)
tree_view.column("Staff", width=50)
tree_view.pack()
tree_scroll.config(command=tree_view.yview)




root.mainloop()
