import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
import os
import openpyxl



canvas = ctk.CTk()
canvas.title("Entry Form")


ctk.set_default_color_theme("green")  # Themes: "blue" (standard), "green", "dark-blue"

frame = ctk.CTkFrame(canvas)
frame.pack()

user_input = ctk.CTkLabel(frame, text = "User Input")
user_input.grid(row=0, column=0, padx=80, pady=20)

first_name_label = ctk.CTkLabel(user_input, text = "First Name")
first_name_label.grid(row=0, column=0)
first_name_input = tk.Entry(user_input)
first_name_input.grid(row=1, column=0)

last_name_label = ctk.CTkLabel(user_input, text = "Last Name")
last_name_label.grid(row=0, column=1)
last_name_input = tk.Entry(user_input)
last_name_input.grid(row=1, column=1)

title = ctk.CTkLabel(user_input, text = "Title")
title_combobox = ttk.Combobox(user_input, values = ["","Mr. ", "Ms. ", "Dr. "])
title_combobox.grid(row=1, column = 2)
title.grid(row=0, column = 2)

age_label = ctk.CTkLabel(user_input, text = "Age")
age_spinbox = ttk.Spinbox(user_input, from_=0, to=100)
age_spinbox.grid(row=3,column=0)
age_label.grid(row=2,column=0)

country = open("country", "r").read()

nationality_label = ctk.CTkLabel(user_input, text = "Nationality")
nationality_combobox = ttk.Combobox(user_input,values = country)
nationality_label.grid(row=2,column=1)
nationality_combobox.grid(row=3,column=1)

for widget in user_input.winfo_children():
    widget.grid_configure(padx=15, pady=10)

def entry():

    fn = first_name_input.get()
    ln = last_name_input.get()
    title = title_combobox.get()
    age = age_spinbox.get()
    nationality = nationality_combobox.get()

    info = {"fn":fn,
             "ln":ln,
             "title":title,
             "age":age,
             "nationality":nationality}

    filepath = "data.xlsx"

    if not os.path.exists(filepath):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        heading = ["First Name", "Last Name", "Title", "Age", "Nationality"]
        sheet.append(heading)
        workbook.save(filepath)

    workbook = openpyxl.load_workbook(filepath)
    worksheet = workbook.active
    worksheet.append([fn, ln, title, age, nationality])
    workbook.save(filepath)
    print(info)

button = ctk.CTkButton(frame, text = "Enter", command = entry)
button.grid(row=4,column=0, sticky="news", padx=20, pady=20)



canvas.mainloop()

