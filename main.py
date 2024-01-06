import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
import os
import openpyxl



canvas = ctk.CTk()
canvas.title("Entry Form")

ctk.set_appearance_mode("System")
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
title_combobox = ctk.CTkComboBox(user_input, values = ["","Mr. ", "Ms. ", "Dr. "])
title_combobox.grid(row=1, column = 2)
title.grid(row=0, column = 2)

slider = tk.IntVar()
slider.set(0)


age_label = ctk.CTkLabel(user_input, text = "Age")
age_spinbox = ctk.CTkSlider(user_input,from_=0, to=100, command=lambda s:slider.set(int(s)))
age_spinbox.grid(row=3,column=0)
age_label.grid(row=2,column=0)
age_value = ctk.CTkLabel(user_input,  textvariable=slider)
age_value.grid(row=3,column=1)

country = ["Afghanistan","Aland Islands","Albania","Algeria","American Samoa","Andorra","Angola","Anguilla","Antarctica","Antigua and Barbuda","Argentina","Armenia","Aruba","Australia","Austria","Azerbaijan","Bahamas","Bahrain","Bangladesh","Barbados","Belarus","Belgium","Belize","Benin","Bermuda","Bhutan","Bolivia,Plurinational State of","Bonaire,Sint Eustatius and Saba","Bosnia and Herzegovina","Botswana","Bouvet Island","Brazil","British Indian Ocean Territory","Brunei Darussalam","Bulgaria","Burkina Faso","Burundi","Cambodia","Cameroon","Canada","Cape Verde","Cayman Islands","Central African Republic","Chad","Chile","China","Christmas Island","Cocos (Keeling) Islands","Colombia","Comoros","Congo","Congo,The Democratic Republic of the","Cook Islands","Costa Rica","Côte d'Ivoire","Croatia","Cuba","Curaçao","Cyprus","Czech Republic","Denmark","Djibouti","Dominica","Dominican Republic","Ecuador","Egypt","El Salvador","Equatorial Guinea","Eritrea","Estonia","Ethiopia","Falkland Islands (Malvinas)","Faroe Islands","Fiji","Finland","France","French Guiana","French Polynesia","French Southern Territories","Gabon","Gambia","Georgia","Germany","Ghana","Gibraltar","Greece","Greenland","Grenada","Guadeloupe","Guam","Guatemala","Guernsey","Guinea","Guinea-Bissau","Guyana","Haiti","Heard Island and McDonald Islands","Holy See (Vatican City State)","Honduras","Hong Kong","Hungary","Iceland","India","Indonesia","Iran,Islamic Republic of","Iraq","Ireland","Isle of Man","Israel","Italy","Jamaica","Japan","Jersey","Jordan","Kazakhstan","Kenya","Kiribati","Korea,Democratic People's Republic of","Korea,Republic of","Kuwait","Kyrgyzstan","Lao People's Democratic Republic","Latvia","Lebanon","Lesotho","Liberia","Libya","Liechtenstein","Lithuania","Luxembourg","Macao","Macedonia,Republic of","Madagascar","Malawi","Malaysia","Maldives","Mali","Malta","Marshall Islands","Martinique","Mauritania","Mauritius","Mayotte","Mexico","Micronesia,Federated States of","Moldova,Republic of","Monaco","Mongolia","Montenegro","Montserrat","Morocco","Mozambique","Myanmar","Namibia","Nauru","Nepal","Netherlands","New Caledonia","New Zealand","Nicaragua","Niger","Nigeria","Niue","Norfolk Island","Northern Mariana Islands","Norway","Oman","Pakistan","Palau","Palestinian Territory,Occupied","Panama","Papua New Guinea","Paraguay","Peru","Philippines","Pitcairn","Poland","Portugal","Puerto Rico","Qatar","Réunion","Romania","Russian Federation","Rwanda","Saint Barthélemy","Saint Helena,Ascension and Tristan da Cunha","Saint Kitts and Nevis","Saint Lucia","Saint Martin (French part)","Saint Pierre and Miquelon","Saint Vincent and the Grenadines","Samoa","San Marino","Sao Tome and Principe","Saudi Arabia","Senegal","Serbia","Seychelles","Sierra Leone","Singapore","Sint Maarten (Dutch part)","Slovakia","Slovenia","Solomon Islands","Somalia","South Africa","South Georgia and the South Sandwich Islands","Spain","Sri Lanka","Sudan","Suriname","South Sudan","Svalbard and Jan Mayen","Swaziland","Sweden","Switzerland","Syrian Arab Republic","Taiwan,Province of China","Tajikistan","Tanzania,United Republic of","Thailand","Timor-Leste","Togo","Tokelau","Tonga","Trinidad and Tobago","Tunisia","Turkey","Turkmenistan","Turks and Caicos Islands","Tuvalu","Uganda","Ukraine","United Arab Emirates","United Kingdom","United States","United States Minor Outlying Islands","Uruguay","Uzbekistan","Vanuatu","Venezuela,Bolivarian Republic of","Viet Nam","Virgin Islands,British","Virgin Islands,U.S.","Wallis and Futuna","Yemen","Zambia","Zimbabwe"]


nationality_label = ctk.CTkLabel(user_input, text = "Nationality")
nationality_combobox = ctk.CTkComboBox(user_input,values = country)
nationality_label.grid(row=2,column=2)
nationality_combobox.grid(row=3,column=2)

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

