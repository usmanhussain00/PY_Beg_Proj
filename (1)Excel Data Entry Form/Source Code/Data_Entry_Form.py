import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl
#This is a simple data entry box
#------------------------------------------------------------------------------------

def enter_data():
#This is the def function we created where we get the specified Entry boxes data
#and print it on the terminal and after that we created a function using openpyxl
#where we entered the following data from the entry form and put it inside a excel file
    accepted =accept_var.get()

    if accepted=="Accepted":
        # User info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemester = numsemester_spinbox.get()

            print("First name:", firstname)
            print("Last name:", lastname)
            print("Title:", title)
            print("Age:", age)
            print("Nationality:", nationality)
            print("Numcourses:", numcourses)
            print("Numsemester:", numsemester)
            print("Registration:", registration_status)
            #-----------NOTE---------------#
            #specify the file path according to your pc
            filepath = "C:\\Users\\HP\\Desktop\\coding\\python\\Python project intermediate\\Excel Data Entry Form\\Database\\Data.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name","Last Name","Title","Age","Nationality"
                           ,"# Courses","# Semesters","Registration Status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname,lastname,title,age,nationality,numcourses
                          ,numsemester,registration_status])
            workbook.save(filepath)


        else:
            tkinter.messagebox.showwarning(title="ERROR", message="First name and Last name are required")

    else:
        tkinter.messagebox.showwarning(title="ERROR",message="You have not accepted the term")

#------------------------------------------------------------------------------------

window = tkinter.Tk()
window.title("Data Entry Form")

frame = tkinter.Frame(window)
frame.pack()

#------------------------------------------------------------------------------------

#Saving user info
user_info_frame = tkinter.LabelFrame(frame,text="user information")
user_info_frame.grid(row=0,column=0,padx=20,pady=20)

first_name_label = tkinter.Label(user_info_frame,text="First Name")
first_name_label.grid(row=0,column=0)

last_name_label = tkinter.Label(user_info_frame,text="Last Name")
last_name_label.grid(row=0,column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1,column=0)
last_name_entry.grid(row=1,column=1)

title_label = tkinter.Label(user_info_frame,text="Title")
title_combobox = ttk.Combobox(user_info_frame,values=["","Mr.","Mrs.","Dr."])
title_label.grid(row=0,column=2)
title_combobox.grid(row=1,column=2)

age_label = tkinter.Label(user_info_frame,text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame,from_=18, to=110)
age_label.grid(row=2,column=0)
age_spinbox.grid(row=3,column=0)

nationality_label = tkinter.Label(user_info_frame,text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame,values=["Pakistan","India" ,"Canada", "Italy", "Germany", "Japan", "Kazakhstan", "Russia", "South Korea", "United States" ])
nationality_label.grid(row=2,column=1)
nationality_combobox.grid(row=3,column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

#------------------------------------------------------------------------------------

#Saving course info
course_frame = tkinter.LabelFrame(frame)
course_frame.grid(row=1,column=0, sticky="news",padx=20,pady=20)
registered_label = tkinter.Label(course_frame,text="Registration")

reg_status_var = tkinter.StringVar(value="Registered")
registered_check = tkinter.Checkbutton(course_frame, text="Currently Registered ",
                                       variable=reg_status_var,onvalue="Registered",offvalue="Not Registered")

registered_label.grid(row=0,column=0)
registered_check.grid(row=1,column=0)

numcourses_label = tkinter.Label(course_frame,text="# Completed")
numcourses_spinbox = tkinter.Spinbox(course_frame, from_=0,to="infinity")
numcourses_label.grid(row=0,column=1)
numcourses_spinbox.grid(row=1,column=1)

numsemester_label = tkinter.Label(course_frame,text="# Semesters")
numsemester_spinbox = tkinter.Spinbox(course_frame, from_=0,to="infinity")
numsemester_label.grid(row=0,column=2)
numsemester_spinbox.grid(row=1,column=2)

for widget in course_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

#------------------------------------------------------------------------------------

#Accept terms
terms_frame = tkinter.LabelFrame(frame,text="Terms & Conditions")
terms_frame.grid(row=2,column=0, sticky="news",padx=20,pady=10)

accept_var = tkinter.StringVar(value="Not expected")
terms_frame = tkinter.Checkbutton(terms_frame,text="I accept the terms & conditions",
                                  variable=accept_var,onvalue="Accepted",offvalue="Not Accepted")
terms_frame.grid(row=0,column=0)

#------------------------------------------------------------------------------------

#Button
button = tkinter.Button(frame,text="Enter Data",command= enter_data)
button.grid(row=3,column=0,sticky="news",padx=20,pady=10)

#------------------------------------------------------------------------------------

window.mainloop()
