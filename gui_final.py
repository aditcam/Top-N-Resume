from PyPDF2 import PdfFileReader
import os
import numpy
from os import listdir
from os.path import isfile, join
import re
import nltk
from pyresparser import ResumeParser
import pandas as pd
import operator
import shutil
from datetime import datetime
import json
import matplotlib

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
import tkinter.font as tkFont
from tkinter import ttk
from tkinter.ttk import *
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg,NavigationToolbar2Tk
from matplotlib.figure import Figure
from math import pi
import matplotlib.pyplot as plt
import time
from PIL import Image, ImageTk
from itertools import count
import threading

LARGE_FONT = ("Verdana", 12)

#The paths for the Resume and CSV files

resume_path = 'None' # This is where the resumes are stored post categorization  and it Should be decided by the front end
profile_path = 'csv'
savefile_path = 'save_path'
resume_final_path = None

profile_data = {}  # a dictionary to store all the dataframes

job_profiles_list = None
job_skills_list = None
employee_data = []
user_requirements = []
num_resumes = None
yrs_exp = None
desired_job = None
top_n = None

time_taken = None
job_profile_count = None

state = 0 # unable to solve error, for some reason the functions in pageTwo are executing before the previous pages are being executed

class resumeParser(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        resumeParser.container = tk.Frame(self, width=800, height=1000)
        resumeParser.container.pack(side="top", fill="both", expand=True)
        self.geometry("800x600+220+30")

        self.title("Resume Parser")

        resumeParser.container.grid_rowconfigure(0, weight=1)
        resumeParser.container.grid_columnconfigure(0, weight=1)

        resumeParser.frames = {}

        for F in (StartPage, PageOne, SplashScreen):
            frame = F(resumeParser.container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(SplashScreen)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)


        label_header = tk.Label(self, text="Resume Analyser", font=('Verdana', 15)).pack(side=TOP, pady=15)
        label_operation = tk.Label(self, text="Select the filepath which contains all your resume pdf and docx files ", font=('Verdana', 10)).pack(side=TOP, pady=15)

        folder_selected = None
        conf_button = Button(self)
        filePath_label = tk.Label(self)

        def selectfold():
            folder_selected = filedialog.askdirectory()
            # print(folder_selected)
            filePath_label.config(text="\n    You have selected: " + folder_selected +'    \n',font=('Verdana', 9), bg = "#BBBBBB")
            filePath_label.pack(fill=Y, pady=30)
            #conf_button = Button(self, text='Confirm')  # INSERT Ccommand = new-Window
            conf_button.configure(text='Confirm',command=lambda: controller.show_frame(PageOne))
            conf_button.pack()
            res_path_final = folder_selected.replace('/', '\\')
            print(res_path_final)
            messagebox.showinfo("Path Selection ", "You have selected " + res_path_final)
            global resume_path
            resume_path = res_path_final

            #creating all the output and input directories required for the program to function

            # get a list consisting of the paths of all the csv files
            global profile_path
            profile_files = [os.path.join(profile_path, f) for f in os.listdir(profile_path) if
                             os.path.isfile(os.path.join(profile_path, f))]
            # print(profile_files)


            for file in profile_files:
                keyname = os.path.basename(file)[:-len('.csv')]  # this is just to obtain every file name sans the extension
                x = pd.read_csv(file)
                global profile_data
                profile_data[keyname] = x.apply(lambda x: x.astype(str).str.lower())

            global job_profiles_list
            job_profiles_list = list(profile_data.keys())  # Storing the job profile names in list format

            global job_profile_count
            job_profile_count = {}
            for profile in job_profiles_list:
                job_profile_count[profile] = 0

            global resume_final_path
            resume_final_path = os.path.join(resume_path, 'final_categorized_resumes')
            if not os.path.exists(resume_final_path):
                os.mkdir(resume_final_path)
            for job_profile in job_profiles_list:
                job_profile_final_path = os.path.join(resume_final_path, job_profile)
                if not os.path.exists(job_profile_final_path):
                    os.mkdir(job_profile_final_path)

        photo = PhotoImage(file="folder_img.png")
        photoimage = photo.subsample(15,15)
        Button1 = Button(self, text ='   Select folder   ', compound=LEFT, command=selectfold, image = photoimage)
        Button1.image = photoimage
        Button1.pack(side=TOP, pady=20)

class PageOne(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label_1 = tk.Label(self, text="Enter Job Profile Requirements", font=LARGE_FONT)
        label_1.pack(pady=10, padx=10)

        label_2 = tk.Label(self,text="Select the Required Domain").pack(pady=5)
        profile_chooser = ttk.Combobox(self)
        profile_chooser['values'] = (' Business Analyst',
                              ' Cloud Architect',
                              ' Data Scientist',
                              ' Full Stack Dev',
                              ' Testing'
        )
        profile_chooser['state'] = 'readonly'
        profile_chooser.pack()
        profile_chooser.current(0)

        label_3 = tk.Label(self,text=" Enter the skills required ").pack(pady=10)
        label_4 = tk.Label(self,text=" (Eg. Java, Python, SQL, C++) ").pack(pady=2)
        entry_1 = Entry(self,width=50)
        entry_1.pack(pady=5)

        label_5 = tk.Label(self,text=" Enter the experience required (in years)").pack(pady=5)
        entry_2 = Entry(self,width=50)
        entry_2.pack(pady=5)

        label_5 = tk.Label(self,text=" Enter number of top resumes required ").pack(pady=5)
        entry_3 = Entry(self,width=50)
        entry_3.pack(pady=5)

        label_6 = tk.Label(self,text='Please wait while the files are being categorized ')
        label_7 = tk.Label(self,text=' 0% completed ')
        progress = Progressbar(self, orient=HORIZONTAL, length=300, mode='determinate')

        def load_task():

            label_6.pack(pady=30)
            progress.pack(pady=15)
            label_7.pack(pady=5)

            global resume_path
            file_count = len([name for name in os.listdir(resume_path) if os.path.isfile(os.path.join(resume_path, name))])
            percentage_completion = 0
            for f in os.listdir(resume_path):
                if os.path.isfile(os.path.join(resume_path, f)):

                    data = ResumeParser(os.path.join(resume_path, f)).get_extracted_data()
                    # print(data)
                    # print(" ")
                    percentage_completion += 100/file_count
                    progress['value'] = percentage_completion
                    label_7['text'] = '{}% completed'.format(round(percentage_completion))
                    self.update_idletasks()

                    proficiency = []

                    # Iterating through all the job dataframes in profile_data and
                    # For each job, iterate through all skills
                    # If skills match skills from resume increment specific_skills by one

                    skills_lower = [i.lower() for i in data['skills']]  # convert all the skills to lowercase
                    global profile_data
                    global job_skills_list
                    job_skills_list = [{} for i in job_profiles_list]  # a dictionary of skills for representing the skills of employee in each job
                    for job_no, tup in enumerate(profile_data.items()):

                        jobs, jobs_df = tup

                        talent_in_job = 0  # a metric to measure skill in a job profile based on number of mentions in resume

                        for skills_col in jobs_df.columns:

                            specific_skill_level = len(list(set(skills_lower).intersection(set(jobs_df[skills_col]))))

                            job_skills_list[job_no][skills_col] = specific_skill_level

                            talent_in_job += specific_skill_level

                        proficiency.append(talent_in_job)
                        # proficiency_measure[jobs] = talent_in_job

                    # get the name of the job that has most matches and move the file to the job's folder

                    most_proficient = proficiency.index(max(proficiency))

                    data['skills'] = skills_lower[:]
                    data['proficiency'] = job_profiles_list[most_proficient]
                    data['job_skills_dict'] = job_skills_list[most_proficient].copy()
                    global employee_data
                    employee_data.append(data.copy())
                    global job_profile_count
                    job_profile_count[data['proficiency']] +=1

                    new_directory = os.path.join(resume_final_path, job_profiles_list[most_proficient],f)
                    old_directory = os.path.join(resume_path, f)
                    shutil.move(old_directory, new_directory)

                    print(data)
                    print(" ")

            print("Loop is over")
            global user_requirements
            global yrs_exp


            #th = threading.Thread(target=threadFunc)
            #th.start()
            #bar()
            #th.join()
            #print("SSop bro ?? ")

        def ranker():
            global employee_data
            global job_profiles_list
            global desired_job
            job = job_profiles_list[desired_job]
            print(job)

            x = [i for i in employee_data if i['proficiency'] == job]
            # print(x)

            for job_dicts in x:
                # print(job_dicts['skills'])
                exp = job_dicts['total_experience']
                exp_score = 50 * (exp / yrs_exp) if yrs_exp>0 and exp<=yrs_exp else 50 if yrs_exp>0 else 0
                # print(set(user_requirements).intersection(set(job_dicts['skills'])))
                req_met = len(list(set(user_requirements).intersection(set(job_dicts['skills']))))
                req_score = 50 * (req_met / len(user_requirements))
                job_dicts['score'] = exp_score + req_score

            x = sorted(x, key=lambda i: i['score'], reverse=True)

            return (x)

        def check_fields():

            if len(entry_1.get()) == 0 or len(entry_2.get()) == 0 or len(entry_3.get()) == 0 or len(profile_chooser.get()) == 0:
                messagebox.showerror(message="Please Fill All The Fields", title="error")
            else:
                if not (entry_2.get().isdigit() and entry_3.get().isdigit()):
                    messagebox.showerror(message="Enter a number ", title="error")
                    if not (entry_2.get().isdigit()):
                        entry_2.delete(0, END)
                    if not (entry_3.get().isdigit()):
                        entry_3.delete(0, END)
                else:
                    skillList = entry_1.get()
                    sl = skillList.replace(" ", "")
                    global user_requirements
                    global yrs_exp
                    global desired_job
                    global num_resumes
                    user_requirements = sl.split(',')
                    num_resumes = int(entry_3.get())
                    print("**{}**".format(num_resumes))
                    yrs_exp = int(entry_2.get())
                    desired_job = profile_chooser.current()
                    global time_taken
                    start = time.time()
                    load_task()
                    end = time.time()
                    time_taken = end - start
                    print("Time take is {}".format(time_taken))
                    global top_n
                    top_n = ranker()
                    print(top_n)
                    resumeParser.frames[PageTwo] = PageTwo(resumeParser.container, self)
                    resumeParser.frames[PageTwo].grid(row=0, column=0, sticky="nsew")
                    print(resumeParser.frames)
                    controller.show_frame(PageTwo)

                    # return and store  the list of required skills in sl (skill list)

        button_1 = tk.Button(self, text='Next', command=check_fields)
        button_1.pack(pady=5)
        Button_2 = tk.Button(self, text="Back to Homepage", command=lambda: controller.show_frame(StartPage))
        Button_2.pack(pady=10)


class PageTwo(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        def btn_click():
            popup = Toplevel()
            popup.geometry("1100x800+100+10")
            popup.grab_set()

            global job_profile_count

            labels = [l.replace("_"," ").capitalize() for l in job_profile_count.keys() if job_profile_count[l]>0]
            values = [v for v in job_profile_count.values() if v > 0]

            # now to get the total number of failed in each section
            actualFigure = plt.figure(figsize=(1,1),dpi = 80)
            actualFigure.suptitle("Profile Stats of Analyzed resumes ", fontsize=10)

            # as explode needs to contain numerical values for each "slice" of the pie chart (i.e. every group needs to have an associated explode value)
            explode = list()
            for k in labels:
                explode.append(0.1)

            pie = plt.pie(values,explode=explode, shadow=True, autopct='%1.1f%%')
            plt.legend(pie[0], labels, loc="upper right")

            canvas = FigureCanvasTkAgg(actualFigure, popup)
            canvas.get_tk_widget().pack()
            canvas.draw()

            toolbar = NavigationToolbar2Tk(canvas, popup)
            toolbar.update()
            canvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            global time_taken

            label_1 = tk.Label(popup,text="Total number of Resumes parsed : {}".format(sum(values)))
            label_2 = tk.Label(popup,text = "Total time taken for parsing : {} sec".format(round(time_taken,2)))
            label_3 = tk.Label(popup,text = "Average time taken for parsing a resume : {} sec ".format(round(time_taken/sum(values),2)))
            label_4 = tk.Label(popup,text = "A Distribution of Parsed resumes ")
            label_1.pack(pady = 2)
            label_2.pack(pady =2)
            label_3.pack(pady =2)
            label_4.pack(pady=2)

            for l,v in zip(labels,values):
                label_5 = tk.Label(popup,text="{} : {}".format(l,v))
                label_5.pack(pady=1)

        def onSelectListItem(evt):
            w = evt.widget
            index =int(w.curselection()[0])
            print(index)
            popup = Toplevel()
            popup.geometry("1100x800+100+10")
            popup.grab_set()
            global top_n
            label_1 = Label(popup,text = "{}".format(top_n[index-1]['name']).capitalize(),font = "Helvetica 14 bold italic")
            label_2 = Label(popup, text="Mob: {} | Email: {}".format(top_n[index - 1]['mobile_number'],top_n[index - 1]['email']))
            label_3 = Label(popup, text="Skills Extracted", font = "Helvetica 12 bold ")
            label_4 = Label(popup, text = "{}".format(top_n[index-1]['skills']).replace("'"," ").replace("[","").replace("]",""),wraplength = 500)
            label_5 = Label(popup, text="Work Experience (In years) : {}".format(top_n[index-1]['total_experience']))
            label_1.pack(pady=4)
            label_2.pack(pady=1)
            label_3.pack(pady=5)
            label_4.pack(pady=5)
            label_5.pack(pady=5)

            profile = top_n[index-1]

            skill_pts = list(profile['job_skills_dict'].values())  # get the skill values

            skill_names = list(profile['job_skills_dict'].keys())  # get the skill names


            f = Figure(figsize=(2, 2), dpi=75)
            a = f.add_subplot(111)

            data = numpy.array(skill_pts)

          # the x locations for the groups
            width = .5

            rects1 = a.bar(skill_names, data, width)
            a.set_title("Skill graph")
            a.set_xlabel('Technologies -->')
            a.set_ylabel('Number of Technologies Known -->')

            canvas = FigureCanvasTkAgg(f, popup)
            canvas.draw()
            canvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

            toolbar = NavigationToolbar2Tk(canvas, popup)
            toolbar.update()
            canvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        global num_resumes
        global state
        global top_n
        print('Hello')
        listbox = Listbox(self,width=200)
        label_1 = tk.Label(self, text="The top {} resume are ".format(num_resumes))
        label_1.pack(pady=10)
        counter = 2
        listbox.insert(1,"{} {} {} {}".format('Rank'.ljust(20),'Name'.ljust(50),'Email Id'.ljust(80),'Contact No.'.ljust(50)))
        for profile in top_n[:min(num_resumes,len(top_n))]:
            listbox.insert(counter,"{} {} {} {}".format(str(counter-1).ljust(20),profile['name'].ljust(50),profile['email'].ljust(60),profile['mobile_number'].ljust(50)))
            counter+=1
        listbox.pack(pady=10)

        listbox.bind('<<ListboxSelect>>', onSelectListItem)

        Button_1 = tk.Button(self,text="Additional stats",command = btn_click)
        Button_1.pack()


class SplashScreen(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.config(bg='#FFFFFF')
        photo = PhotoImage(file=r"C:\Users\AdityaShekharCamarus\Desktop\ResumeParserLogo.png")
        photoimage = photo.subsample(2,2)
        Logo_ResumeAnalyser = tk.Label(self, compound=LEFT, image = photoimage)
        Logo_ResumeAnalyser.image = photoimage
        Logo_ResumeAnalyser.pack(pady=50)
        fontStyle = tkFont.Font(family="Lucida Grande", size=8)
        Label_designers = tk.Label(self, text = 'Designed by Aditya, Pranav x 2, Lalit',bg='white',fg = '#000000', font = fontStyle).pack(side = BOTTOM, pady = 10)
        fontStyle = tkFont.Font(family="Lucida Grande", size=12)
        Button_instruction = tk.Button(self, text = 'Click here to start the application ', borderwidth = 0,bg='white',fg = '#0591AB', font = fontStyle, command = lambda: controller.show_frame(StartPage)).pack(pady = 15)

app = resumeParser()
app.mainloop()