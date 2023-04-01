import tkinter as tk
from tkinter import *
import tkinter.messagebox
import docx
import tkinter as tk
from tkinter import ttk
import openpyxl
from tkcalendar import DateEntry
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import random
from functools import partial
import shutil
import os
import pyautogui
from PIL import ImageTk, Image
import PIL.Image


# defining fonts for usage in project
NORM_FONT = ('times new roman', 13, 'normal')
NORM_FONT_MEDIUM_HIGH = ('times new roman', 15, 'normal')
NORM_FONT_MEDIUM_LOW = ('times new roman', 14, 'normal')
TIMES_NEW_ROMAN_BIG = ('times new roman', 16, 'normal')
NORM_VERDANA_FONT = ('verdana', 10, 'normal')
BOLD_VERDANA_FONT = ('verdana', 11, 'normal')
LARGE_VERDANA_FONT = ('verdana', 13, 'normal')
XXL_FONT = ('times new roman', 25, 'normal')
XL_FONT = ('times new roman', 20, 'normal')
L_FONT = ('times new roman', 15, 'normal')

def reset_MCQCreatorWindow(assessmentName_text,total_questions_entry,low_complexity_entry,medium_complexity_entry,high_complexity_entry,duration_entry,totalMarks_entry):
    assessmentName_text.delete(0, tk.END)
    total_questions_entry.delete(0, tk.END)
    low_complexity_entry.delete(0, tk.END)
    medium_complexity_entry.delete(0, tk.END)
    high_complexity_entry.delete(0, tk.END)
    duration_entry.delete(0, tk.END)
    totalMarks_entry.delete(0, tk.END)
    assessmentName_text.insert(0, "")
    total_questions_entry.insert(0, "")
    low_complexity_entry.insert(0, "")
    medium_complexity_entry.insert(0, "")
    high_complexity_entry.insert(0, "")
    duration_entry.insert(0, "")
    totalMarks_entry.insert(0, "")

def close_app():
    root.destroy()

def write_to_word_file(total_question_count, low_complexity_percentage, medium_complexity_percentage,
                       high_complexity_percentage):
    # Calculate the actual count of questions for each complexity
    print("Total Question Count : ",total_question_count.get())
    total_questions = float(total_question_count.get())
    low_complexity_count = round((float(low_complexity_percentage) / 100) * total_questions)
    medium_complexity_count = round((float(medium_complexity_percentage) / 100) * total_questions)
    high_complexity_count = round((float(high_complexity_percentage) / 100) * total_questions)

    print("Low complexity count: ", low_complexity_count," High Complexity count :", high_complexity_count," Medium complexity count : ", medium_complexity_count)
    # create a new word file
    doc = docx.Document()

    # add the contents to the file
    name_para = doc.add_paragraph("Name:\t\t\t\t\t\tMCQ1\t\t\t\t\t Date:")
    name_para.style.font.name = "Yu Gothic"
    name_para.style.font.size = docx.shared.Pt(10)
    name_para.style.font.bold = True

    emp_para = doc.add_paragraph("Emp#\t\t\t\t\t\tC++")
    emp_para.style.font.name = "Yu Gothic"
    emp_para.style.font.size = docx.shared.Pt(10)
    emp_para.style.font.bold = True

    time_para = doc.add_paragraph("Time Duration: 120 Minutes\t\t\t\t\t\tTotal Marks: 50")
    time_para.style.font.name = "Yu Gothic"
    time_para.style.font.size = docx.shared.Pt(10)
    time_para.style.font.bold = True

    note_para = doc.add_paragraph(
        "------------------------------------------------------------------------------------------------------------")
    note_para.style.font.name = "Yu Gothic"
    note_para.style.font.size = docx.shared.Pt(10)
    note_para.style.font.bold = True

    note_para = doc.add_paragraph("Note: Marks for every question mentioned along with the question it")
    note_para.style.font.name = "Yu Gothic"
    note_para.style.font.size = docx.shared.Pt(10)
    note_para.style.font.bold = True

    # Read data from excel file
    workbook = openpyxl.load_workbook("Master.xlsx")
    sheet = workbook.active
    complexity_count = [low_complexity_count, medium_complexity_count, high_complexity_count]
    complexity = ["low", "medium", "high"]
    question_fetched = 0
    low_complexity_fetched = 0
    medium_complexity_fetched = 0
    high_complexity_fetched = 0

    while question_fetched < total_questions:
        print("Question fetched : ",question_fetched)
        random_index = random.randint(0, 2)
        if complexity_count[random_index] > 0:
            for row in sheet.iter_rows(values_only=True):
                if row[2].lower() == complexity[random_index]:
                    bIsComplexityCountStillValid = False;
                    if ((row[2].lower() == 'low'and low_complexity_fetched < low_complexity_count) or
                        (row[2].lower() == 'medium' and medium_complexity_fetched < medium_complexity_count) or
                        (row[2].lower() == 'high' and high_complexity_fetched < high_complexity_count)):
                        bIsComplexityCountStillValid = True
                        if(True == bIsComplexityCountStillValid):
                            serial_no, question_description, _, topic, mark,answer = row
                            para = doc.add_paragraph("Question Description: {}\t\tMark: {}".format(question_description, mark))
                            para.style.font.name = "Yu Gothic"
                            para.style.font.size = docx.shared.Pt(10)
                            para.style.font.bold = False
                            complexity_count[random_index] -= 1
                            if row[2].lower() == 'low':
                                low_complexity_fetched +=1
                            elif row[2].lower() == 'medium':
                                medium_complexity_fetched +=1
                            else:
                                high_complexity_fetched +=1

                            question_fetched += 1
                            if question_fetched == total_question_count:
                                break

    # save the file
    doc.save("MCQ.docx")
    # get the path of the user's desktop folder
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

    # path of the destination file on the desktop
    dest_path = os.path.join(desktop_path, 'MCQ.docx')

    # copy the file from source to destination
    shutil.copy('MCQ.docx', dest_path)

def generate_check(total_questions_entry,low_complexity_entry,medium_complexity_entry,high_complexity_entry):
    low = float(low_complexity_entry.get())
    medium = float(medium_complexity_entry.get())
    high = float(high_complexity_entry.get())
    total = low + medium + high
    if total == 100:
        write_to_word_file(total_questions_entry, low, medium,
                           high)
        tk.messagebox.showinfo("Success", "MCQ is ready")
    else:
        tk.messagebox.showerror("Error", "The sum of the low, medium, and high complexity percentages is not equal to 100.")
def donothing(event=None):
    print("Button is disabled")
    pass
def create_MCQWindow(master):
    create_MCQWindow = Toplevel(master)
    create_MCQWindow.title("Assessment Paper Generator ")
    create_MCQWindow.geometry('455x440+700+250')
    create_MCQWindow.configure(background='wheat')
    create_MCQWindow.resizable(width=False, height=False)
    create_MCQWindow.protocol('WM_DELETE_WINDOW', donothing)
    heading = Label(create_MCQWindow, text="New Assessment Creation",
                    font=('ariel narrow', 15, 'bold'),
                    bg='wheat')
    dataEntryFrame = Frame(create_MCQWindow, width=200, height=130, bd=4, relief='ridge',
                           bg='snow')
    default_text1 = StringVar(dataEntryFrame, value='')
    default_text2 = StringVar(dataEntryFrame, value='')
    default_text3 = StringVar(dataEntryFrame, value='')
    default_text4 = StringVar(dataEntryFrame, value='')

    # lower frame added to show the result of transactions
    infoFrame = Frame(create_MCQWindow, width=70, height=20, bd=4, relief='ridge')
    # create a Book Name label
    name_label = Label(dataEntryFrame, text="Assessment Name", width=15, anchor=W, justify=LEFT,
                        font=NORM_FONT,
                        bg='snow')

    assessment_date_label = Label(dataEntryFrame, text="Assessment Date", width=15, anchor=W,
                         justify=LEFT,
                         font=NORM_FONT, bg='snow')

    totQuest_label = Label(dataEntryFrame, text="Total Questions", width=15, anchor=W,
                       justify=LEFT,
                       font=NORM_FONT, bg='snow')

    low_label = Label(dataEntryFrame, text="Low %", width=15, anchor=W, justify=LEFT,
                         font=NORM_FONT,
                         bg='snow')

    medium_label = Label(dataEntryFrame, text="Medium %", width=15, anchor=W,
                         justify=LEFT,
                         font=NORM_FONT, bg='snow')
    high_label = Label(dataEntryFrame, text="High %", width=15, anchor=W,
                         justify=LEFT,
                         font=NORM_FONT, bg='snow')
    totolMarks_label = Label(dataEntryFrame, text="Total Marks", width=15, anchor=W,
                         justify=LEFT,
                         font=NORM_FONT, bg='snow')
    duration_label = Label(dataEntryFrame, text="Duration", width=15, anchor=W,
                         justify=LEFT,
                         font=NORM_FONT, bg='snow')

    infolabel = Label(infoFrame, text="All fields are mandatory !!", width=40, anchor='center',
                      justify=LEFT,
                      font=NORM_VERDANA_FONT, bg='snow', fg="black")

    heading.grid(row=0, column=0, columnspan=2)
    dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
    name_label.grid(row=0, column=0, pady=5)
    assessment_date_label.grid(row=1, column=0, pady=5)
    totQuest_label.grid(row=2, column=0, pady=5)
    low_label.grid(row=3, column=0, pady=5)
    medium_label.grid(row=4, column=0, pady=5)
    high_label.grid(row=5, column=0, pady=5)
    totolMarks_label.grid(row=6, column=0, pady=5)
    duration_label.grid(row=7, column=0, pady=5)

    # create a text entry box
    # for typing the information
    assessmentName_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow')
    cal = DateEntry(dataEntryFrame, width=28, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                    anchor=W, justify=LEFT)
    totQuest_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow')
    lowPerc_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow')
    mediumPerc_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow')
    highPerc_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow')
    totMarks_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow')
    duration_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow')

    assessmentName_text.grid(row=0, column=1, pady=5)
    cal.grid(row=1, column=1, pady=5)
    totQuest_text.grid(row=2, column=1, pady=5)
    lowPerc_text.grid(row=3, column=1, pady=5)
    mediumPerc_text.grid(row=4, column=1, pady=5)
    highPerc_text.grid(row=5, column=1, pady=5)
    totMarks_text.grid(row=6, column=1, pady=5)
    duration_text.grid(row=7, column=1, pady=5)

    # ---------------------------------Button Frame Start----------------------------------------
    buttonFrame = Frame(create_MCQWindow, width=200, height=100, bd=4, relief='ridge')
    buttonFrame.grid(row=20, column=1, pady=8)
    generate_deposit = Button(buttonFrame)

    insert_result = partial(generate_check,totQuest_text,lowPerc_text, mediumPerc_text,highPerc_text)

    # create a Save Button and place into the create_MCQWindow window
    generate_deposit.configure(text="Generate", fg="Black", command=insert_result,
                             font=NORM_FONT, width=8, bg='light cyan', underline=0, state=NORMAL)
    generate_deposit.grid(row=0, column=0)

    clear_result = partial(reset_MCQCreatorWindow,assessmentName_text,totQuest_text,lowPerc_text, mediumPerc_text,highPerc_text,duration_text,totMarks_text)
    reset_button = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                   font=NORM_FONT, width=8, bg='light cyan', underline=0)
    reset_button.grid(row=0, column=1)

    # create a Cancel Button and place into the create_MCQWindow window
    cancel = Button(buttonFrame, text="Close", fg="Black", command=master.destroy,
                    font=NORM_FONT, width=8, bg='light cyan', underline=0)
    cancel.grid(row=0, column=2)
    # ---------------------------------Button Frame End----------------------------------------

    infoFrame.grid(row=21, column=1, pady=5)
    infolabel.grid(row=0, column=0, padx=2, pady=3)

    create_MCQWindow.bind('<Return>', lambda event=None: submit_deposit.invoke())
    create_MCQWindow.bind('<Alt-d>', lambda event=None: submit_deposit.invoke())
    create_MCQWindow.bind('<Alt-c>', lambda event=None: cancel.invoke())
    create_MCQWindow.bind('<Alt-r>', lambda event=None: clear.invoke())

    create_MCQWindow.focus()
    create_MCQWindow.grab_set()
    #mainloop()
    
def insert_questions(master):
    insertQ_window = Toplevel(master)

    headingForm = "Add Assessment Questions"
    insertQ_window.title("Question Bank Creation ")

    insertQ_window.geometry('760x615+700+250')
    insertQ_window.configure(background='wheat')
    insertQ_window.resizable(width=True, height=True)

    heading = Label(insertQ_window, text=headingForm, font=('ariel narrow', 15, 'bold'),
                    bg='wheat')
    heading.grid(row=0, column=0, columnspan=3)
    right_frame = Frame(insertQ_window, width=600, bd=6, relief='ridge', bg='light yellow')
    left_frame = Frame(insertQ_window, width=600, bd=6, relief='ridge', bg='light yellow')

    question_label =  Label(right_frame, text="Question Description :", width=20, anchor=W, justify=LEFT,
                          font=NORM_FONT,
                          bg='light yellow')
    question_label.grid(row=0, column=0, padx=10, pady=10, sticky="W")
    question_entry = tk.Text(right_frame, height=15, width=80, font=NORM_FONT)
    question_entry.grid(row=1, column=0, padx=10, pady=10, sticky="W")

    right_frame.grid(row=1, column=1, padx=2, pady=5, sticky=W)
    left_frame.grid(row=2, column=1, padx=2, pady=5, sticky=W)

    infoFrame = Frame(insertQ_window, width=200, height=100, bd=8, relief='ridge', bg='light yellow')
    infoFrame.grid(row=4, column=0, padx=90, pady=5, columnspan=4, sticky=W)

    # ---------------------------------Preparing display Area - start ---------------------------------

    itemnametext = StringVar(left_frame)
    itemnamelabel = Label(left_frame, text="Complexity", width=12, anchor=W, justify=LEFT,
                          font=NORM_FONT,
                          bg='light yellow')
    itemnamelabel.grid(row=0, column=1, padx=10, pady=5)

    complexities = ["High", "Medium", "Low"]
    complexity_dropdown = tk.ttk.Combobox(left_frame, values=complexities, state="readonly",width=23, justify=LEFT, font=NORM_FONT)
    complexity_dropdown.current(0)
    complexity_dropdown.grid(row=0, column=2, pady=5)

    # Display item Id - Row 4
    descriptiontext = StringVar(left_frame)
    descriptionlabel = Label(left_frame, text="Subject", width=12, anchor=W, justify=LEFT,
                             font=NORM_FONT,
                             bg='light yellow')
    descriptionlabel.grid(row=0, column=3, padx=10, pady=5)
    description_Text = Entry(left_frame, text="", width=25, justify=LEFT, textvariable=descriptiontext,
                             font=NORM_FONT,
                             bg='snow')
    description_Text.grid(row=0, column=4, padx=5, pady=5)

    # Display Father name - Row 5

    # Display Country Name - Row 5
    quantitytext = StringVar(left_frame)
    quantityLabel = Label(left_frame, text="Topic", width=12, anchor=W, justify=LEFT,
                          font=NORM_FONT,
                          bg='light yellow')
    quantityLabel.grid(row=1, column=1, padx=10, pady=5)
    quantity_Text = Entry(left_frame, text="", width=25, justify=LEFT, textvariable=quantitytext,
                          font=NORM_FONT,
                          bg='snow')
    quantity_Text.grid(row=1, column=2, pady=5)

    unitpricetext = StringVar(left_frame)
    unitpriceLabel = Label(left_frame, text="Marks", width=12, anchor=W, justify=LEFT,
                           font=NORM_FONT,
                           bg='light yellow')
    unitpriceLabel.grid(row=1, column=3, padx=10, pady=5)
    unitprice_Text = Entry(left_frame, text="", textvariable=unitpricetext, width=25, justify=LEFT,
                           font=NORM_FONT,
                           bg='snow')
    unitprice_Text.grid(row=1, column=4, padx=5, pady=5)

    racktext = StringVar(left_frame)

    infoLabel = Label(infoFrame, text="Press Save button to save the modified records", width=60,
                      anchor='center',
                      justify=CENTER,
                      font=NORM_FONT,
                      bg='light yellow')



    infoLabel.grid(row=1, column=0, padx=10, pady=5)

    # ---------------------------------Button Frame Start----------------------------------------
    buttonFrame = Frame(insertQ_window, width=200, height=100, bd=4, relief='ridge')
    buttonFrame.grid(row=3, column=1, pady=8)
    submit_deposit = Button(buttonFrame)

    # insert_result = partial(registerlocalCenter, trust_nametext, pledge_text, infolabel)

    # create a Save Button and place into the new_center_window window
    submit_deposit.configure(text="Save", fg="Black", command=NONE,
                             font=NORM_FONT, width=8, bg='light cyan', underline=0, state=NORMAL)
    submit_deposit.grid(row=0, column=0)
    """
    clear_result = partial(clearRegisterPledgeForm,
                           pledge_text,
                           trust_nametext, infolabel)
    """
    clear = Button(buttonFrame, text="Reset", fg="Black", command=NONE,
                   font=NORM_FONT, width=8, bg='light cyan', underline=0)
    clear.grid(row=0, column=1)

    # create a Cancel Button and place into the new_center_window window
    # cancel_Result = partial(destroyWindow, new_center_window)
    cancel = Button(buttonFrame, text="Close", fg="Black", command=NONE,
                    font=NORM_FONT, width=8, bg='light cyan', underline=0)
    cancel.grid(row=0, column=2)
    # ---------------------------------Button Frame End----------------------------------------

    insertQ_window.bind('<Return>', lambda event=None: submit.invoke())
    insertQ_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
    insertQ_window.bind('<Alt-r>', lambda event=None: self.print_button.invoke())

    insertQ_window.focus()
    insertQ_window.grab_set()
    #mainloop()

def designMainScreen(master):
    labelFrame = Label(master, text="Assessment Creator", justify=CENTER,
                       font=XXL_FONT,
                       fg='black')
    # labelFrame.place(x=200, y=10)
    result_btnAddQuestion = partial(insert_questions,master)
    btn_addQues = Button(master, text="Add Question", fg="Black", command=result_btnAddQuestion,
                           font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')
    # labelFrame.place(x=200, y=10)
    result_createPaper = partial(create_MCQWindow,master)
    btn_createPaper = Button(master, text="Create Paper", fg="Black", command=result_createPaper,
                         font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')

    btn_usrCtrl = Button(master, text="User Control", fg="Black", command=None,
                             font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')
    btn_exit = Button(master, text="Exit", fg="Black", command=master.destroy,
                      font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')

    btn_addQues.place(x=65, y=220)
    btn_createPaper.place(x=65, y=275)
    btn_usrCtrl.place(x=65, y=330)
    btn_exit.place(x=65, y=385)

    master.bind('<Escape>', lambda event=None: btn_exit.invoke())

    master.bind('<I>', lambda event=None: btn_inventory.invoke())
    master.bind('<i>', lambda event=None: btn_inventory.invoke())
    master.bind('<S>', lambda event=None: btn_sales.invoke())
    master.bind('<s>', lambda event=None: btn_sales.invoke())
    master.bind('<c>', lambda event=None: btn_shopper.invoke())
    master.bind('<C>', lambda event=None: btn_shopper.invoke())

    mainloop()


root = tk.Tk()
root.configure(bg="wheat")
root.geometry("400x400")
root.title("MCQ Creator")
width, height = pyautogui.size()
root.geometry('{}x{}+{}+{}'.format(1000, 800, 200, 100))
width, height = pyautogui.size()
root.geometry(
    '{}x{}+{}+{}'.format(int(width / 1.35), int(height / 1.25), int(width / 9), int(height / 12)))
root.configure(bg='AntiqueWhite1')

canvas_width, canvas_height = width, height
canvas = Canvas(root, width=canvas_width, height=canvas_height)
myimage = ImageTk.PhotoImage(PIL.Image.open("..\\image\\Geometry-Header-1920x1080.jpg").resize((width * 2, height * 2)))
canvas.create_image(0, 0, anchor=NW, image=myimage)
canvas.pack()
designMainScreen(root)
root.mainloop()

