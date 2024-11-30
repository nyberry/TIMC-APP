'''
This program allows users at TIMC to import pdf files, reads the data, generates a report and can send by email.

Modules:
1. read_pdfs.py: contains the code to read data from the different types of pdf.
2. write_xls.py: contains the code to populate and edit a Mircosoft Excel worksheet from a template.
3. write_docx.py: contains the code to generate a Microsoft Word document.
4. handle_files.py: contains the code to handle file and directory management.
 
data{} is a dictionary of the information for an individual patient.
   - data['(non test)'] has the structure {keyword: value}
   - data['(test)'] has the structure {keyword: [value (int or str), units (str), lower bound of normal (int or str), upper bound of normal (int or str)]}.  
'''

# CONSTANTS
VERSION = 'Version 2.0.4 15th May 2024'
TYPES_OF_DOCUMENTS = ["Null", "Medical report", "Letter with lab results", "Letter to patient", "Letter of referral", "Letter TWIMC"]  # the types of documents that can be written
MIN_BUTTON_WIDTH = 14
run_mode ='offline'  # 'user', 'developer', 'offline' , 'update'          


# IMPORTS
import os, sys, shutil, psutil, win32com.client, random
import xlwings, subprocess
import tkinter as tk
from tkinter import PhotoImage, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
from datetime import datetime
from time import sleep
from pywhatkit import sendwhatmsg_instantly
from pdf.handle_pdfs import  read_data_from_pdfs, encrypt_pdf_with_PyPDF2, get_pdfs_from_inbox, get_pdfs_from_desktop, attempt_pdf_import, add_attachment_pdfs_to_main_pdf,assign_pdf_hash,assign_hashed_pdf_filename
from word_generator.write_docx import write_word_document
from file_handling.handle_files import set_absolute_directory_path, set_filepaths_and_filenames
from file_handling.handle_files import check_if_templates_need_updating, read_TIMCusers_file, write_TIMCusers_file, delete_temp_files_and_close_temp_windows
from classes import *


# GLOBAL VARIABLES
pdfs = {}                      # a dictionary in the form key = hash: value = pdf class data structure
patients = {}                  # a dictonary of the form key = patient name: value = class patient data structure
data ={}
patient = None                 # this will be the active TIMC_patient class
document = None                # this will be the active document class
user = None
selected_patient = None       # the name of the chosen patient
selected_patients = []        # the names of patients selected from listbox  
word, excel = None, None                            # global variables for the open worksheet and word apps
workbook, sheet, doc = None, None, None             # global variables for the Excel and Word relative paths
temp_folder_window_titles= []        # a dynamically named folder to display temporary files



def main(run_mode):
    
    global user

    # Show Loading image
    load_window = show_loading_window()
 
    # Kill any Excel or Word objects that may be open, otherwise may crash when tries to open a new sheet
    quit_office_apps_without_saving()
    delete_temp_files_and_close_temp_windows(temp_folder_window_titles)

    # Identify current user
    user = identify_user(load_window)

    # set run mode
    arguments = sys.argv[1:]  # Exclude the script name
    if arguments != []:
        argument = arguments[0]
        if argument in ['developer', 'user', 'offline', 'update']:
            run_mode = str(arguments[0])
    print ("running in "+run_mode+" mode.")
    if run_mode != 'offline':    
        check_if_templates_need_updating(run_mode)
    
    # Run the GUI
    load_window.destroy()
    GUI()
    set_GUI_to_initial_state()
    if run_mode == 'user':
        check_inbox_button_pressed() 
    window.mainloop()

    # Exit button pressed, tidy up any files and close any open temporary windows
    delete_temp_files_and_close_temp_windows(temp_folder_window_titles)

def build_dict_of_unique_patients() -> dict:
    # This sub looks at all the categorized pdf files in the temp directory, and return a dict of patient names.
    # input will be a dictionary of TIMC_pdf instances
    # output will be a dictionary of TIMC_patient instances
    patients = {}
    for hash, pdf in pdfs.items():
        name = pdf.patient
        first_name= name.split()[0]
        surname= name.split()[1]
    
        # if this person does not already exist in the dictionary we are making, create a new entry:
        if name not in patients:
            patients[name]= TIMC_patient(None, first_name, surname, None, None, [], [], [], None, None, None)
        
        # append to the list of attachment filepaths if necessary
        if pdf.type.lower()=="medas dump":
            patients[name].medas_attachments.append(pdf.filepath)
        if pdf.type.lower()=="ml lab report" or pdf.type.lower()=="al arabi lab report":
            patients[name].lab_attachments.append(pdf.filepath) 
        if pdf.type.lower()=="other" or pdf.type.lower()=="timc document":
            patients[name].other_attachments.append(pdf.filepath) 

        # create the entry
        medas_filepaths = patients[name].medas_attachments
        lab_filepaths = patients[name].lab_attachments
        other_filepaths = patients[name].other_attachments
        patients[name].display_name=first_name+' '+surname+' ('+(len(medas_filepaths)*'Medas✔')+((len(lab_filepaths)!=0)*'Lab'+len(lab_filepaths)*'✔')+((len(other_filepaths)!=0)*'Other'+len(other_filepaths)*'✔')+')'
        
    #return the list
    return(patients)

def build_word_doc(document, patient, user, workbook, data):

    #global variables
    global doc

    #open Word template and generate the report, return it as doc, save it as temporary file to pass to win32com
    status_label.config(text=f'writing word document for {document.patient}...') 
    window.update()
    doc=write_word_document(document, patient, user, workbook, data)

    #save doc as a temporary file
    temp_dir=set_absolute_directory_path('temp')
    temp_filepath = os.path.join(temp_dir,"temp.docx")
    try:
        doc.save(temp_filepath)
    except Exception as e:
        print ("problem saving temp doc as",temp_filepath)
        exception_message = str(e)
        print(f"Exception message: {exception_message}")

    #open the document in word
    word = win32com.client.Dispatch("Word.Application")
    try:
        doc=word.Documents.Open(temp_filepath)
    except Exception as e:
        print("problem opening word document to display it")
        print("the document was ",data['docx draft filepath'])
        print(f"Exception message: {e}")

    if document.type == 'send by whatsapp':
        save_button_pressed()
    else:
        #make it visible and in the focus
        word.Visible = True
        # Check whether to email, save the report,  delete it or edit the XLS
        show_buttons(['delete_report','save'])
        status_label.config(text=f'{document.type} generated for {document.patient}. Please review the document.')
    if workbook:
        edit_excel_data_button.grid(row=4, column=1, columnspan=1)

def build_workbook(workbook,data):
    # updates the excel spreadsheet with the data extracted from the pdfs
    # and then call the function to parse a copy of that sheet                      

    sheet = workbook.sheets['Data']

    #call function to calculate first and last row of the bloods
    bloods_start_row, bloods_last_row = calculate_bloods_range(sheet)

    # iterate through the name cells in patinet data field. If the contents match an item in the data list, paste data to column D
    for row in range(1,bloods_start_row):
        cell = sheet.range(f'C{row}')
        term=cell.value  
        if term in data: 
            try:
                sheet.range(f'D{row}').value = data[term]
            except:
                print("problem with data",term)

    # iterate through the name cells in bloods fields column. If the contents match an item in the data list, paste data to column D,E,F,G
    for row in range(bloods_start_row,bloods_last_row+1):
        cell = sheet.range(f'C{row}')
        term=cell.value  
        if term in data: 
            try:
                sheet.range(f'D{row}').value = data[term][0]
                sheet.range(f'E{row}').value = data[term][1]
                sheet.range(f'F{row}').value = data[term][2]
                sheet.range(f'G{row}').value = data[term][3]
            except:
                print("problem with data",term)

    # if an examination is recorded, set the o/e fields to normal
    if "Examination" in data:
        try:
            sheet['Examination'].value = data["Examination"]
            sheet['CVS_exam'].value = "normal"
            sheet['RS_exam'].value = "normal"
            sheet['abdo_exam'].value = "normal"
            sheet['MSK_exam'].value = "normal"
            sheet['skin_exam'].value = "normal"
            sheet['prostate_exam'].value = "not done"
        except Exception as e:
            print (f'Error recording examination in spreadsheet {e}')

    #keep a backup of the loaded spreadsheet before parsing, in case we want to add more data later
    backup_filepath = os.path.abspath("../temp/HStemplateBackup.xlsm")
    workbook.api.SaveCopyAs(backup_filepath)
  
    #call the function to parse:
    parse_worksheet(workbook)
    return

def calculate_bloods_range(sheet):    
    # Search for the markers of the start and end of the bloods fields:    
    try:
        for row_num in range(1, sheet.cells.last_cell.row):
            cell_value = sheet.range(f'B{row_num}').value
            if cell_value == 'Group':
                bloods_start_row = row_num
            if cell_value == 'End of blood tests':
                bloods_last_row = row_num-1
                break
    except Exception as e:
        print ("Error identifying start and end of blood tests on spreadsheet",e)
        bloods_start_row = 10
        bloods_last_row = 200
    return (bloods_start_row,bloods_last_row)

def check_inbox_button_pressed():
    # checks outlook inbox and uploads any lab reports to the working directory.
    global pdfs              
  
    # count how many pdf files are currently in the temp directory:
    num_pdfs_before=len(pdfs)
    
    # import any new ones from inbox: 
    show_status('Checking Email Inbox','Scanning for pdf attachments')
    try:  
        pdfs = get_pdfs_from_inbox(pdfs)
        pdfs = get_pdfs_from_desktop(pdfs)
    except:
        show_status('','Error checking inbox')

    # count and display how many pdfs have been imported:
    num_files_imported=len(pdfs)-num_pdfs_before
    if num_files_imported:
        show_status('',f'{num_files_imported} pdf attachments found')
    else:
        show_status('','No new pdfs found in inbox')
    refresh_patient_list()

def check_sheet_for_new_data(sheet,data):
    # check the manually updated spreadsheet for new fields and values, updates data dictionary
    row = 3
    while sheet.range(f'B{row}').value != 'End of data fields':
        try:
            field_name = sheet.range(f'C{row}').value
            field_value = sheet.range(f'D{row}').value
        except:
            print("ERROR READING SPREADSHEET AT ROW",row)
        if field_name is not None:
            #check if in datetime format, cant have those in json dictionary
            if isinstance(field_value, datetime):
                    #so convert to string
                    field_value=field_value.strftime('%d %B %Y')  #need to convert datetime to string fpr json
            if field_name not in data and field_value is not None:
                # add new field and value to dictonary
                data[field_name]=field_value
        row+=1

    return(data)

def clear_all_buttons():
    # get rid of all buttons and dialogues, typically after button press, so that new options can be given
    continue_button.grid_remove()
    delete_report_button.grid_remove()
    generate_document_button.grid_remove()
    encrypt_button.grid_remove()
    edit_excel_data_button.grid_remove()
    check_inbox_button.grid_remove()
    clear_list_button.grid_remove()
    save_button.grid_remove()
    check_if_finished_button.grid_remove()
    send_email_button.grid_remove()
    merge_patients_button.grid_remove()
    report_error_button.grid_remove()
    quit_button.grid_remove()
    send_whatsapp_button.grid_remove()

def clear_list_button_pressed():
    global pdfs, patients
    delete_temp_files_and_close_temp_windows([])
    pdfs = {}
    patients = refresh_patient_list()

def edit_excel_data_button_pressed():

    #global variables
    global data, workbook

    # close word app, open Excel app
    quit_office_apps_without_saving()

    # Make a copy of the backup Excel workbook, connect to it
    template_filepath = os.path.abspath("../temp/HStemplateBackup.xlsm")
    workbook_filepath = os.path.abspath("../temp/HStemplate.xlsm")
    shutil.copy(template_filepath, workbook_filepath)

    xlapp = xlwings.App(visible=True)
    workbook = xlapp.books.open(workbook_filepath)

    # allow editing and wait for continue button         
    status_label.config(text="Review the worksheet and edit if necessary.")
    show_buttons(['continue'])
   
    # wait for a continue prompt
    window.wait_variable(continue_button_status)
    continue_button_status.set(False)                          #reset the button
    clear_all_buttons()

    #check the worksheet for any new data
    sheet = workbook.sheets['Data']
    data=check_sheet_for_new_data(sheet,data)

    #parse the sheet
    parse_worksheet(workbook)

    #build the building of new word document
    build_word_doc(document, patient, user, workbook ,data)
    
    return

def encrypt_button_pressed():
    # 1. open the pdf  2. read the pdf for data for password and filename  3. call the encrpt pdf function
    global data, pdfs, run_mode, selected_patient, document, patient
    document = TIMC_document('Pdf to encrypt',None,None,None,None,selected_patient, None, None)

    # call the function to read data from pdfs
    patient = patients[selected_patient]
    data=read_data_from_pdfs(patient, run_mode)
    patient.QID = data.get('QID')
    patient.email = data.get('Email')
    document.author = f'{user.title} {user.firstname} {user.surname}'
    document.password = patient.QID

    # set the filepaths and filenames for the sheet and report
    data=set_filepaths_and_filenames(data)
    temp_dir=set_absolute_directory_path('temp')

    for hash, pdf in pdfs.items():
        name = pdf.patient
        first_name= name.split()[0]
        surname= name.split()[1]
        filename = pdf.original_filename
        if first_name == patient.firstname and surname == patient.surname:
            unencrypted_filepath = pdf.filepath
            encrypted_filepath = data['pdf final filepath encrypted']
            print (f'trying to encrypt:\n  unencrypted_filepath: {unencrypted_filepath}\n encrypted_filepath: {encrypted_filepath}') 
            if encrypt_pdf(unencrypted_filepath, encrypted_filepath) == True:
                messagebox.showinfo('Document protected',f'Document saved with password {document.password}')
            else:
                messagebox.showerror('Error',f'Unable to encrypt document with a password')
            break

    # refresh display to give choice of sending by email or returning to start
    show_buttons(['send_email','check_if_finished','report_error'])

def encrypt_pdf(unencrypted_filepath, encrypted_filepath):
    global document, data
    # make a copy to be encrypted:
    try:
        shutil.copy(unencrypted_filepath, encrypted_filepath)
        print(f"Copy successful: {unencrypted_filepath} -> {encrypted_filepath}")
    except FileNotFoundError:
        print(f"File not found: {unencrypted_filepath}")
        return False
    except Exception as e:
        print(f"An error occurred: {e}")
        return False

    # assign a password
    if not document.password:
        password = simpledialog.askstring("Password", "Please enter a password for this document (usually QID):", show='*')
        document.password = password

    # encrypt the new copy of the pdf
    if document.password:
        try:
            encrypt_pdf_with_PyPDF2(data['pdf final filepath encrypted'],document.password)
            comment_line1 = "Report saved with password as " + data['pdf final filepath encrypted'] +"\n"
            comment_line2 = f'The password is: {document.password}\n'
            comment_line3 = "For data security, please use this encrypted copy if sending to patient, or outside of TIMC."+"\n"
            comment = comment_line1+comment_line2+comment_line3
            explanation_label.config(text=f'{data['pdf final filepath encrypted']}')
            status_label.config(text=f'Saved with password: {document.password}')
            print (comment)

            # Creating an information messagebox
            # messagebox.showinfo("Report saved", comment)

        except Exception as e:
            print ("problem encrypting pdf report", data['pdf final filepath'],": ",e)

            # delete the file named encrypted if it exists
            try:
                file_path = data['pdf final filepath encrypted']
                if os.path.isfile(file_path):
                    os.remove(file_path)
            except:
                print ("Error: Please note there may be a file saved as encrypted, which is not in fact encrypted.")
            return False
        
    else:
        print ("Could not encrypt pdf as no password in data")
        return False
    return True

def generate_document_button_pressed():
    # Asks user to select the type of report from a checklist

    # focus on this patient and remove all others from display
    global document, patients, patient   
    patient = patients[selected_patient]
    update_displayed_names([patient.display_name])
    show_status('',selected_patient)

    # function to create a new instance of class document
    def on_radiobutton_select():
        global document
        document = TIMC_document(None, None, None, None, None, None, None, None)
        document.type = selected_option.get()
        print (f'Type of document selected: {document.type}')
        document.author = f'{user.title} {user.firstname} {user.surname}'
        document.patient = selected_patient
        checkbox.destroy()
        generate_document()

    def cancel_patient_selection():
        checkbox.destroy()
        set_GUI_to_initial_state()

    # Create a checkbox
    checkbox = tk.Toplevel()
    checkbox.title("Select document")
    header_label = tk.Label(checkbox, text="Select type of document to generate")
    header_label.pack(padx = 15, pady = 10 )

    # Create variables to store the state of each checkbox
    selected_option = tk.StringVar()
    selected_option.set(None)
    
    # Create radio buttons for each option; only doctors can create reports
    button_options=TYPES_OF_DOCUMENTS 
    for item in range(1,len(button_options)):
        if (button_options[item] == 'Medical report' or button_options[item] == 'Letter of referral') and (user.role.lower() != 'doctor' and user.role.lower() != 'dr'):
            pass
        else:
            tk.Radiobutton(checkbox, text=button_options[item], variable = selected_option, value=button_options[item], command=on_radiobutton_select).pack(anchor='w')
    
    cancel_button = tk.Button(checkbox, text="Cancel", bg='orange', width = MIN_BUTTON_WIDTH, command = cancel_patient_selection)
    cancel_button.pack(pady = 10)

    # make the checkbox modal
    checkbox.grab_set()
    checkbox.wait_window()

def generate_document():
    #global variables
    global data, document, patient
    global pdfs
    global word, doc
    global workbook

    #clear the buttons, disable the listbox, and worklist refresh the GUI window
    clear_all_buttons()
    show_status(f'Generating a {document.type} for {document.patient}','Gathering data...')
    patient_listbox.config(state=tk.DISABLED)
    print (f"\n< Generating a {document.type} for {document.patient} >")

    # call the function to read data from pdfs
    patient = patients[document.patient]
    data=read_data_from_pdfs(patient,run_mode)

    # update the patient instance of TIMC_patient class
    patient.title = data.get('Title')
    patient.email = data.get('Email')
    patient.QID = data.get('QID')
    patient.phone = data.get('Phone')
    patient.sex = data.get('Sex')

    # assign a password to the document
    document.password = patient.QID
    document.date = data['Date of report']

    # set the filepaths and filenames for the sheet and report
    data=set_filepaths_and_filenames(data)

    # If we are writing a report, open the XLS spreadsheet and import the data
    if document.type == "Medical report":
        status_label.config(text=f'updating spreadsheet for {document.patient}...')
        window.update()

        # Make a copy of the Excel workbook, connect to it, update it with data
        template_filepath = os.path.abspath("../data/HStemplate.xlsm")
        workbook_filepath = os.path.abspath("../temp/HStemplate.xlsm")
        shutil.copy(template_filepath, workbook_filepath)

        # Open the Excel application
        xlapp = xlwings.App(visible=False)

        # Open the Excel workbook
        workbook = xlapp.books.open(workbook_filepath)
        build_workbook(workbook,data)

        # Move on to building the word document
        build_word_doc(document, patient, user, workbook ,data)

        #Save the modified workbook, close it, and quit Excel
        workbook.save()
        workbook.close()
        xlapp.quit()

    else:
        build_word_doc(document, patient, user, None, data)

    print ("Finished building word document\n")
    return

def GUI():
    # define window, labels, listboxes, status flags, and buttons as global variables
    global window
    global top_banner, explanation_label, status_label, patient_listbox, continue_button_status
    global generate_document_button,  encrypt_button, edit_excel_data_button, clear_list_button, save_button
    global check_if_finished_button, send_email_button, send_whatsapp_button
    global merge_patients_button, report_error_button, check_inbox_button, continue_button, delete_report_button, quit_button

    window = TkinterDnD.Tk()
    window.title("TIMC Report Generator")

    #Top and bottom banners
    images_dir=set_absolute_directory_path('images')
    top_banner_image=PhotoImage(file=images_dir+"\\banners\\"+"top banner 640.png")
    bottom_banner_image=PhotoImage(file=images_dir+"\\banners\\"+"plain bottom banner 640.png")
    top_banner = tk.Label(window, image=top_banner_image)
    top_banner.grid(row=0, column=0, columnspan=5)  # Centered label
    bottom_banner = tk.Label(window, image=bottom_banner_image)
    bottom_banner.grid(row=6, column=0, columnspan=5)

    # Create labels for explanation and status text
    explanation_label = tk.Label(window)
    explanation_label.grid(row=1, column=1, columnspan=3)
    status_label = tk.Label(window, pady= 5)

    #Create a canvas to hold the patient listbox within a frame
    canvas = tk.Canvas(window)
    canvas.grid(row=2, column=1, columnspan=3)
    listbox_frame = tk.Frame(window, width = 60, padx = 5, pady= 5)
    listbox_frame.grid(row=2, column=1, columnspan=3)
    patient_listbox = tk.Listbox(listbox_frame, width=60, selectmode=tk.MULTIPLE)

    canvas.drop_target_register(DND_FILES)
    canvas.dnd_bind('<<Drop>>', on_drop)
    canvas.place(x=0, y=explanation_label.winfo_height(), width=patient_listbox.winfo_width(), height=patient_listbox.winfo_height())
    patient_listbox.bind("<ButtonRelease-1>", on_patient_select)
    patient_listbox.drop_target_register(DND_FILES)
    patient_listbox.dnd_bind('<<Drop>>', on_drop)

    # Create buttons
    check_inbox_button = tk.Button(window, text="Check inbox",  command=check_inbox_button_pressed, bg = 'white', width = MIN_BUTTON_WIDTH)
    clear_list_button = tk.Button(window, text="Clear list", command=clear_list_button_pressed, bg = 'white', width = MIN_BUTTON_WIDTH)
    quit_button = tk.Button(window, text="Exit", bg = 'white', command=lambda: window.destroy(), width = MIN_BUTTON_WIDTH)
    generate_document_button = tk.Button(window, text = "Letter/ Report", bg="light blue",  command=generate_document_button_pressed, width=MIN_BUTTON_WIDTH)
    encrypt_button = tk.Button(window, text="encrypt", bg="light grey",  command=encrypt_button_pressed,  width = MIN_BUTTON_WIDTH)
    merge_patients_button = tk.Button(window, text="merge", bg="light blue",  command=merge_patients_button_pressed, width=MIN_BUTTON_WIDTH)
    edit_excel_data_button = tk.Button(window,text="Excel", bg='light blue', command = edit_excel_data_button_pressed, width = MIN_BUTTON_WIDTH)
    delete_report_button = tk.Button(window,text="Discard", bg='orange', command = lambda:start_afresh(), width = MIN_BUTTON_WIDTH)
    continue_button_status = tk.BooleanVar()
    continue_button_status.set(False)
    continue_button = tk.Button(window, text="Continue", bg='light green', command = lambda: continue_button_status.set(True), width = MIN_BUTTON_WIDTH)
    save_button = tk.Button(window, text="Save", bg='light green', command=save_button_pressed, width=MIN_BUTTON_WIDTH)
    send_email_button = tk.Button(window, text="Email", bg='yellow', command=send_email_button_pressed, width=MIN_BUTTON_WIDTH)
    report_error_button = tk.Button(window, text="Report Bug", bg = 'orange', command = report_error_button_pressed, width= MIN_BUTTON_WIDTH)
    send_whatsapp_button = tk.Button(window, text="WhatsApp", bg = 'light green', command = send_whatsapp_button_pressed, width= MIN_BUTTON_WIDTH)
    check_if_finished_button = tk.Button(window, text="Finished", bg='light green', command=start_afresh, width = MIN_BUTTON_WIDTH)

    # set the  minimum row heights to get some space bewteen buttons
    rows_to_configure = [1, 3, 4, 5]
    min_sizes = [20, 20, 30, 30]
    for row, min_size in zip(rows_to_configure, min_sizes):
        window.grid_rowconfigure(row, minsize=min_size)
   
def identify_user(load_window):

    # function to ask user to input their name
    def get_user_details(current_user_login, user):
        while True:
            user.firstname = simpledialog.askstring("New User", f'New user {current_user_login} detected. Please enter your First Name:',parent = load_window)
            if len(user.firstname)>=1:
                break

        while True:   
            user.surname = simpledialog.askstring("New User", "Please enter your surname:", parent = load_window)
            if len(user.surname)>=1:
                break

        while True:
            user.title = simpledialog.askstring("New User", "Please enter your title (like Dr, Nurse, Mr or Mrs):", parent = load_window)
            if user.title.lower().strip() in ['doctor','dr','nurse','rn','rgn','sister','dentist','physio','admin','manager','mr','mrs','ms''miss']:
                break

        # assign role
        if user.title.lower().strip()=="dr" or user.title.lower().strip()=="doctor":
            user.role = 'doctor'
        elif user.title.lower().strip()=="nurse" or user.title.lower().strip()=="rn" or user.title.lower().strip()=="rgn" or user.title.lower().strip()=="sister":
            user.role = 'nurse'
        else:
            user.role = user.title
        return(user)

     # retrieve the user list from file
    try:
        users = read_TIMCusers_file()
    except:
        print ('unable to find users list file so creating a new list')
        users = {}
    
    # retrieve the details of the currenty logged in user
    user = TIMC_user(None, None, None, None, None, None)
    try:
        current_user_login = os.getlogin()
    except:
        print('problem with getting current user login')
        current_user_login = "Unknown user"

    # if the user is known, get details
    if current_user_login in users:
        try:
            user.title = users[current_user_login][0]
            user.firstname = users[current_user_login][1]
            user.surname = users[current_user_login][2]
            user.role = users[current_user_login][3]
            user.stamp = users[current_user_login][4]
            user.banner = users[current_user_login][5]

            print(f'User {user.title} {user.firstname} {user.surname} logged in.')
        except Exception as e:
            print(f'problem getting details of {current_user_login}: {e}')

    else:
        # ask for user details text input
        while user.surname== None or user.firstname== None:
            try:
                user = get_user_details(current_user_login, user)
                if user.surname and user.firstname:
                    users[current_user_login]=[user.title, user.firstname, user.surname, user.role, "", ""]
                    print(f'New user added {user.title} {user.firstname} {user.surname}')
                else:
                    print(f'Invalid details')
            except Exception as e:
                print(f'problem getting user details and extending userlist {e}')
        
         #save updated userlisr
        try:
            write_TIMCusers_file(users)
        except Exception as e:
            print(f'problem writing userlist {e}')
    return (user)

def merge_patients_button_pressed():
  
    global pdfs, selected_patient, selected_patients

    # function for selecting from checklist & then destroying the checklist
    def on_radiobutton_select():
        global selected_patient
        selected_patient=selected_choice_for_merge.get()
        print("We are going to assign all these documents to the name:", selected_patient)
        checkbox.destroy()

    def on_cancel_merge():
        global selected_patient
        selected_patient = None
        print("Merge cancelled")
        checkbox.destroy()
    
    # determine how many patients we are merging
    number_of_patients_to_merge = len(selected_patients)
    print ()
    print ("Merge function activated.")
    print ("There are ",number_of_patients_to_merge," patients selected to merge")
    print ("These are:", selected_patients)

    # show a checkbox
    print ("<Show a check box asking which patient name to keep>")
    checkbox = tk.Toplevel()  # Use Toplevel instead of Tk
    checkbox.title("Choose the correct name")
    header_label = tk.Label(checkbox, text=f'You may wish to merge documents if they belong to the same person, but one name has a typo.\n Which is the correct spelling of their name?')
    header_label.pack()

    # Create variables to store the state of each checkbox
    selected_choice_for_merge = tk.StringVar()
    selected_choice_for_merge.set(None)
    
    # Create Checkbuttons with two or three choices
    radiobutton1 = tk.Radiobutton(checkbox, text=selected_patients[0], variable=selected_choice_for_merge, value=selected_patients[0],command=on_radiobutton_select)
    radiobutton2 = tk.Radiobutton(checkbox, text=selected_patients[1], variable=selected_choice_for_merge, value=selected_patients[1],command=on_radiobutton_select)
    if number_of_patients_to_merge == 3:
        radiobutton3 = tk.Radiobutton(checkbox, text=selected_patients[2], variable=selected_choice_for_merge, value=selected_patients[2],command=on_radiobutton_select)
    
    # Pack the Checkbuttons into the window
    radiobutton1.pack()
    radiobutton2.pack()
    if number_of_patients_to_merge == 3:
        radiobutton3.pack()
    
    # make a cancel merge button
    cancel_button = tk.Button(checkbox, text="Cancel", bg='orange', width = MIN_BUTTON_WIDTH, command = on_cancel_merge)
    cancel_button.pack(pady = 10)    

    # make the checkbox modal
    checkbox.grab_set()
    checkbox.wait_window()

    # A selection has now been made, so tidy the list
    # iterate through all pdf files in pdfs. The dictionary has the form {hash:[filename,filetype, first_name, surname]}
    # if any of the first_name and surname combos match those to be trimmed, replace them with the chosen name.

    # work out which patients need to be reassigned
    if selected_patient:
        try:
            patients_to_reassign=[]
            for patient in selected_patients:
                if patient != selected_patient:
                    patients_to_reassign.append(patient)
            print ("The following patients need to be reassigned:",patients_to_reassign)

            for pdf in pdfs:
                name= pdfs[pdf].patient

                # check if that patient's files need to be reassigned
                if name in patients_to_reassign:
                    print("report currently assigned to "+ name+ " will be reassigned to "+selected_patient)
                    selected_patient_names = selected_patient.split()
                    selected_patient_first_name = selected_patient_names[0]
                    selected_patient_surname = selected_patient_names[1]

                    #reassign
                    pdfs[pdf].patient =selected_patient
            messagebox.showinfo('Documents merged', f'Documents merged and now all belong to {selected_patient}')
        except:
            messagebox.showerror('Error',f'It was not possible to merge those documents')
        print ()
        
    #refresh the list and display
    refresh_patient_list()
    status_label.config(text="", anchor="center")
    selected_patients=[]
    selected_patient=None

def on_drop(event):
    # this code handles what to do when a file is dragged and dropped into the listbox frame

    # first check the list box is not disabled
    listbox_state = patient_listbox.cget("state")
    if listbox_state!="disabled":
        
        # declare global variable:
        global pdfs
        
        # count how many pdf files are currently in the temp directory:
        num_pdfs_before=len(pdfs)

        # build a list of dropped file paths
        dropped_files = event.data
        file_list = window.tk.splitlist(dropped_files)
        number_of_files_dropped = len(file_list)
        print (f'\n{str(number_of_files_dropped)} file(s) dropped. Checking files...\n')

        # check if the file is a pdf
        for original_filepath in file_list:
            if original_filepath.lower().endswith('.pdf'):
                try:
                    # save a copy of the attachment in the temp folder
                    temp_dir= set_absolute_directory_path('temp')
                    temp_path = os.path.join(temp_dir, 'temp.pdf')
                    shutil.copy(original_filepath, temp_path)
                    print(f'file copied to {temp_path}')

                    #create an empty instance of TIMC_pdf
                    pdf = TIMC_pdf(temp_path, original_filepath, None, None, None)
                    print("checking: "+ original_filepath)
                    hash_value = assign_pdf_hash(pdf)

                    #skip if the attachment already exists in the temp folder
                    if hash_value and hash_value in pdfs:
                        print (">> this pdf already exists in the directory so not imported")
                        break
                    
                    # this pdf doesn't exist so attempt to populate the instance 
                    pdf = attempt_pdf_import(pdf)

                    # assign a filename
                    pdf.filepath = assign_hashed_pdf_filename(pdf, hash_value)

                    # if no patient, assign qid or "Unknown"
                    if not pdf.patient:
                        pdf.patient = "Unknown Patient"

                    # add to pdf dictionary, if valid (succesfully populated with name and type)
                    if pdf.patient and pdf.type and pdf.filepath:
                        pdfs[hash_value]=pdf

                except Exception as e:
                    print (f'Problem attempting to import pdf file: {original_filepath} , {e}')

                # tidy up by deleting the temporary file, if it still exists
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except:
                        # may have crashed if deleted too soon???
                        sleep(1)
                        try:
                            os.remove(temp_path)
                        except:
                            print (f"error deleting {temp_path}")      

        # count how many pdfs have been imported:
        num_files_imported=len(pdfs)-num_pdfs_before
        if num_files_imported ==0:
                status_label.config(text="no valid files imported")
        else:
                status_label.config(text=str(num_files_imported)+" documents imported")

        # update the list of patients on the screen and return this list as pdf_filedict
        refresh_patient_list()

def on_patient_select(event):

    global selected_patients, selected_patient
    
    # first check the list box is not disabled
    listbox_state = patient_listbox.cget("state")
    if listbox_state!="disabled":
        # if one or more patients are selected:
        if patient_listbox.curselection():
            # build a list of the names selected 
            selected_display_names = [patient_listbox.get(index) for index in patient_listbox.curselection()]     
            selected_patients=[]
            for name in selected_display_names:
                words=name.split()
                selected_patients.append(words[0]+' '+words[1])
            number_of_selected_patients=len(selected_patients)

            # if just 1 selected, show generate report button
            if number_of_selected_patients == 1:
                selected_patient = selected_patients[0]
                show_buttons(['send_whatsapp', 'generate_document','quit', 'encrypt'])
                show_status('Select a patient, or drag new pdfs into the window', f'{selected_patient}')
                
            # if 2 or 3 patients selected, show merge patients buttons
            if number_of_selected_patients >=2 and number_of_selected_patients <=3:
                merge_patients_button.config(state=tk.NORMAL, bg = 'goldenrod', text=f"Merge")
                string_of_patients_to_merge = [f'{name} ' for name in selected_patients]
                show_buttons(['merge_patients','quit']) 
                show_status(f'Merge the documents of {len(selected_patients)} patients',f'{str(string_of_patients_to_merge)}')

            # if 4 or more patients selected
            if number_of_selected_patients >=4:
                merge_patients_button.config(state=tk.DISABLED, bg='light grey', text=f"Merge") 
                show_buttons(['merge_patients','quit'])
                show_status("please select no more than three patients","Too many patients are selected")
                                    
        # if no patients are selected:
        else:
            selected_patient=None
            show_buttons(['check_inbox','clear_list','quit'])
            show_status("Select a patient from the list","")

def parse_worksheet(workbook):

    sheet = workbook.sheets['Data']

    #call function to calculate first and last row of the bloods
    bloods_start_row, bloods_last_row = calculate_bloods_range(sheet)

    # tidy up by deleting empty rows
    sheet= workbook.sheets['Data']
    row = bloods_start_row+1
    while sheet.range(f'B{row}').value!="End of data fields":
        cellC = sheet.range(f'C{row}')
        cellD = sheet.range(f'D{row}')
        cellAboveC = sheet.range(f'C{row-1}')
        cellAboveD = sheet.range(f'D{row-1}')  
        if cellD.value == None and (cellC.value != None or (cellAboveC.value== None and cellAboveD.value== None)):
            sheet.range(f'A{row}').api.EntireRow.Delete()
        else:
            row+=1
            
    #tidy up by deleting duplicate headers from column B
    new_header=""
    for row in range(bloods_start_row, bloods_last_row+1):
        if sheet.range(f"B{row}").value == new_header:
            sheet.range(f"B{row}").value=""
        else: 
            new_header = sheet.range(f"B{row}").value

    #now the sheet is ready for manual input, so return
    return

def quit_office_apps_without_saving():
    
    # function to close the excel app
    def quit_excel_app_without_saving():
        #close and exit excel without saving file
        excel_objects_flag=False
        try:
            # Close Excel application
            for proc in psutil.process_iter(attrs=['pid', 'name']):
                if 'EXCEL.EXE' in proc.info['name']:
                    excel_objects_flag=True
                    os.kill(proc.info['pid'], 9)
            xl_app = win32com.client.Dispatch("Excel.Application")
            xl_app.Quit()
            if excel_objects_flag==True:
                print("Excel application closed with unsaved work discarded.")
        except Exception as e:
            print(f"An error occurred: {str(e)}")
    
    # function to close the word app
    def quit_word_app_without_saving():
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Quit()
            print("Word applications closed with unsaved work discarded.")
        except Exception as e:
            print(f"An error occurred: {str(e)}")

    quit_excel_app_without_saving()
    quit_word_app_without_saving()

def refresh_top_banner(user):
    try:
        if user.banner:
            top_banner_image=PhotoImage(file=f"../images/banners/{user.banner}")
        else:
            top_banner_image=PhotoImage(file="../images/banners/top banner 640.png")
        top_banner.config(image=top_banner_image)
        top_banner.image=top_banner_image
        window.update()
    except Exception as e:
        print("problem updating banner {e}")
    return()

def refresh_patient_list():
    global pdfs, patients
    patients = build_dict_of_unique_patients()
    display_names = sorted([patients[patient].display_name for patient in patients])
    update_displayed_names(display_names)                      
    show_buttons(['check_inbox','clear_list','quit'])                             

def report_error_button_pressed():
    user_input = simpledialog.askstring("Send bug report",'Thanks! This app is always under development. If you noticed a bug, please let me know by describing it here:')
    if user_input:
        #send the email to nick:
        try:
            # Create an instance of Outlook
            outlook = win32com.client.Dispatch('Outlook.Application')

            # Create a new email
            mail = outlook.CreateItem(0)

            # Set the email properties
            mail.Sender = f'{data["Doctor"]}'
            mail.Subject = f'TIMC APP bug report from {data["Doctor"]}'
            mail.Body = user_input
            mail.To = 'nckbrry@gmail.com'

            # Send the email
            mail.Send()

            # update status label
            messagebox.showinfo("Bug Report Sent",f'Bug report sent by email to nckbrry@gmail.com')
            status_label.config(text='bug report sent by email to nckbrry@gmail.com')
        
        except Exception as e:
            messagebox.showinfo("Bug Report Not Sent",f'Sorry, could not send bug report sent by email to nckbrry@gmail.com , {str(e)}. Please could you let me know manually?')

         # Quit Outlook application
        try:
            outlook.Quit()
        except:
            pass 

    show_buttons(['check_if_finished', 'quit'])

def set_GUI_to_initial_state():
    global selected_patient
    # deselect any patient(s)
    patient_listbox.config(state=tk.NORMAL)
    if selected_patient:
        selected_patient = None
    selected_indices = patient_listbox.curselection()
    for index in selected_indices:
        patient_listbox.selection_clear(index)
    #return window to initial state
    show_buttons(['check_inbox', 'clear_list','quit'])
    status_label.config(text = "")
    refresh_top_banner(user)
    refresh_patient_list()
    
def save_button_pressed():

    # define global variable
    global doc, document, temp_folder_window_title
  
    clear_all_buttons()
    
    # this is the function to save the final report.
    # It is saved in 3 forms: DOCX, pdf, and encrypted pdf
    # any pdf lab files are added to the end of the report
    # then check if it should be emailed, or quit.
 
    #save a final copy of docx
    try:
        doc.SaveAs(data['docx final filepath'])                 
        print ("Report saved as",data['docx final filepath'])
    except Exception as e:
        print ("problem saving report as", data['docx final filepath'],":",e)

    # save an unencrypted pdf without attachments
    try:
        doc.SaveAs(data['pdf final filepath'], FileFormat=17)       #FileFormat 17 is for PDF
        print("Report saved without password as ",data['pdf final filepath'])
    except Exception as e:
        print ("problem saving final report as", data['pdf final filepath'],": ",e)
    
    # forget the doc
    doc = None

    # add any pdf lab reports to the unencrypted pdf
    add_attachment_pdfs_to_main_pdf(data['pdf final filepath'], patients[selected_patient], document.type)

    # save an encrypted pdf, if there is a password already
    unencrypted_filepath = data['pdf final filepath']
    if document.password:
        encrypted_filepath = data['pdf final filepath encrypted']
        encrypt_pdf(unencrypted_filepath, encrypted_filepath)

    # refresh display to give choice of sending by email or returning to start
    show_buttons(['send_email', 'check_if_finished', 'send_whatsapp'])
    show_status (document.description, 'report saved')

    # create a temporary folder with the documents for use
    try:
        temp_folder_window_title= document.description
        temporary_final_document_folder_path = set_absolute_directory_path(f"temp/{temp_folder_window_title}")
        temp_folder_window_titles.append(temp_folder_window_title)

        # copy the documents in 
        temp_unencrypted_filename = f'{document.description} {document.date}.pdf'
        temp_unencrypted_filepath = os.path.join(temporary_final_document_folder_path, temp_unencrypted_filename)
        shutil.copy(unencrypted_filepath, temp_unencrypted_filepath)
        if document.password:
            temp_encrypted_filename = "Password protected document.pdf"
            temp_encrypted_filepath = os.path.join(temporary_final_document_folder_path, temp_encrypted_filename)
            shutil.copy(encrypted_filepath, temp_encrypted_filepath)
        
        # display that folder in a window
        command = 'explorer "{}"'.format(temporary_final_document_folder_path)
        subprocess.Popen(command, shell=True)

    except Exception as e:
        print (f'Problem displaying temporary folder {e}')
    
def send_email(email, password_var, password, message):
    patient.email = email
    document.password = password
    # first, save a new encrypted version of the file in case the password has changed
    if password_var == "yes":
        unencrypted_filepath = data['pdf final filepath']
        encrypted_filepath = data['pdf final filepath encrypted']
        encrypt_pdf(unencrypted_filepath, encrypted_filepath)   

    try:
        # Create an instance of Outlook
        outlook = win32com.client.Dispatch('Outlook.Application')

        # Create a new email
        mail = outlook.CreateItem(0)

        # Set the email properties
        mail.Sender = "info@theimcentre.com"
        mail.Subject = 'Message from The International Medical Centre'
        mail.Body = message
        if password_var == "yes":
            mail.Attachments.Add(data['pdf final filepath encrypted'])
            print ('Sending password protected copy by email')
        else:
            mail.Attachments.Add(data['pdf final filepath'])
            print ('sending unprotected copy by email')
        mail.To = email

        # Send the email
        mail.Send()
    
    except Exception as e:
        print (f'Error sending email {e}')
        return False
    
    # Quit Outlook application
    try:
        outlook.Quit()
    except:
        pass
    
    show_buttons(['check_if_finished', 'report_error', 'send_email'])
    return True     

def send_email_button_pressed():
    # save the file and then send out as email

    global patient, document, user

    def confirm_email_details():
        # check that the email address is valid
        if not email_var.get() or '@' not in email_var.get():
            messagebox.showerror("Invalid email","Please provide a valid email address")
            return
        # check that a password exists, if needed
        if not showpassword_var.get() and password_var.get() == "yes":
            messagebox.showerror("No password","Please provide a password")
            return
        # call the function to send the email
        if send_email(email_var.get(), password_var.get(), showpassword_var.get(), showmessage_text.get("1.0", tk.END)) == True:
            messagebox.showinfo('Email sent', f'Email sent to {email_var.get()}')
        else:
            messagebox.showerror('Error',f'Unable to send email to {email_var.get}')
        # Close email dialogue window and return
        dialogue.destroy()

    def turn_password_on():
        showpassword_var.set(document.password),
        new_message = message_if_password
        display_body_of_email(new_message)

    def turn_password_off():
        showpassword_var.set(None),
        new_message = message_if_no_password
        display_body_of_email(new_message)

    def display_body_of_email(message):    
        showmessage_text.delete(1.0 , tk.END)
        showmessage_text.insert(1.0, message)  # Insert the message into the Text widget

    # Asks user to select the type of report from a checklist
    # Create a window
    dialogue = tk.Toplevel()
    dialogue.title("Send document by email")
    header_label = tk.Label(dialogue, text=f'Send by Email to {patient.firstname} {patient.surname}')
    header_label.grid()

    # Create a frame for text entry fields
    entry_frame = tk.Frame(dialogue)
    entry_frame.grid(padx=10, pady=10)

    # Confirm Email Address
    email_label = tk.Label(entry_frame, text="Email:")
    email_label.grid(row=0, column=0, sticky="w")
    email_var = tk.StringVar()
    email_entry = tk.Entry(entry_frame, textvariable=email_var)
    email_entry.grid(row=0, column=1)
    if patient.email:
        email_var.set(patient.email)

    # Confirm Password
    showpassword_label = tk.Label(entry_frame, text="Password:")
    showpassword_label.grid(row=2, column=0, sticky="w")
    showpassword_var = tk.StringVar()
    showpassword_entry = tk.Entry(entry_frame, textvariable=showpassword_var)
    showpassword_entry.grid(row=2, column=1)
   

    # Message for email 
    footer = f'\nThe International Medical Centre\nTel: +974 4488 4292 / +974 6644 4282\nEmail: info@theimcentre.com\nWeb:  www.theimcentre.com\n14 Sahat Street, Jelaiah, Duhail\nPO Box 19941, Doha, Qatar\n'
    message_if_password = f'Dear {patient.firstname},\n\nA message is attached. The password to open the attachment is your national identification number (QID).\nPlease let me know if you have any difficulty opening it.\n\nThanks,\n{user.title} {user.firstname} {user.surname}'+footer
    message_if_no_password = f'Dear {patient.firstname},\n\nA message is attached.\n\nThanks,\n{user.title} {user.firstname} {user.surname}' + footer
    message = message_if_password
    
    # Display the message
    showmessage_text= tk.Text(dialogue)
    showmessage_text.grid(row=3, column=0, padx = 10, pady = 10)
    showmessage_text.configure(state="normal", wrap = tk.WORD)  # Allow editing
    showmessage_text.insert(1.0, message)


    # password yes or no radiobuttons
    password_var = tk.StringVar()
    password_label = tk.Label(entry_frame, text="Use password:")
    password_label.grid(row=1, column=0, sticky="w")
    yes_radio = tk.Radiobutton(entry_frame, text="Yes", variable=password_var, value="yes", command = turn_password_on)
    yes_radio.grid(row=1, column=1, sticky="w")
    no_radio = tk.Radiobutton(entry_frame, text="No", variable=password_var, value="no", command = turn_password_off)
    no_radio.grid(row=1, column=2, sticky="w")
    if document.password:
        password_var.set("yes")
        showpassword_var.set(document.password)
    else:
        password_var.set("no")

    # Confirm Button
    submit_button = tk.Button(dialogue, text="Confirm", bg = 'light green', width = MIN_BUTTON_WIDTH, command = confirm_email_details)
    submit_button.grid(row =4, pady = 10)

    # Cancel button
    cancel_button = tk.Button(dialogue, text="Cancel", bg='orange', width = MIN_BUTTON_WIDTH, command = lambda:dialogue.destroy())
    cancel_button.grid(row =5, pady = 10)

    # make the checkbox modal
    dialogue.grab_set()
    dialogue.wait_window()

def send_whatsapp_button_pressed():
    global user, data, document, patient, pdfs
    
    if not document:
        patient = patients[selected_patient]
        update_displayed_names([patient.display_name])
        show_status('',selected_patient)
        document = TIMC_document(None, None, None, None, None, None, None, None)
        document.type = "send by whatsapp"
        document.author = f'{user.title} {user.firstname} {user.surname}'
        document.password = patient.QID
        document.patient = selected_patient
        generate_document()
    
    print (f"\n< Send a WhatsApp message to {document.patient} >")

    # call the function to read data from pdfs
    data=read_data_from_pdfs(patient,run_mode)
    patient.phone = data.get('Phone')
    
    def confirm_phone_details():
        if not phone_var.get() or len(phone_var.get())<8:
            messagebox.showerror("Invalid phone number","Please provide a valid phone number")
            return
        try:
            phone_number = phone_var.get()
            message = showmessage_text.get("1.0", tk.END)
            sendwhatmsg_instantly(phone_number, message)
            messagebox.showinfo(f'Send Whatsapp to {patient.firstname} {phone_number}', f"Please send the message directly from the Whatsapp window.\nDon't forget to add any attachments, by dragging from the temporary folder window.\nThis window will disappear when you close this dialogue box but any documents will still be availabe in the 'reports' folder.")
        except:
            messagebox.showerror('Error',f'Unable to send Whatsapp to {patient.firstname} {phone_number}')
        dialogue.destroy()
        
    # Create a window
    dialogue = tk.Toplevel()
    dialogue.title("Send WhatsApp message")
    header_label = tk.Label(dialogue, text=f'Send WhatsApp to {patient.firstname}')
    header_label.grid()

    # Create a frame for text entry fields
    entry_frame = tk.Frame(dialogue)
    entry_frame.grid(padx=10, pady=10)

    # Confirm Phone number
    phone_label = tk.Label(entry_frame, text="Phone:")
    phone_label.grid(row=0, column=0, sticky="w")
    phone_var = tk.StringVar()
    phone_entry = tk.Entry(entry_frame, textvariable=phone_var)
    phone_entry.grid(row=0, column=1)
    if patient.phone:
        phone_var.set(patient.phone)

    # Confirm Message
    header = f'Hi {patient.firstname},\n\nA message is attached. The password is your QID. Please let us know if any problems.\n\n'
    footer = f'Thanks, {user.title} {user.firstname} {user.surname}\nThe International Medical Centre'
    message = header + footer
    showmessage_text = tk.Text(dialogue, width=30, height=18, padx=10,pady=10, wrap = tk.WORD, bg="#DCF8C6")
    showmessage_text.grid(row=3, column=0, padx = 10, pady = 10 )
    showmessage_text.insert(tk.END, message)  # Insert the message into the Text widget
    showmessage_text.configure(state="normal")  # Allow editing
    showmessage_text.focus_set()
    showmessage_text.mark_set("insert", "2.0")
  
    # Confirm Button
    submit_button = tk.Button(dialogue, text="Send Whatsapp", bg = 'light green', width = MIN_BUTTON_WIDTH, command = confirm_phone_details)
    submit_button.grid(row =4, pady = 5)

    # Cancel button
    cancel_button = tk.Button(dialogue, text="Cancel", bg='orange', width = MIN_BUTTON_WIDTH, command = lambda:dialogue.destroy())
    cancel_button.grid(row =5, pady = 10)

    # make the checkbox modal
    dialogue.grab_set()
    dialogue.wait_window()

    # when the whatsapp has been sent
    start_afresh()

def show_buttons(buttons):

    clear_all_buttons()
    if 'continue' in buttons:
        continue_button.grid(row=5, column=1, columnspan=3)
    if 'delete_report' in buttons:
        delete_report_button.grid(row=4, column=3, columnspan=1)
    if 'generate_document' in buttons:
        generate_document_button.grid(row=4, column=2, columnspan=1)
    if 'encrypt' in buttons: 
        encrypt_button.grid(row=5, column=2, columnspan=1)
    if 'edit_excel_data' in buttons:
        edit_excel_data_button.grid(row=4, column=1, columnspan=1)
    if 'check_inbox' in buttons:
        check_inbox_button.grid(row=4, column=1, columnspan=1)
    if 'clear_list' in buttons:
        clear_list_button.grid(row=4, column=2, columnspan=1)
    if 'save' in buttons:
        save_button.grid(row=4, column=2, columnspan=1)
    if 'check_if_finished' in buttons:
        check_if_finished_button.grid(row=4, column=3, columnspan=1)
    if 'send_email' in buttons:
        send_email_button.grid(row=4, column=2, columnspan=1)
    if 'merge_patients' in buttons:
        merge_patients_button.grid(row=4, column=2, columnspan=1)
    if 'report_error' in buttons:
        report_error_button.grid(row = 4, column =2, columnspan=1)
    if 'send_whatsapp' in buttons:
        send_whatsapp_button.grid(row = 4, column =1, columnspan=1)
    if 'quit' in buttons:
        quit_button.grid(row = 4, column = 3, columnspan = 1)
    window.update()

def show_status(explanation, status):
    explanation_label.config(text=explanation)
    status_label.config(text=status)
    explanation_label.grid(row=1, column=1, columnspan=3)
    status_label.grid(row=3, column=0, columnspan=5)

def show_loading_window():
    try:
        load_window = tk.Tk()
        load_window.title("Loading...")
        # Create a Label widget to display the image
        images_dir=set_absolute_directory_path('images')
        banner_image = tk.PhotoImage(file = f'{images_dir}\\banners\\top banner 640.png')
        banner = tk.Label(load_window, image=banner_image)
        banner.pack()
        loading_image_files = os.listdir(f'{images_dir}\\loading')
        image_file = random.choice(loading_image_files)
        loading_image = tk.PhotoImage(file = f'{images_dir}\\loading\\{image_file}')
        photo = tk.Label(load_window, image=loading_image)
        photo.pack()
        load_status = tk.Label(load_window, text = 'Loading...' )
        load_status.pack()
        load_window.update()
    except Exception as e:
        print(f'problem showing loading window {e}')
    return (load_window)

def start_afresh():
    # refresh data dictionary and forget document instance
    global data, document
    data  = {}
    document = None
    quit_office_apps_without_saving()
    set_GUI_to_initial_state()

def update_displayed_names(display_names) -> None:
    print (f'\n< Refresh displayed patient list >')  
    patient_listbox.config(state = tk.NORMAL)
    patient_listbox.delete(0, tk.END)
    for name in display_names:
        patient_listbox.insert(tk.END, name)
    patient_listbox.grid()
    if display_names:
        explanation_label.config(text="Select a patient, or drag new pdfs into the window")
        clear_list_button.config(state=tk.NORMAL, bg = 'white')
    else:
        explanation_label.config(text="Drag pdf files into the window")
        clear_list_button.config(state= tk.DISABLED, bg = 'light grey')
    explanation_label.grid()

if __name__ == "__main__":
    main(run_mode)


