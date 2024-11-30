import PyPDF2
import os
import re
import time
from datetime import date, datetime, timedelta, timezone
import win32com
import hashlib
import shutil

from file_handling.handle_files import set_absolute_directory_path
from pdf.AA_lab import read_report as read_AA
from pdf.AA_lab import train_test_names as train_AA
from pdf.ML_lab import read_report as read_ML
from pdf.ML_lab import train_test_names as train_ML
from pdf.medas import read_medas_dump as read_medas

from classes import *


def get_pdfs_from_desktop(pdfs):
    
    try:
        # check inbox for patient test pdfs, if there are any then pass them to attempt_pdf_import module
        print()
        print("< Checking desktop for patient pdfs >")

        # list the paths of all PDF files on your desktop
        desktop_path = os.path.expanduser("~/Desktop")
        desktop_pdf_filepaths = [os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if f.endswith(".pdf")]
        if desktop_pdf_filepaths:
            print(f"found:\n{desktop_pdf_filepaths}")

        # iterate through those files
        for desktop_filepath in desktop_pdf_filepaths:

            original_filename = os.path.basename(desktop_filepath)

            # save a copy of the attachment in the temp folder
            temp_dir= set_absolute_directory_path('temp')
            temp_path = os.path.join(temp_dir, 'temp.pdf')
            shutil.copy(desktop_filepath, temp_path)

            #create a new instance of TIMC_pdf
            pdf = TIMC_pdf(temp_path, original_filename, None, None, None)
            print("checking: "+ original_filename)
            hash_value = assign_pdf_hash(pdf)

            #skip if the attachment already exists in the temp folder
            if hash_value and hash_value in pdfs:
                print (">> this pdf already exists in the directory so not imported")
            
            else:
                # this pdf doesn't exist so attempt to populate the instance 
                pdf = attempt_pdf_import(pdf)

                # assign a filename
                pdf.filepath = assign_hashed_pdf_filename(pdf, hash_value)

                # add to pdf dictionary, if valid (succesfully populated with name and type)
                if pdf.patient and pdf.type and pdf.filepath:
                    pdfs[hash_value]=pdf

            # delete the temp file
            delete_temp_file(temp_path)

        print ("< Finished checking Inbox for lab reports >")
        return pdfs
    
    except Exception as e:
        print (f'< Problem occured while checking inbox for lab reports: {e}/n')
        return pdfs

def get_pdfs_from_inbox(pdfs):

    try:
        # check inbox for patient test pdfs, if there are any then pass them to attempt_pdf_import module
        print("< Checking Microsoft Outlook inbox for recent lab reports (within 30 days) >")

        # Create an Outlook application object
        outlook = win32com.client.Dispatch("Outlook.Application")
        # Get the namespace (which represents the current user's session)
        namespace = outlook.GetNamespace("MAPI")
        # Open the default inbox folder
        inbox = namespace.GetDefaultFolder(6)  # 6 represents the inbox folder

        # Calculate the date from three months ago
        three_months_ago = datetime.now(timezone.utc) - timedelta(days=90)
        one_year_ago= datetime.now(timezone.utc) - timedelta(days=365)

        # Get all items in the inbox
        messages = inbox.Items

        # Filter messages based on ReceivedTime
        recent_messages = [message for message in messages if message.ReceivedTime >= three_months_ago]

        if not recent_messages:
            print ("There are no messages from within the last 90 days.")
            print ("< Checking Microsoft Outlook inbox for older lab reports (within 365 days) >")
            recent_messages = [message for message in messages if message.ReceivedTime >= one_year_ago]

        if not recent_messages:
            print ("There are no recent messages in inbox")
                
        else:
            number_of_emails=len(recent_messages)
            print("There are ",number_of_emails," recent messages in inbox. Checking each for attachments that might be a lab report")
        
            # Iterate through items in the inbox
            for message in recent_messages:

                # Check if the item has attachments
                if message.Attachments.Count > 0:

                    # check if the attachments are likely to be test reports?
                    # ML and AA labs report all have the words "TEST REPORT" in the title. This is how we will identify a lab report.
                    if "TEST" in message.subject.upper():
                    
                        # Iterate through attachments, checking if each is a pdf file
                        for i in range(len(message.Attachments)):
                            attachment=message.Attachments[i]
                            original_filename = attachment.FileName
                            if original_filename.endswith(".pdf"):
                                try:
                                    # save a copy of the attachment in the temp folder
                                    temp_dir= set_absolute_directory_path('temp')
                                    temp_path = os.path.join(temp_dir, 'temp.pdf')
                                    attachment.SaveAsFile(temp_path)

                                    #create an empty instance of TIMC_pdf
                                    pdf = TIMC_pdf(temp_path, original_filename, None, None, None)
                                    print("checking: "+ original_filename)
                                    hash_value = assign_pdf_hash(pdf)

                                    #skip if the attachment already exists in the temp folder
                                    if hash_value and hash_value in pdfs:
                                        print (">> this pdf already exists in the directory so not imported")

                                    else:
                                        # this pdf doesn't exist so attempt to populate the instance 
                                        pdf = attempt_pdf_import(pdf)

                                        # assign a filename
                                        pdf.filepath = assign_hashed_pdf_filename(pdf, hash_value)

                                        # add to pdf dictionary, if valid (succesfully populated with name and type)
                                        if pdf.patient and pdf.type and pdf.filepath:
                                            pdfs[hash_value]=pdf

                                except Exception as e:
                                    print (f"Problem attempting to import pdf file from Outlook: {original_filename}, {e}")

                                finally:
                                     # delete the temp file
                                     delete_temp_file(temp_path)

                            
                           
        print ("< Finished checking inbox for lab reports >")
        return pdfs
    
    except Exception as e:
        print (f'< Problem occured while checking inbox for lab reports: {e}/n')
        return pdfs

def delete_temp_file(temp_path):
    # tidy up by deleting the temporaty file, if it still exists
    if os.path.exists(temp_path):
        try:
            os.remove(temp_path)
        except:
            print ("error trying to delete " + temp_path)
            print ("I'll wait 1 second and try again...")
            time.sleep(1)  # Sleep for 1 second
            try:
                os.remove(temp_path)
            except Exception as e:
                print ("Nope that didn't work. Error: ",e)  

def assign_pdf_hash(pdf):
    # attempt to compute a hash value for the pdf file specified by temp_filepath
    try:
        algorithm="sha256"
        hasher = hashlib.new(algorithm)
        with open(pdf.filepath, "rb") as file:
            while chunk := file.read(8192):
                hasher.update(chunk)
        hash_value = hasher.hexdigest()
        return hash_value
    except:
        print (f'">> unable to create a hash value for {pdf.filepath}')    
        return None

def attempt_pdf_import(pdf):
    # try to read the contents of the file:
    try:
        pdf.content = get_pdf_content(pdf.filepath)
        if not pdf.content:
            print (">> this file does not contain readable text.")
            pdf.type = 'other'
            return pdf
    except Exception as e:
        print (f'">> Unable to read the content of {pdf.filepath}, {e}')
        return pdf
    
    # try to identify and name the file based on the content
    try:
        pdf = identify_pdf_type_and_patient(pdf)
    except:
        print (">> error attempting to identify the pdf file ", pdf.filepath)
        return pdf
    
    if pdf.type == None:
        print (">> Could not identify type of file")
        return pdf
    
    # handle the case where no name is returned
    if pdf.patient == None:
        print (">> Could not identify patient to assign this file to.")
        return pdf
    
    # Handle the case where only one name has been returned
    if len(pdf.patient.split()) == 1:
        pdf.patient += " Unknown"

    # we have now established that the pdf file is valid
    print(f'pdf identified as type: {pdf.type} for {pdf.patient}')
    return pdf

def assign_hashed_pdf_filename(pdf,hash_value):

    try:
        hashed_pdf_filename = str(hash_value)+".pdf"
    except Exception as e:
        print (f'>> Unable to rename pdf file with hash: {str(e)}')
        return None

    # assign filepath:
    try:
        temp_dir= set_absolute_directory_path('temp')
        new_filepath = os.path.join(temp_dir, hashed_pdf_filename)
    except Exception as e:
        print (f'>> Unable to assign file path to hash.pdf: {str(e)}')
        return None

    # try renaming the temp file with its new unique hash name
    try:
        os.rename(pdf.filepath, new_filepath)
        pdf.filepath = new_filepath
        print(f'>> pdf imported as {pdf.filepath}')
        return pdf.filepath
      
    except FileNotFoundError:
        print(f">> ERROR: File not found.")
        return None

    except FileExistsError:
        print(f">> ERROR: File already exists in the directory")
        return None
            
    except Exception as e:
        print(f">> An error occurred: {str(e)}")
        return None
   
def get_pdf_content(file_path):

    # attempt to open the pdf file specified in file_path
    # read and return the content, or None if unable

    try:
        pdfFileObj = open(file_path, 'rb')
        pdfReader = PyPDF2.PdfReader(pdfFileObj)
        numpages=len(pdfReader.pages)
        lines=[]
        for page in range(numpages):
            pageObj = pdfReader.pages[page]
            text=(pageObj.extract_text())
            page_lines=text.splitlines()
            lines.extend(page_lines)
        pdfFileObj.close()    
        return(lines)
    except Exception as e:
        pdfFileObj.close()    
        print ("Unable to read pdf: ",e)
        return None
    
def identify_pdf_type_and_patient(pdf):
    
    # looks through the pdf file content for terms which identify it.
    # - it may be a medas dump, a ML lab report, or an Al Arabi lab report, or maybe a report, or maybe something else
    # finds patient first name, surname, and the type of report.    

    try:
        # determine the pdf type. Use strings of text that only appear in a particular file type.  
        for line in pdf.content:
            if "TIMC: " in line:
                pdf.type = "TIMC document"
                break
            elif "Authorized on :" in line:
                pdf.type="ML lab report"
                break
            elif "This electronic copy of your tests result has been finalized by the laboratory director" in line or "AL Arabi Laboratory" in line:
                pdf.type="Al arabi lab report"
                break
            elif "VISIT NOTES" in line:
                pdf.type="Medas dump"
                break
            else:
                pdf.type="other"

        if not pdf.type:
            return pdf

        # seperate the pdf lines into a list of lines and words
        print (f'pdf type: {pdf.type}')
        lines, words = [], []
        for line in pdf.content:
            lines.append(line.strip())
            linewords = line.strip().split()
            for word in linewords:
                words.append(word)

        # Now we know that the file is valid: find the patient's first name and surname
        if pdf.type=="Medas dump":
            for line in range(len(pdf.content)):
                if pdf.content[line] =="Patient's Name":
                    nextline=pdf.content[line+1]
                    words=re.findall(r'\w+',nextline)
                    first_name = words[0].capitalize()
                    surname = words[-1].capitalize()
                    pdf.patient = f'{first_name} {surname}'

        if pdf.type=="ML lab report":
            try:
                for i in range(len(lines)):
                    if lines[i] == "Name":
                        name = lines[i+2]
                        # handles the case where the name spans 2 lines
                        if 'referred' in lines[i+4].lower() or 'clinic' in lines[i+4].lower():
                            name += f' {lines[i+3]}'
                        name = name.strip().split()
                        first_name = name[0].capitalize()
                        surname = name[-1].capitalize()
                        pdf.patient = f'{first_name} {surname}'
            except Exception as e:
                print (f"problem getting name from ML report {e}")

                    
        if pdf.type=="Al arabi lab report":        
            for line in pdf.content:
                words=line.split()
                if "Visit  No" in line:
                    surname = words[-1].capitalize()
                    if "Gender" in words:
                        idx = words.index("Gender")
                        first_name = words[idx+1].capitalize()
                        pdf.patient = f'{first_name} {surname}'
                    
        
        if pdf.type=="TIMC document":        
            for line in pdf.content:
                if "TIMC: " in line:
                    words = line.split()
                    pdf.patient = f'{words[1]} {words[-1]}'
                
        if pdf.type=="other":
            for i in range(len(words)):
                word = re.sub(r'[^a-zA-Z0-9]', '', words[i]) # just keep alphanumeric
                if not pdf.patient:
                    if word.lower() in ['name','patientname']:
                        try:
                            first_name = re.sub(r'[^a-zA-Z0-9]', '', words[i+1]).capitalize()
                            surname = re.sub(r'[^a-zA-Z0-9]', '', words[i+2]).capitalize()
                            pdf.patient = f'{first_name} {surname}'
                            break  
                        except:
                            pass

        return pdf
       
    except Exception as e:
        print("error identifyig patient and type from PDF content",e)
        return None
    
def identify_lab(lines):
    #figure out which lab we are dealing with    
    lab=None
    # String to check for (convert to lowercase for case-insensitive search)
    search_string1_AA = "AL Arabi Laboratory".lower()
    search_string2_AA ="This electronic copy of your tests result has been finalized by the laboratory director".lower()
    search_string_ML = "Authorized on :".lower()
    # Check if the search stringS appear in any of the strings in the list (after converting to lowercase)
    found_AA = False  
    if any(search_string1_AA in string.lower() for string in lines) or any (search_string2_AA in string.lower() for string in lines):
        found_AA = True
    found_ML = any(search_string_ML in string.lower() for string in lines)
    if found_AA:
        lab = "Al Arabi"
    elif found_ML:
        lab = "ML" 
    return(lab)

def read_data_from_pdfs(patient,run_mode):

    # This command handles reading data from medas and lab files and updating the Excel sheet 

    # initialise the dictionary
    data={}

    # identify list of medas and lab files in temp directory matching FIRST NAME and SURNAME
    temp_dir=set_absolute_directory_path('temp')
    print (patient.medas_attachments)
    print (patient.lab_attachments)
    for medas_file in patient.medas_attachments:
        print ("< extracting data from medas document >")
        data.update(extract_data_from_medas_dump(medas_file))

    for lab_report in patient.lab_attachments:
        print(f"< extracting data from {lab_report} >")
        data.update(extract_data_from_lab_report(lab_report,run_mode))

    for other_pdf in patient.other_attachments:
        data.update(extract_data_from_other_file(other_pdf))

    #reconcile Medas and Lab data if necessary, tidy up
    data=choose_best_data(data)
    return(data)

def extract_data_from_medas_dump(input_filename):
    
    medas_data={}

    #absolute path of the input file
    temp_dir=set_absolute_directory_path('temp')
    input_filepath = os.path.join(temp_dir, input_filename)

    # open and read pdf 
    pdfFileObj = open(input_filepath, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    numpages=len(pdfReader.pages)
    pagetext=""
    for page in range(numpages):
        pageObj = pdfReader.pages[page]
        text=(pageObj.extract_text())
        pagetext+=text

    lines=pagetext.splitlines(True)
    medas_data.update(read_medas(lines))

    return(medas_data)

def extract_data_from_other_file(input_filename):

    def read_other_file(lines):
        otherDict = {} 
        
        # check if QID contained in text   
        for i in range(len(lines)):
            line = lines[i].strip()
            words = line.split()
            if "QID" in words or "Qatar ID":
                try:
                    qid = words[words.index("QID")+2]
                    otherDict["QID"] = qid
                except:
                    pass

        # if not explicitly found, check for another 11 digit number       
        if "QID" not in otherDict:
            for i in range(len(lines)):
                line = lines[i].strip()
                words = line.split()
                for word in words:
                    if word.isdigit() and len(word) == 11:
                        otherDict["QID"] = word

        return(otherDict)
    
    #absolute path of the input file
    temp_dir=set_absolute_directory_path('temp')
    input_filepath = os.path.join(temp_dir, input_filename)

    # open and read pdf 
    pdfFileObj = open(input_filepath, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    numpages=len(pdfReader.pages)
    pagetext=""
    for page in range(numpages):
        pageObj = pdfReader.pages[page]
        text=(pageObj.extract_text())
        pagetext+=text

    lines=pagetext.splitlines(True)
    other_data = read_other_file(lines)
    print (other_data)
    return(other_data)

def extract_data_from_lab_report(input_filename,run_mode):

    #absolute path of the input file
    temp_dir=set_absolute_directory_path('temp')

    input_filepath = os.path.join(temp_dir, input_filename)

    # open and read pdf, generate a list of lines 
    pdfFileObj = open(input_filepath, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    numpages=len(pdfReader.pages)
    lines=[]
    for page in range(numpages):
        pageObj = pdfReader.pages[page]
        text=(pageObj.extract_text())
        page_lines=text.splitlines()
        lines.extend(page_lines)
    pdfFileObj.close()

    #figure out which lab the report is from
    lab = identify_lab(lines)

    if run_mode=='developer':
    #run the training sub to check if any new blood testdicthave been identified
        if lab == "ML":
            train_ML(lines)
        elif lab == "Al Arabi":
            train_AA(lines)
    
    #define the terms to search for:
    data={}
    if lab=="ML":
        data.update(read_ML(lines))
    elif lab=="Al Arabi":
        data.update(read_AA(lines))
 
    return(data)

def add_attachment_pdfs_to_main_pdf(filepath, patient, doctype):
    
    if doctype == "send by whatsapp":
        pdf_filepaths=[]
    else:
        pdf_filepaths=[filepath]

    for attachment in patient.lab_attachments:
        pdf_filepaths.append(attachment)

    for attachment in patient.other_attachments:
        pdf_filepaths.append(attachment)
    
    # Create a PDF object to write the merged PDF
    pdf_writer = PyPDF2.PdfWriter()

    # Open each existing PDF file
    for input_filepath in pdf_filepaths:
    
        with open(input_filepath, 'rb') as input_pdf_file:
            pdf_reader = PyPDF2.PdfReader(input_pdf_file)

            # Add all pages from the input PDF to the writer
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                pdf_writer.add_page(page)

    # Save the password-protected PDF to a file
    with open(filepath, 'wb') as output_pdf_file:
        pdf_writer.write(output_pdf_file)
    return

def encrypt_pdf_with_PyPDF2(filepath, password):

    # Open the existing PDF file
    with open(filepath, 'rb') as input_pdf_file:
        pdf_reader = PyPDF2.PdfReader(input_pdf_file)

        # Create a PDF object to write the password-protected PDF
        pdf_writer = PyPDF2.PdfWriter()

        # Add all pages from the input PDF to the writer
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            pdf_writer.add_page(page)

        # Encrypt the PDF with the password
        pdf_writer.encrypt(password)

    # Save the password-protected PDF to a file
    with open(filepath, 'wb') as output_pdf_file:
        pdf_writer.write(output_pdf_file)

    return

def choose_best_data(data):

    # this tidies the data dictionary, by looking for special cases where the data may not be optimally formatted

    print("tidying the data")

    # check QID has correct format, delete it if not
    if 'QID' in data:
        qid = data['QID'].strip()
        valid_flag=True
        if len(qid)!=11:
            print('Invalid QID {qid}, it does not have 11 digits')
            valid_flag=False
    
        if valid_flag==True:
            for char in qid:
                valid_flag = False if (char<"0" or char>"9") else valid_flag
                if not valid_flag:
                    print ('Invalid QID {qid}, not all characters are digits')

        if valid_flag==True:
            data['QID']=qid
            print (f'QID: "{data["QID"]}"')

        elif valid_flag==False:
            del data['QID']
      
    # try converting age to an integer, and assiging title based on sex
    if "Age/Sex" in data:
        if data["Age/Sex"]!="":
            data["Age"]=(data["Age/Sex"].split(" ")[0]).strip()
            data["Sex"]=(data["Age/Sex"].split("/")[-1]).strip()

    if "Title" not in data:
        data['Title']=""

    if "Age" in data:
        try:
            data["Age"]=int(data["Age"])
        except:
            print("Error: Age string cannot be converted to an integer")
    if "Sex" in data:
        if data["Sex"].capitalize()=="Male":
            data["Title"]="Mr"
            data["Pronoun"]="he"
        elif data["Sex"].capitalize()=="Female":
            data["Title"]="Ms"
            data["Pronoun"]="she"
    else:
        data["Title"]=""    

    # ensure that a doctor is assigned
    if "Doctor" not in data:
        if "Referred Doctor" in data:
            data["Doctor"]=data["Referred Doctor"]
        elif "Physician" in data:
            data["Doctor"]=data["Physician"]
        else:
            data["Doctor"]="Unknown"

    # tidy the doctor's name
    doctor = data['Doctor'].lower()
    if "nick" in doctor or "berry" in doctor or "nicholas" in doctor:
        data['Doctor']="Dr Nicholas Berry"
    elif "suzy" in doctor or "duckworth" in doctor: 
        data['Doctor']="Dr Suzy Duckworth"
    elif "lubna" in doctor or "saghir" in doctor: 
        data['Doctor']="Dr Lubna Saghir"
    elif "muna" in doctor or "farooqi" in doctor: 
        data['Doctor']="Dr Muna Farooqi"
    elif "julie" in doctor or "oh" in doctor: 
        data['Doctor']="Dr Julie Oh"

    if "Date of birth" not in data:
        data["Date of birth"]=""

    # populate clinic ref and strip any whitespace
    if "Clinic reference" not in data:
        data['Clinic reference']=""
        if "Clinic File No." in data:
            data["Clinic reference"]= data["Clinic File No."]
    data["Clinic reference"]=data["Clinic reference"].strip()
    
    # format phone number in form +97433788063
    if "Phone" not in data and "Contact No." in data:
        data['Phone']=data['Contact No.']

    if "Phone" in data:
        try:
            raw_phone = str(data['Phone'])
            digits = re.findall(r'\d', raw_phone)
            phone = ''.join(digits)
            if len(phone) ==8:
                phone = '974' + phone
            phone = '+'+ phone
            data['Phone']=phone
        except Exception as e:
            print ("Problem formatting phone number {e}")
    

   # format dates nicely

    if "Authorized on" in data:
        data['Date of report']=data['Authorized on']

    if "Date of report" in data:
        try:
            date_str=data['Date of report']
            # Convert the date string to a datetime object
            date_obj = datetime.strptime(date_str, "%d/%m/%Y")
            # Format the datetime object to the desired format
            formatted_date = date_obj.strftime("%dth %B %Y")
            data["Date of report"]=formatted_date
        except Exception as e:
            print("error generating date of report string",e)

    if "Date of report" not in data:
        try:
            today=date.today()
            data['Date of report']=today.strftime("%dth %B %Y")
        except Exception as e:
            print("error generating date of report string",e)

    # adjust for FOB which can only be positive or negative and must be lower case
    if "Occult blood, stool" in data:
        result=data["Occult blood, stool"][0]
        formatted_result=(result.lower()).strip()
        data["Occult blood, stool"][0]=formatted_result

    # format glucose level convert glu mg/dL to mmoml/l if necessary
    if "Fasting Glucose" in data:
        if data["Fasting Glucose"][1]=="mg/dL":
            try:
                value = data["Fasting Glucose"][0]
                if isinstance(value, str):
                    try:
                        value = float(value)
                    except ValueError:
                        print("The glucose value is a string but cannot be converted to an float.")
                data["Fasting Glucose"][0]=round(value/18.0182,1)
                data["Fasting Glucose"][1]="mmol/mL"
                data["Fasting Glucose"][2]=3.9
                data["Fasting Glucose"][3]=5.6
                
            except Exception as e:
                print("something went wrong converting glucouse from mg/dl to mmol/ml"&e)    

    if "Glucose-G (Random)" in data:
        if data["Glucose-G (Random)"][1]=="mg/dL":
            try:
                value = data["Glucose-G (Random)"][0]
                if isinstance(value, str):
                    try:
                        value = float(value)
                    except ValueError:
                        print("The glucose value is a string but cannot be converted to an float.")
                data["Glucose-G (Random)"][0]=round(value/18.0182,1)
                data["Glucose-G (Random)"][1]="mmol/mL"
                data["Glucose-G (Random)"][2]=3.9
                data["Glucose-G (Random)"][3]=5.6
                
            except Exception as e:
                print("something went wrong converting glucouse from mg/dl to mmol/ml"&e)    
            

    # patch upper and lower range values for female FBC values, if necessary
    if 'Sex' in data:
        if data['Sex'].lower() == "female":
            if 'Haemoglobin' in data:
                data['Haemoglobin'][2]=12.0,
                data['Haemoglobin'][3]=16.0
            if 'HCT' in data:
                data['HCT'][2] = 33
                data['HCT'][3] = 51
            if 'RBC count' in data:
                data['RBC count'][2]=3.8
                data['RBC count'][3]=5.2
            if 'MCH' in data:
                data['MCH'][2] = 26
                data['MCH'][3] = 34

    # patch PSA upper range for age values, if necessary
    if 'Prostate Specific Antigen (PSA Total)' in data and 'Age' in data:
        if data['Age'] <=50:
            data['Prostate Specific Antigen (PSA Total)'][3]=2.5
        elif data['Age'] <=60:
            data['Prostate Specific Antigen (PSA Total)'][3]=3.5
        else:
            data['Prostate Specific Antigen (PSA Total)'][3]=4.5

    # Reconcile Vitamin D as 25 Hydroxy (OH) Vitamin D, serum
    if 'Vitamin - D (25-Hydroxyvitamin D)' in data:
        data['25 Hydroxy (OH) Vitamin D, serum']=data['Vitamin - D (25-Hydroxyvitamin D)']

    if 'Vitamin D' in data:
        data['25 Hydroxy (OH) Vitamin D, serum']=data['Vitamin D']

    # Reconcile Vitamin B12 as B12
    if 'Vitamin B12' in data:
        data['B12']= data['Vitamin B12']

    if 'Magnesium,' in data:
        data['Magnesium'] = data['Magnesium,']

    if "Phosphorous," in data:
        data['Phosphorus'] = data['Phosphorous,']

    if "D-DIMER" in data:
        data['D-dimer']= data['D-DIMER']

    # format the date of report for filename
    today = date.today()
    data['Date of report for filename']=today.strftime("%Y%m%d")
    return(data)

