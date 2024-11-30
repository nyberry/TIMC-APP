import os
import json
import re
from pdf.format_date import format_date
from file_handling.handle_files import set_absolute_directory_path

def extract_number(s):
    try:
        # Use regular expression to find all numbers in the string
        numbers = re.findall(r'\d+\.\d+|\d+', s)
    
        # Join the list of numbers
        joined_number = ''.join(numbers)

        # Check if the joined number has a decimal point
        if '.' in joined_number:
            result = float(joined_number)
        else:
            result = int(joined_number)
        return result    
    except:
        print (f'problem extracting value from string {s}')
        return None

def read_medas_dump(lines):
   
    #constants
    DATA_DIR = set_absolute_directory_path('data')
    SEARCH_TERMS_FILE=os.path.join(DATA_DIR,"search_terms_medas.json")

    #import list of search terms
    with open(SEARCH_TERMS_FILE,'r') as json_file:    # these are the terms that Medas uses to organise the data. We will refer to them during search
        search_terms=json.load(json_file)
    json_file.close()

    # intialise a data dictionary
    medas_data = {}

    #define the terms to manually search for:
    title,full_name,first_name,surname,reg_no,sex="","","","","",""
    date_of_birth,date_of_medical="",""
    doctor=""
    BP_systolic,BP_diastolic="",""
    pulse,SpO2,height,weight,temperature="","","","",""
    history,examination="",""
    occupation=""
    smears=""
    pmh=""
    dh=""
    fh=""
    phone=""
    email=""
    allergies =""   

    # carry out the manual search  
    cursor = 0
    while cursor < len(lines)-2:
        line = lines[cursor].strip()
        nextline= lines[cursor+1].strip()

        if line =="Patient's Name":
            try:
                words=re.findall(r'\w+',nextline)
                full_name=words
                first_name = words[0]
                surname = words[-1]
                first_name = first_name.capitalize()
                surname = surname.capitalize()
            except:
                print (f'problem extracting patients name')

        if line =="Email Id":
            try:    
                email=nextline
            except:
                print (f'problem extracting email')

        if line == "Contact No":
            try:
                phone=nextline
            except:
                print (f'problem extracting phone')

        if line =="Age / Sex" or line=="Age/Sex":
            try:
                words=re.findall(r'\w+',nextline)
                age = words[0]
                sex = words[-1]
                if sex.upper() == "MALE":
                    title = "Mr"
                elif sex.upper() == "FEMALE":
                    title = "Mrs"
                else:
                    title = ""
            except:
                print (f'problem extracting age/ sex')

        if line =="Date of Birth":
            try:
                words=re.findall(r'\w+',nextline)
                words.remove(words[0])
                day,month,year = words[1],words[0],words[2]
                # in long form
                date_of_birth_long, date_of_birth_DDMMYY,date_of_birth_American= format_date(day, month, year)
            except:
                print(f'problem extracting date of birth')

        if line =="Doctor":
            try:
                doctor=nextline
            except:
                print("problem extractinf doctor")

        if line[0:10] =="Entered On":
            try:
                words=re.findall(r'\w+',line)
                words=words[3:]
                day,month,year = words[1],words[0],words[2]
                # in long form
                date_of_medical = format_date(day, month, year)[0]
            except:
                print ("problem extracting date of entry")

        if line=="Chief Complaint":
            try:
                history=""
                scanline=cursor  
                while scanline < len(lines)-2 and lines[scanline+1].strip() not in search_terms and lines[scanline+1].strip()[0:11]!="VITAL SIGNS":
                    scanline += 1
                    history += lines[scanline].strip()+" "
            except Exception as e:
                print (f"problem extracting history {e}")

        if line=="B.P (Systolic)":
            try:
                BP_systolic=extract_number(nextline)
            except:
                print ("Error reading systolic BP")

        if line=="B.P (Diastolic)":
            try:
                BP_diastolic=extract_number(nextline)
            except:
                print ("Error reading diastolic BP")

        if line=="Temperature":
            try:
                temperature=extract_number(nextline)
            except:
                print ("Error reading temperature")

        if line=="Pulse":
            try:
                pulse=extract_number(nextline)
            except:
                print ("Error reading pulse")

        if line=="O2 Saturation":
            try:
                SpO2=extract_number(nextline)
            except:
                print ("Error reading O2 sats")

        if line=="Height":
            try:
                height=extract_number(nextline)
            except:
                print ("Error reading height")

        if line=="Weight":
            try:
                weight=extract_number(nextline)
            except:
                print ("Error reading weight")

        if line=="EXAMINATION NOTES":
            try:
                scanline=cursor+1
                examination=""  
                while scanline < len(lines) and lines[scanline+1].strip() not in search_terms:
                    scanline+=1
                    examination+=lines[scanline].strip()+" "
            except:
                print ("Error reading exam findings")

        for word in ['mother', 'father', 'brother', 'sister', 'grandmother', 'grandfather']:
            if word in line.lower():
                try:
                    fh+=str(line)+'\n'
                except:
                    print ("error reading family history")
        
        if 'work:' in line.lower():
            try:
                occupation = line.replace('work:','')
            except:
                print ("error reading occupation")
        if "works as " in line.lower():
            try:
                occupation = line.replace('works as ','')
            except:
                print ("error reading occupation")

        if line=="Past Cervical Smears":
            try:
                smears=nextline
            except:
                print ("error reading smears")

        if "PMH: " in line:
            try:
                pmh+=(str(line).replace('PMH: ',''))+" "
            except:
                print ("error reading PMH")

        if "DH: " in line:
            try:
                dh+=str(line).replace('DH: ','')
                if 'nil' in line:
                    dh="nil"
            except:
                print ("error reading DH")
        
        if "FH: " in line:
            try:
                fh+=str(line).replace('FH: ','')
                if 'nil' in line:
                    fh="nil"
            except:
                print ("error reading FH")

        if line == "Drug Allergy":
            try:
                allergies=nextline
            except:
                print ("error reading allergies")
        
        
        cursor+=1

    '''
    # scan the medas doc: check if each line matches a search term, and if so, read the value of the next lines until another search term is encountered
    cursor = 0
    while cursor <len(lines):
        term = lines[cursor].strip()
        cursor += 1
        print(term)

        if term in search_terms and term not in medas_data:
            instruction = search_terms[term] # get the instruction about how to proceed
            value = ""
            continue_flag = True

            while continue_flag == True:
                a = input ("continue")
                if cursor == len(lines):
                    continue_flag = False
                else:
                    nextline = lines[cursor].strip()
                    if nextline in search_terms:
                        continue_flag = False
                    else:
                        # if we are told to read to the end of the next line only:
                        if instruction == "line" and value == "":
                            value = nextline
                        # if we are told to read several lines:
                        elif instruction == "lines":  
                            value += nextline if value == "" else '\n'+nextline
                        # if we are told to read one word:
                        elif instruction == "word" and value == "":
                            words = nextline.split()
                            value = "" if not words else words[0]
                        #finally:
                        cursor+=1
            # we now have the valut to assign to the term
            print(f'{term} : {value}')     
            medas_data[term]= value
    '''

    #write data to text file:

    medas_data.update({"Title" : title,
        "First name" : first_name,
        "Surname": surname,
        "Full name": full_name,
        "Clinic reference":reg_no,
        "Age":age,
        "Sex":sex,
        "Email":email,
        "Phone":phone,
        "Date of birth": date_of_birth,
        "Date of birth DDMMYY": date_of_birth_DDMMYY,
        "Date of medical": date_of_medical,
        "Physician": doctor,
        "Occupation": occupation,
        "Cervical smear":smears,
        "Systolic BP":BP_systolic,
        "Diastolic BP":BP_diastolic,
        "Temperature": temperature,
        "Resting heart rate":pulse,
        "Oxygen Saturation":SpO2,
        "Height":height,
        "Weight":weight,
        "History":history,
        "Examination":examination,
        "Family History":fh,
        "Past Medical History":pmh,
        "Cervical smears":smears,
        "Medications":dh,
        "Allergies":allergies})

    print ('Data extracted from Medas document:')    
    for key,value in list(medas_data.items()):  #list so can delete empty keys while iterating
        if value!="":
            print (f'{key}: {value}')
        else:
            try:
                del medas_data[key]
            except:
                print (f'problem deleting empty key {key}')
    print ()

    return(medas_data)
