import os
import json
import re
from file_handling.handle_files import set_absolute_directory_path

def read_report(lines):

    # Constants for file paths
    DATA_DIR=set_absolute_directory_path('data')
    PATIENT_DATA_FIELD_NAMES=os.path.join(DATA_DIR,"patient_data_field_names_ml.json")
    SEARCH_TERMS=os.path.join(DATA_DIR,"search_terms_ML.json")
    IGNORE_TERMS=os.path.join(DATA_DIR,"ignore_terms.json")

    # initialise data
    ML_data={}
    with open(PATIENT_DATA_FIELD_NAMES,'r') as json_file:
        patient_data=json.load(json_file)
    json_file.close()
    with open(SEARCH_TERMS,'r') as json_file:
        search_terms=(json.load(json_file))
    json_file.close()
    with open(IGNORE_TERMS,'r') as json_file:
        raw_ignore_terms=(json.load(json_file))
        global_ignore_terms=[]
        for term in raw_ignore_terms:
            if raw_ignore_terms[term]=="ML labs":
                global_ignore_terms.append(term)
       
   
    # check if each line matches a demographic or blood test name, and if so, read the value of the next line   
    for i in range(len(lines)-3):
        #examine a line (word) and the next line (word)
        raw_line=lines[i]
        line = raw_line.strip()
        nextline= lines[i+1].strip()

        for test_name in search_terms:

            # check the list of terms we need to search for assosciated with this test

            query = search_terms[test_name][0]

            # if query returns a string, make it into a list of one item
            if type(query)==str:
                searchstring_list =[]
                searchstring_list.append(query)
            elif type(query)==list:
                searchstring_list = query
            else:
                print("error search string for {test_name} must be a string or a list")
                next

            #iterate through the list of searchstrings
            for searchstring in searchstring_list:    

                if searchstring in raw_line:
                    #check if we need to ignore it based on global ignore terms:
                    ignore_flag=False
                    for term in global_ignore_terms:
                        if term in raw_line:
                            print(f"ignore because {term} is in {raw_line} ")                                
                            ignore_flag=True

                    # check if this particular test name contains any words that we should look out for to ignore the result
                    # they will be appended as a list at the 5th value in the dictionary entry

                    if len(search_terms[test_name])==5:
                        ignorestrings = search_terms[test_name][4]
                    
                        if ignorestrings:
                            for ignorestring in ignorestrings:
                                #is there an ignorestring in this line?
                                if ignorestring in raw_line:
                                    print(f"ignore because {ignorestring} is in {raw_line} ")  
                                    ignore_flag=True

                    # if we should not ignore it:
                    if not ignore_flag:
                        try:
                            #get the values
                            test_value = nextline
                            units=search_terms[test_name][1]
                            lower=search_terms[test_name][2]
                            upper=search_terms[test_name][3]
                            
                        except:
                            #don't crash if there's a problem
                            print("problem retrieving ML data about",test_name)
                            print("the line I read was",raw_line)
                            test_name,test_value,units,lower,upper=None,None,None,None

                        #write the data
                        try:
                            ML_data[test_name]=[test_value,units,lower,upper]
                            print(f"added: {test_name}: {test_value} {units} ({lower} - {upper})")
                        except:
                            print ("ERROR adding",test_name,": ",test_value,units,lower,upper)
                        break
            

        if line in patient_data:
            #get patient field data
            ML_data[line]=lines[i+2]
        
    # check if there are special lines like these:
    for i in range(len(lines)-1):
        line = lines[i]
        try:
            if 'Authorized on' in line:
                ML_data["Authorized on"] = (line.split(" ")[-2]).replace("-","/")    
            if 'Qatar ID.' in line:
                ML_data["QID"]=lines[i+2]
            if 'Clinic File No.' in line:
                ML_data["Clinic File No."]= lines[i+2]
        except:
            print("Problem reading special data")    
    return(ML_data)

def train_test_names(lines):

    #Constants
    DATA_DIR=set_absolute_directory_path('data')
    SEARCH_TERMS=os.path.join(DATA_DIR,"search_terms_ML.json")
    IGNORE_TERMS=os.path.join(DATA_DIR,"ignore_terms.json")
    TEST_NAMES=os.path.join(DATA_DIR,"test_names.txt")

    # this is just for ML labs
    # scan the imported lines for numeric values.
    # Each time a numeric value is found, check if the precedeing term called query_term
    # is in test_names, patient_data_field_names, or not_test_names.
    # if it is not in any of these, ask if it should be added.

    with open(SEARCH_TERMS,'r') as json_file:
        search_terms = json.load(json_file)
    json_file.close()

    with open(IGNORE_TERMS,'r') as json_file:
        ignore_terms = json.load(json_file)
    json_file.close()

    #import list of test names
    with open(TEST_NAMES, 'r') as file:
        test_names = [line.strip() for line in file.readlines()]
        file.close()


    for i in range(len(lines)-4):
        line = lines[i].strip()
        try:
            number = float(line)
            numberFlag=True
        except:
            numberFlag=False

        if numberFlag==True:
            #it's a value:
            term = lines[i-1].strip()
            if term ==":":
                term = lines[i-2].strip()

            if term not in ignore_terms:

                #check if there is a search string already associated with this term
                known_search_term=False
                for test_name in search_terms:
                    query = search_terms[test_name][0]
                    # if query returns a string, make it into a list of one item
                    if type(query)==str:
                        searchstring_list =[]
                        searchstring_list.append(query)
                    elif type(query)==list:
                        searchstring_list = query

                    else:
                        print("error search string for {item} must be a string or a list")
                        next

                    for search_term in searchstring_list:
                        if search_term in term:
                            known_search_term=True
                            break    

                if known_search_term==False:
                    units=lines[i+1].strip()
                    refRange=lines[i+2].strip()
                    print ("It looks like I have encountered a new term:")
                    print(term,number,units,refRange)
                    ans=input("Do you want to remember this a blood test name (b), remember it as a term to ignore(i), or do nothing (n)?")
                    if ans.upper()=="B":
                        #figure out the lower and upper reference ranges by spliting the term
                        row=[0,units,refRange,0]
                        pattern = r'\d+\.\d+'
                        matches = re.findall(pattern, row[2])
                        numbers = [float(match) for match in matches]
                        if len(numbers)==2:
                            row[2]=numbers[0]
                            row[3]=numbers[1]
                        elif len(numbers)==1:
                            row[2]=0
                            row[3]=numbers[0]
                        else:
                            row[2]=0
                            row[3]=0
                        #update the dictionary of blood tests
                        search_terms[term]=([[term],row[1],row[2],row[3]])
                    elif ans.upper()=="I":
                        #update the dictiorary of terms to ignore
                        ignore_terms.update({term:"ML labs"})

    #save the lists
    sorted_json = json.dumps(search_terms, indent=4, sort_keys=True)
    
    try:
        with open(SEARCH_TERMS,'w') as file:
            file.write(sorted_json)
    except Exception as e:
        print("ERROR SAVING SEARCH TERMS ML:",str(e))

   
    try:
        with open(IGNORE_TERMS,'w') as json_file:
            json.dump(ignore_terms,json_file,indent=4)
        json_file.close()
    except Exception as e:
        print("ERROR SAVING IGNORE TERMS ML:",str(e))




    #reconcile files if needed
    for term in search_terms:
        if term not in test_names:
            print(term+" is not in test_names.txt so I will update")             
            with open(TEST_NAMES,'a') as file:
                file.write(str(term) + '\n') 
            file.close()



    return
