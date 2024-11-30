import os
import re
import json
from file_handling.handle_files import set_absolute_directory_path


## AA lab report uses the data structure:
#  data{test:[searchterm, anchorterm, index of value rel to anchorterm, index of units rel to anchor term, lower bound, upper bound, +/-[ignoreterms] ]}

def read_report(lines):

    #Constants
    DATA_DIR=set_absolute_directory_path('data')
    SEARCH_TERMS=os.path.join(DATA_DIR,"search_terms_AA.json")

    #initialise the data dictionary
    AA_data={}

    #import list of search terms
    with open(SEARCH_TERMS,'r') as json_file:    # these are the terms to search for which will tell us if we have found data
        search_terms=json.load(json_file)
    json_file.close()

    # iterate through lines
    for line_item in range(len(lines)):
        raw_line = lines[line_item]
        line = raw_line.strip()
        words = (line.split())
        words = [word.strip(',') for word in words]

        # check if patient data is contained: age, doctor, QID
        keyword='Year(s)'
        if keyword in words:
            # suggests a line like this: ['Muna', 'Iqbal', 'FarooqiMale52', 'Year(s)', '10', 'Month(s)', '13', 'Day(s)'] for biochem,
            # or like this [['44', 'Year(s)', '8', 'Month(s)', '9', 'Day(s)', 'Patient', 'Name']] for haem
            try:
                # find the index position of the keyword in the string
                idx=words.index(keyword)

                if idx == 1: # = haematology
                    # find the age string
                    AA_data['Age']= words[0]
                    print ('Found Age:', AA_data['Age'])
                    # find the doctor_sex string
                    for nextline_item in range(line_item, line_item+7):
                        raw_nextline = lines[nextline_item]
                        nextline= raw_nextline.strip()
                        nextwords= (nextline.split())
                        nextwords = [nextword.strip(',') for nextword in nextwords]
                        for nextword in nextwords:
                            # Use regular expressions to try to split the string into (1) Doctor, (2) Sex
                            match = re.match(r'([A-Za-z]+)(Male|Female)', nextword)
                            if match:
                                AA_data['Doctor'] = match.group(1)
                                AA_data['Sex'] = match.group(2)
                                print("Found Doctor: ", AA_data['Doctor'])
                                print("Found Sex: ", AA_data['Sex'])
                                break
                        
                else: # = not haematology
                    # find the doctor_sex_age string
                    doctor_sex_age=words[idx-1]
                    # Use regular expressions to tyr to split the string into (1) Doctor, (2) Sex, and (3) Age
                    match = re.match(r'([A-Za-z]+)(Male|Female)(\d+)', doctor_sex_age)           
                    # does it match?
                    if match:
                        # add to patient data
                        doc_surname = match.group(1)
                        doc_surname = doc_surname.lower()
                        AA_data['Doctor'] = match.group(1)
                        AA_data['Sex'] = match.group(2)
                        AA_data['Age'] = match.group(3)
                        print("Found Doctor: ", AA_data['Doctor'])
                        print("Found Sex: ", AA_data['Sex'])
                        print("Found Age: ", AA_data['Age'])                       
                    else:
                        # something went wrong
                        print("I tried to extract doctor, age and sex, but the string does not match the expected format.")
                        print("I found:", words)
            except:
                #something went wrong
                print("something went wrong trying to read doctor, age and sex from AA lab report")

        # check if sample date is contained
        if "Sample Date Consulant" in raw_line:
            try:
                AA_data["Date of test"]=words[0]
            except:
                print("problem reading date of test")


        # check if QID is contained
        if "QID" in raw_line:
            try:
                QID_identified=False
                # look for ['QID','27903600171'] ... if biochem
                # look for ['27903600171'],['QID'] ... if haematology
                QIDwords = words
                QIDwords.append(lines[line_item-1])
                for word in QIDwords:
                    pattern = re.compile(r'^\d{11}$')
                    if pattern.match(word):
                        AA_data['QID']=word
                        QID_identified = True
                        break
                if QID_identified:
                    print (f'Found QID: {word}')
                else:
                    print ("Unable to identify QID")
                    print ("the line was:", words)
            except:
                print("problem reading QID")

        # check if a test name searchstring is contained in the line:
        for term in search_terms:
            searchstring = search_terms[term][0]
            ignorestrings = None
            ignoreflag = False

            # check if this particular test name contains any words that we should look out for to ignore the result
            # they will be appended as a list at the 7th value in the dictionary entry
            if len(search_terms[term])==7:
                ignorestrings = search_terms[term][6]
            
            # is the searchstring in the line?
            if searchstring in raw_line:
                # are there ignore strings?
                if ignorestrings:
                    for ignorestring in ignorestrings:
                        #is there an ignorestring in this line?
                        if ignorestring in raw_line:
                            ignoreflag=True


                if not ignoreflag:
                # build a list of the words of the line before, the line, and the line after
                    try:
                        threelines = lines[line_item-1]+" , "+lines[line_item]+" , "+lines[line_item+1]
                        words = (threelines.split())
                        words = [word.strip(',') for word in words]
                        try:
                            keyword=search_terms[term][1]
                            # check that the keyword is actually in the list (it may not be, eg keyword MCH is not in tne MCHC list):
                            if keyword in words:
                                # read the data from the list
                                keyword_index= words.index(keyword)
                                valueRI=search_terms[term][2]
                                unitsRI=search_terms[term][3]
                                #get the values
                                test_value = words[keyword_index+valueRI]
                                units=words[keyword_index+unitsRI]
                                lower=search_terms[term][4]
                                upper=search_terms[term][5]
                                #write the data
                                try:
                                    AA_data[term]=[test_value,units,lower,upper]
                                    print(f"Found {term}: {test_value} {units} ({lower}, {upper})")
                                except:
                                    print ("ERROR adding",term,test_value,units,lower,upper)
                        except Exception as e:
                            #don't crash if there's a problem
                            print("problem retrieving data about",term,": ",e)
                            print("the line I read was",words)
                    except Exception as e:
                        print('Error trying to build a list of words to scan for {searchstring}: {e}')
                                    
    return(AA_data)

def train_test_names(lines):

    # this is just for Al Arabi labs
    # scan the imported lines for numeric values.
    # Each time a numeric value is found, check if its line contains an Al_Arabi search term
    # if not, ask if it should be added 
    # 
    # Constants
    DATA_DIR=set_absolute_directory_path('data')
    TEST_NAMES=os.path.join(DATA_DIR,"test_names.txt")
    SEARCH_TERMS=os.path.join(DATA_DIR,"search_terms_AA.json")
    IGNORE_TERMS=os.path.join(DATA_DIR,"ignore_terms.json")

    #import list of test names
    with open(TEST_NAMES, 'r') as file:
        test_names = [line.strip() for line in file.readlines()]
        file.close()

    # import dictionary of terms to identify information to ignore
    with open(IGNORE_TERMS,'r') as json_file:
        ignore_terms = json.load(json_file)
    json_file.close()

    # import dictionary of terms to indentify infornation to include
    # term = {test_name:[searchstring, keyword, valueRelIdx, unitsRelIdx, LowerRedIdx, UpperRelIdx]}
    with open(SEARCH_TERMS,'r') as json_file:
        search_terms = json.load(json_file)
    json_file.close()

    # iterate w through each line in the pdf file
    for line_item in range(len(lines)):
        raw_line= lines[line_item]
        line = lines[line_item].strip()
        words = (line.split())
        words = [word.strip(',') for word in words]  #words is now a list of words appearing in the line, with no commas

        #check if we ought to ignore this line:
        ignore_flag=False
        for ignore_item in ignore_terms:
            if ignore_item in words and ignore_terms[ignore_item]=="Al Arabi":
                # ignore if the term is in the list of ignore terms
                ignore_flag=True

        #check if there are alphabet characters in the line:
        letters_in_line=False
        for char in "abcdefghijklmnopqrstuvwxyz":
            if char in line:
                letters_in_line = True
        if not letters_in_line:    
            # ignore if no alphabet characters in the line
            ignore_flag = True
        
        #check if there is a numeric value in the line:
        number_found=False
        for item in words:
            try:
                number = float(item)
                number_found=True
                break
            except:
                number_found=False
        if number_found==False:
            # ignore if no number in the list
            ignore_flag=True

        #check this line?
        if ignore_flag!=True:
            #check if the line contains a known search term in the Al Arabi terms file
            recognised_flag=False
            for term in search_terms:  
                searchstring=search_terms[term][0]
                if searchstring in line:
                    recognised_flag=True

            if recognised_flag == False:
            #Ther is no recognised search term in the line:
                print ("The line ", words," does not seem to contain a recognised term.")
            #check if the line is a result which should be added to search terms
                ans=input ("Does this line contain a test result which should be remembered?")
                if ans.upper() == "Y":

                    #check if there is an obvious match iny the main test names file?
                    test_name=None
                    for i in range(len(test_names)):
                        str1 = test_names[i].lower()
                        str2 = line.lower()
                        if str1 in str2:
                            print(str1, "is in", str2)
                            ans=input("I think "+line+" should map to "+ test_names[i] +" ,is this correct?")
                            if ans.upper()=="Y":
                                test_name=test_names[i]
                                break
                                
                    # if not, then seek manual entry
                    if test_name==None:
                        print ("Here is the list of known tests:")
                        for i in range(len(test_names)):
                            print (i,test_names[i])
                        print ("The line I have found is: ",line)
                        valid_input=False
                        while valid_input==False:
                            ans= input ("Would you like to associate the test with one of these? Enter the number to link with, or N to remember a new test").strip()
                            ans=ans.upper()
                            if ans=="N":
                                valid_input=True
                            else:
                                try:
                                    ans=int(ans)
                                    valid_input = True
                                except:
                                    pass
                            
                        if ans!="N":
                            #associate this new term with a known test:
                            test_name=test_names[ans]
                            print ("OK. I will associate values from ",words," with the known term",test_name)

                        else:
                            # otherwise, ask for the new test's name:
                            print ("Let's add a new blood test to the list then.")
                            test_name = input ("What would you like the name to appear as in the report? ")

                    #now ask for the associated search term, value index, units index, lower and upper bounds
                    
                    # build a list of the words of the line before, the line, and the line after
                    try:
                        threelines = lines[line_item-1]+" , "+lines[line_item]+" , "+lines[line_item+1]
                    except:
                        print("Error trying to put 3 lines together")
                        threelines=lines[line_item]
                    words = (threelines.split())
                    words = [word.strip(',') for word in words]

                    # now an inout loop to get data from human
                    finished_flag=False
                    while finished_flag==False:
                        print ("Here is a list of words:")
                        for i in range(len(words)):
                            print (i,words[i])

                        # check wich term to use 
                        print("Which searchword would you like to associate with the test "+test_name+" ?")
                        valid_input=False
                        while valid_input==False:
                            ans= input ("enter a number from the list above, or enter an exact searchstring (case sensitive)")
                            try:
                                ans=int(ans)
                                # this means we have been given a number
                                searchstring=words[ans]
                            except:
                                searchstring=ans
                            valid_input=searchstring in threelines
                            if valid_input==False:
                                print("I'm not seeing that searchterm in the raw line. Here is the raw line: ",raw_line)
                        


                        # and which term should be the anchor: all other terms and values will be measured relative to this 
                        valid_input=False
                        ans = input ("And which is the key/anchor word? Enter a number from the list above.")
                        while valid_input==False:
                            try:
                                keyword_index=int(ans)
                                keyword=words[keyword_index]
                                valid_input=True
                            except:
                                print("invalid keyword")
                            
                        # get the rest of the values and terms, relative to the keyword
                        print("OK. Here is the line again",line)
                        print("Here is the list of words in the line, relative to "+keyword)
                        print("For the following questions, enter a number to choose the word or value, or hit [return] to set as None/blank")
                        for i in range(len(words)):
                            print (i-keyword_index,words[i])
                        valid_input=False
                        while valid_input==False:
                            try:
                                valueRI= int(input ("which word gives the value?"))    
                                unitsRI= int(input ("And which word gives the units?"))
                                lower= float(input ("What is the lower bound of normal? (enter the actual value)"))
                                upper =float(input ("What is the upper bound of normal?"))
                                valid_input=True
                            except:
                                print ("At least one of those answers was not valid")
                        
                        #we can now update the docs of Al Arabi search terms, and the file of test names
                        search_terms.update({test_name:[searchstring,keyword,valueRI,unitsRI,lower,upper]})
                    
                        print("I have remembered that when I see ["+searchstring+"] in an Al Arabi lab report, I should update the data field ["+test_name+"] like this: "+test_name+", units "+words[keyword_index+unitsRI]+", range "+str(lower)+" - "+str(upper))                 
                        
                        ans=input("Is this correct? y/n").upper()
                        if ans=="Y":
                            finished_flag=True    

                    sorted_json = json.dumps(search_terms, indent=4, sort_keys=True)
                    try:
                        with open(SEARCH_TERMS,'w') as file:
                            file.write(sorted_json)
                    except Exception as e:
                            print("ERROR SAVING SEARCH TERMS AA:",str(e))
                    #with open(SEARCH_TERMS,'w') as json_file:
                    #    json.dump(search_terms,json_file,indent=4)
                    #json_file.close()

                elif ans.upper()!="Y":
                    # I have been told not to remember this unknown search term
                    # check if I should actively ignore it in future
                    ans=input("Should I ignore lines like this in future?")
                    if ans.upper()=="Y":
                        term = ""
                        while term == "" or term not in words:
                            term = input("Please enter the exact term you would like to ignore:")
                            if term not in words:
                                print ("that term is not in the line")
                        print ("OK. From now on, I will ignore any lines in Al Arabi reports containing the word ",term,".")
                        ignore_terms.update({term:"Al Arabi"})     

                        with open(IGNORE_TERMS,'w') as json_file:
                            json.dump(ignore_terms,json_file,indent=4)
                        json_file.close()

    #reconcile files if needed
    for term in search_terms:
        if term not in test_names:
            print(term+" is not in test_names.txt so I will update")             
            with open(TEST_NAMES,'a') as file:
                file.write(str(term) + '\n') 
            file.close()


    return
