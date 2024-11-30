#blood test ranges
import os
import PyPDF2
import fitz  # PyMuPDF


input_filepath = os.path.abspath('data/laboratory-reference-ranges.pdf')

ignore_strings = ['Laboratory Tests', 'Reference Ranges','ABIM Laboratory Test Reference Ranges  ̶  January 2024','Revised - January 2024']

doc = fitz.open(input_filepath)
for page in doc:
    text = page.get_text()
    lines = text.split('\n')
    i=0
    group = None
    test = None
    test_value = None
    group = None
    groupcount = 0
    printflag = True
    while True:
        try:
            line = lines[i]
            if line in ignore_strings:
                i += 1

            elif line.strip().isdigit():
                i+=1

            elif not group:
                    
                if lines[i+1][0:1]==" ":                    # new group detected
                    group = line.strip()
                    test = lines[i+1].strip()
                    test_value = lines[i+2].strip()
                    i+=3

                else:                                       # new test, not grouped
                    test = line.strip()
                    test_value = lines[i+1].strip()
                    i+= 2

            elif group:                                     #existing group and not the first test
                
                if line[0:1]==" ":                           #another member of the group
                    test = line.strip()
                    test_value = lines[i+1].strip()
                    i+= 2
                    if lines[i+2][0:1] != " ":
                        group = None

                else:                                       # a new test not in the group               
                    group = None
                    groupcount = 0
                    printflag= False

            
            
            if group:
                if groupcount == 0:
                    output = (f'{group}\n    {test}: {test_value}')
                else:
                    output = (f'    {test}: {test_value}')
                    groupcount += 1
            else:
                output =  (f'{test}: {test_value}')

            if printflag:
                print(output)
            printflag=True          
        



        except Exception as e:
            print (e)
            break
    input ('press a key')
    i=0


'''
pdfFileObj = open(input_filepath, 'rb')
pdfReader = PyPDF2.PdfReader(pdfFileObj)
numpages=len(pdfReader.pages)
numpages = 1
pagetext=""
for page in range(numpages):
    pageObj = pdfReader.pages[page]
    text=(pageObj.extract_text())
    pagetext+=text

lines=pagetext.splitlines(True)
for line in lines:
    print(line)
    input ("new (g)roup), new (t)est within the group")
    '''