import os, shutil
import json
import time
import requests
import pyautogui
from datetime import datetime
from ftplib import FTP

def delete_temp_files_and_close_temp_windows(temp_folder_window_titles):  

    # declare global variables:
    global hash_dict

    # Close temp files window if open
    for window in temp_folder_window_titles:
        try:
            pyautogui.getWindowsWithTitle(window)[0].close() 
        except Exception as e:
            print (f"unable to close {window}, {e}")

    # deletes all files and subfolders from the temp directory 
    directory_path = set_absolute_directory_path('temp')
    try:
        # Check if the directory exists
        if os.path.exists(directory_path) and os.path.isdir(directory_path):
            print("<Cleaning working files and subfolders from temp folder>")
            # Walk through all files and subdirectories in the temp directory
            for root, dirs, files in os.walk(directory_path, topdown=False):
                # Delete all files
                for file_name in files:
                    file_path = os.path.join(root, file_name)
                    try:
                        os.remove(file_path)
                        print("Deleted", file_path)
                    except Exception as e:
                        print("Unable to delete:", e)

                # Delete all subdirectories
                for dir_name in dirs:
                    dir_path = os.path.join(root, dir_name)
                    try:
                        shutil.rmtree(dir_path)
                        print("Deleted directory", dir_path)
                    except Exception as e:
                        print("Unable to delete directory:", e)

            print("All files and subfolders in temp folder have been deleted.")
                #show_status('','Temporary files and subfolders deleted')
        else:
            print("Directory does not exist or is not a directory:", directory_path)
    except Exception as e:
        print("An error occurred:", e)

        # forget all those files
        hash_dict={}


def set_absolute_directory_path(target_directory):
    # make sure that the current working directory is the same as the location of main.py, ie /..:
    os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)),".."))
    # return the absolute directory of the target folder:
    target_directory_absolute_path= os.path.abspath(os.path.join("../", target_directory))
        
    if not os.path.exists(target_directory_absolute_path):
        try:
            os.makedirs(target_directory_absolute_path)
            print (f'{target_directory_absolute_path} did not exist so has been created')
        except:
            print (f'Error: {target_directory_absolute_path} does not exist and cannot be created')

    return(target_directory_absolute_path)

def read_TIMCusers_file():
    # Constants for file paths
    DATA_DIR = set_absolute_directory_path('data')
    users_file = os.path.join(DATA_DIR,'TIMC_users.json')
    with open (users_file, 'r') as json_file:
        users = json.load(json_file)
    print ("users file read")
    return(users)

def write_TIMCusers_file(users):
    # Constants for file paths
    DATA_DIR = set_absolute_directory_path('data')
    users_file = os.path.join(DATA_DIR,'TIMC_users.json')
    with open (users_file, 'w') as json_file:
             json.dump(users,json_file,indent=4)
    print("users file updated")
    upload_to_server(['data/TIMC_users.json'])

def set_filepaths_and_filenames(data):
    # set the filenames and filepaths to use for the documents to be generated:

    # what type of report is this? Medical report, test report, or just document if unknown 
    try:
        type_of_document = "Document_" if data['medas_flag']==True else "Report_"
    except:
        type_of_document = "Document_"
    
    # get the date string Eg. 20240201
    try:
        date_of_report=str(data['Date of report for filename'])+"_"
    except:
        date_of_report=""

    # get the name string Eg. Joe Bloggs
    try:
        name_for_unencrypted_report = f'{data["Surname"]}_{data["First name"]}_'
        name_for_encrypted_report = f'{data["First name"]}{data["Surname"][0]}_'
    except:
        name_for_unencrypted_report = ""
        name_for_encrypted_report=""

    # get the clinic reference Eg. IM12345
    if 'Clinic reference' in data:
        if data['Clinic reference']=="":
            ref_for_report = "TIMC_"
        else:
            ref_for_report = f'{data["Clinic reference"]}_'
    else:
        ref_for_report = "TIMC_"
        
    # Name the Word docx.
    # The docx is for internal use and editing only. It should be easily identifiable. Eg:
    # Joe_Smith_IM12345_Test_Report_20240201.docx
    # it is not protected with a password 
    try:
        data['docx draft filename'] = f'{name_for_unencrypted_report}{ref_for_report}{type_of_document}{date_of_report}draft.docx'
        data['docx final filename'] = f'{name_for_unencrypted_report}{ref_for_report}{type_of_document}{date_of_report}unencrypted.docx'
    except:
        data['docx draft filename'] = "draft document.docx"
        data['docx final filename'] = "final document.docx"

    # Name the PDF documents.
        
    # The unencrptyed pdf is for attaching to the MEDAS file. It should be easily identifiable Eg:
    # Joe_Smith_IM12345_Test_Report_20240201.pdf
    # it is not protected with a password
    try:
        data['pdf final filename'] = f'{name_for_unencrypted_report}{ref_for_report}{type_of_document}{date_of_report}unencrypted.pdf'
    except:
        data['pdf final filename'] = "unencrpyted document.pdf"

    # The encrptyed pdf is for sending out of the organisation. It should NOT be easily identifiable Eg:
    # 20240201_IM23452_JoeS_encrypted.pdf if we know the clinic number, or
    # 20240201_TIMC_JoeS_encrypted.pdf if we do not know the clinic number
    # it is protected with a password, which is the 11 digit QID if we know it, or the DOB in form DDMMYY if we don't
    try:
        data['pdf final filename encrypted'] = f'{date_of_report}{ref_for_report}{name_for_encrypted_report}encrypted.pdf'
    except:
        data['pdf final filename encrypted'] = "Document_TIMC_encrypted.pdf"

    # the XLS filename:    
    data['xlsm filename'] = f'{name_for_unencrypted_report}{ref_for_report}Worksheet_{date_of_report}.xlsm' 
 

    # Check if report directories exist, create if not
    try:
        reports_dir = set_absolute_directory_path('..\\reports')
        if not os.path.exists(reports_dir):
            os.makedirs(reports_dir)
            print(f"Directory '{reports_dir}' created.")
        unencrypted_reports_dir=f'{reports_dir}\\unencrypted'
        if not os.path.exists(unencrypted_reports_dir):
            os.makedirs(unencrypted_reports_dir)
            print(f"Directory '{unencrypted_reports_dir}' created.")
        unencrypted_docx_dir=f'{unencrypted_reports_dir}\\docx'
        if not os.path.exists(unencrypted_docx_dir):
            os.makedirs(unencrypted_docx_dir)
            print(f"Directory '{unencrypted_docx_dir}' created.")
        unencrypted_pdf_dir=f'{unencrypted_reports_dir}\\pdf'
        if not os.path.exists(unencrypted_pdf_dir):
            os.makedirs(unencrypted_pdf_dir)  
            print(f"Directory '{unencrypted_pdf_dir}' created.")
        encrypted_pdf_dir=reports_dir
    except Exception as e:
        reports_dir = set_absolute_directory_path('..')
        unencrypted_reports_dir = reports_dir
        unencrypted_docx_dir = reports_dir
        unencrypted_pdf_dir = reports_dir
        encrypted_pdf_dir = reports_dir
        print("problem setting report directories. Files will be saved to {reports_dir}.", e)

    # make the filepaths
    try:
        data['docx draft filepath'] =  os.path.join(unencrypted_docx_dir,data['docx draft filename'])
        data['docx final filepath'] =  os.path.join(unencrypted_docx_dir,data['docx final filename'])
        data['pdf final filepath'] =  os.path.join(unencrypted_pdf_dir,data['pdf final filename'])
        data['pdf final filepath encrypted'] = os.path.join(encrypted_pdf_dir,data['pdf final filename encrypted'])
        data['xlsm filepath'] =  os.path.join(reports_dir,data['xlsm filename'])
    except Exception as e:    
        data['docx draft filepath'] =  data['docx draft filename']
        data['docx final filepath'] =  data['docx final filename']
        data['pdf final filepath'] =  data['pdf final filename']
        data['pdf final filepath encrypted'] = data['pdf final filename encrypted']
        data['xlsm filepath'] =  data['xlsm filename']
        print("problem setting report filepaths. Files named without path and will save in working directory.", e)
    
    return(data)

def decode_timestamp(timestamp):
    time_struct = time.localtime(timestamp)
    formatted_time = time.strftime("(%d %b %H:%M)", time_struct)
    return formatted_time

def upload_to_server(files):

    #FTP server and credentials
    ftp_host = 'ftp.clini.co.uk'  
    ftp_username = 'clini.co.uk'
    ftp_password = 'clini66!'
 
    # folders
    local_base_folder= set_absolute_directory_path('')
    remote_base_folder= 'https://clini.co.uk/TIMCapp'
    ftp_base_folder = '/HTDOCS/TIMCapp'

    # upload remote files    

    try:            
        with FTP(ftp_host) as ftp:
            ftp.login(ftp_username, ftp_password)

            # load remote timestamp log
            response = requests.get(f'{remote_base_folder}/timestamps.json')
            remote_timestamps = response.json() if response else {}

            # upload the files
            for file in files:
                with open(f'{local_base_folder}/{file}', "rb") as f:
                    directory, filename = os.path.split(file)
                    ftp.cwd(f'{ftp_base_folder}/{directory}')
                    ftp.storbinary(f"STOR {filename}", f)
                print(f'Uploaded {file}')
                remote_timestamps[file]=int(time.time())

            # update remote timestamp file, via a temporary file as cannot json write direct to FTP
            json_data = json.dumps(remote_timestamps)
            with open(f'{local_base_folder}/temp/timestamps.json', "w") as temp_file:
                temp_file.write(json_data)
            ftp.cwd(ftp_base_folder)
            with open(f'{local_base_folder}/temp/timestamps.json',"rb") as temp_file:
                ftp.storbinary('STOR timestamps.json', temp_file)

    except Exception as e:
        print (e)

def check_if_templates_need_updating(run_mode):

    #FTP server and credentials
    ftp_host = 'ftp.clini.co.uk'  
    ftp_username = 'clini.co.uk'
    ftp_password = 'clini66!'
 
    # folders
    local_base_folder= set_absolute_directory_path('')
    remote_base_folder= 'https://clini.co.uk/TIMCapp'
    ftp_base_folder = '/HTDOCS/TIMCapp'
    folders = ['data','images/stamps', 'images/banners', 'images/loading']

    # create local timestamp log
    local_timestamps= {}
    try:
        for folder in folders:
            for filename in os.listdir(f'{local_base_folder}/{folder}'):
                file = f'{folder}/{filename}'
                modification_time = int(os.path.getmtime(f'{local_base_folder}/{file}'))    # time.timezone()??
                local_timestamps[file] = modification_time
    except Exception as e:
        print(e)
                
    # load remote timestamp log
    response = requests.get(f'{remote_base_folder}/timestamps.json')
    remote_timestamps = response.json() if response else {}
    
    # if needed, created a new remote timestamp log
    if not response or run_mode=='update':
        try:
            with FTP(ftp_host) as ftp:
                ftp.login(ftp_username, ftp_password)
                for folder in folders:
                    ftp.cwd(f'{ftp_base_folder}/{folder}')
                    for filename in ftp.nlst():
                        file = f'{folder}/{filename}'
                        modification_time = ftp.sendcmd(f"MDTM {ftp_base_folder}/{file}")
                        # Parse modification time response, it returns a different format to what we need
                        response_parts = modification_time.split()
                        if len(response_parts) == 2:
                            timestamp = time.strptime(response_parts[1], "%Y%m%d%H%M%S")
                            remote_timestamps[file] = int(time.mktime(timestamp))
        except Exception as e:
            print (e)

    # build dict of most recent versions of the files: remote or local
    try:
        print ("Checking files to see which version is most recent:")
        most_recent={}
        for file, remote_timestamp in remote_timestamps.items():
            if file not in local_timestamps:
                most_recent[file] = 'does not exist locally'
        for file, local_timestamp in local_timestamps.items():
            remote_timestamp = remote_timestamps.get(file)
            if not remote_timestamp:
                most_recent[file] = 'does not exist remotely'
            elif local_timestamp > remote_timestamp:
                print (f'{file}: most recent version is local at {decode_timestamp(local_timestamp)}')
                most_recent[file] = 'local'
            elif remote_timestamp > local_timestamp:
                print (f'{file}: most recent version is remote at {decode_timestamp(remote_timestamp)}')
                #print (f'local timestamp is {decode_timestamp(local_timestamp)}\n')
                most_recent[file] = 'remote'
            else:
                most_recent[file] = 'equal'
    except Exception as e:
        print (e)

    # devloper mode: update remote files    
    if run_mode == 'developer' or run_mode == 'update':
        try:            
            with FTP(ftp_host) as ftp:
                ftp.login(ftp_username, ftp_password)

                # delete any files which don't exist locally
                for file, location in most_recent.items():
                    if location == "does not exist locally":
                        ans = input (f'File {file} {location}. Do you want to delete it? (y/n)').lower()
                        if ans == 'y':
                             directory, filename = os.path.split(file)
                             ftp.cwd(f'{ftp_base_folder}/{directory}')
                             try:
                                del remote_timestamps[file]
                                ftp.delete(filename)
                                print (f'{filename} deleted.')
                             except Exception as e:
                                print (f'{e}')

                # update any remote files
                for file, location in most_recent.items():
                    if location == 'local' or location == 'does not exist remotely':
                        with open(f'{local_base_folder}/{file}', "rb") as f:
                            directory, filename = os.path.split(file)
                            try:  # to change working directory to the remote folder
                                ftp.cwd(f'{ftp_base_folder}/{directory}')
                            except: # If directory doesn't exist, create it
                                ftp.mkd(f'{ftp_base_folder}/{directory}')
                                ftp.cwd(f'{ftp_base_folder}/{directory}')
                            ftp.storbinary(f"STOR {filename}", f)
                        print(f'Uploaded {file}')
                        remote_timestamps[file]=int(time.time())

                # update remote timestamp file, via a temporary file as cannot json write direct to FTP
                json_data = json.dumps(remote_timestamps)
                with open(f'{local_base_folder}/temp/timestamps.json', "w") as temp_file:
                    temp_file.write(json_data)
                ftp.cwd(ftp_base_folder)
                with open(f'{local_base_folder}/temp/timestamps.json',"rb") as temp_file:
                    ftp.storbinary('STOR timestamps.json', temp_file)
        except Exception as e:
            print (e)

    # user mode: update local files    
    if run_mode == 'user':
        updated = 0
        try:
            for file, location in most_recent.items():
                if location == 'remote' or location == 'does not exist locally':
                    response = requests.get(f'{remote_base_folder}/{file}')
                    if response.status_code == 200:  #success
                        with open(f'{local_base_folder}/{file}', "wb") as f:
                            f.write(response.content)
                        print (f"Downloaded {file}")
                        updated += 1
                        local_timestamps[file]=int(time.time())
                    else:
                        print (response.status_code)
        except Exception as e:
            print(e)
        if updated:
            print (f'updated {updated} files')
        else:
            print (f'all files are up to date')
            

if __name__ == "__main__":
    check_if_templates_need_updating('developer')