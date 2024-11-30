# CLASSES
class TIMC_user:
    def __init__(self, title, firstname, surname, role, stamp, banner):
        self.title = title
        self.firstname = firstname
        self.surname = surname
        self.role = role
        self.stamp = stamp
        self.banner = banner

class TIMC_document:
    def __init__(self, type, title, author, password, stamp, patient, date, description):
        self.type = type
        self.title = title
        self.author = author
        self.password = password                
        self.stamp = stamp                     
        self.patient = patient
        self.date = date
        self.description = description

class TIMC_patient:
    def __init__(self, title, firstname, surname, display_name, sex, medas_attachments, lab_attachments, other_attachments, QID, phone, email):
        self.title = title
        self.firstname = firstname
        self.surname = surname
        self.display_name = display_name
        self.sex = sex
        self.medas_attachments = medas_attachments          # list of filepaths of pdfs
        self.lab_attachments = lab_attachments              # list of filepaths of pdfs
        self.other_attachments = other_attachments          # list of filepaths of pdfs
        self.QID = QID
        self.phone = phone
        self.email = email

class TIMC_pdf:
    def __init__(self, filepath, original_filename, type, patient, content):
        self.filepath = filepath
        self.original_filename = original_filename
        self.type = type
        self.patient = patient
        self.content = content
        