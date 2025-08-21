import os
import PyPDF2
import shutil

class SearchERA:
    def __init__(self, dest_path, type_id):
        self.dest_path = dest_path
        self.type_id = type_id


    def place_era_file(self, file_path, era_path):
                                     
        shutil.copy(file_path, era_path)

    def search_era_file(self, patient_name, dispute_no_path, ICN, dispute_id, visible_rows):
        matching_files = []
        
        era_folder = os.path.join(self.dest_path, 'INST')
        for foldername, subfolders, filenames in os.walk(era_folder):
            for filename in filenames:
                if patient_name in filename:
                    if filename.endswith('ERA.pdf') or filename.endswith('era.pdf') or \
                    filename.endswith('ERA..pdf') or filename.endswith('ERA01.pdf'):
                        matching_files.append(filename)
                        for files in matching_files:
                            file_path = os.path.join(foldername, files)
                            if os.path.exists(file_path):
                                
                                if self.match_text(file_path, ICN): 
                                    if self.type_id == 'Institutional':
                                        era_path = os.path.join(dispute_no_path, dispute_id+' ERA.pdf')     
                                        self.place_era_file(file_path, era_path)

                                    elif self.type_id == 'Professional' and visible_rows > 1:
                                        era_folder = os.path.join(dispute_no_path, 'ERAS')
                                        if not os.path.exists(era_folder):
                                            os.makedirs(era_folder)

                                        era_path = os.path.join(era_folder)

                                        self.place_era_file(file_path, era_path)

                                    else:
                                        era_path = os.path.join(dispute_no_path, dispute_id+' ERA.pdf')     
                                        self.place_era_file(file_path, era_path)

    def match_text(self, file_path, ICN):
        
        with open (file_path, "rb") as textfile:
            reader = PyPDF2.PdfReader(textfile)
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text = page.extract_text()
                if ICN in text:
                    return True
            return False

