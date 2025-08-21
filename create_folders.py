import os
import re
import shutil
import string
from place_data import PlaceData

class CreateFolders:
    def __init__(self, destination):
        self.dispute_ids_path = None
        self.destination = destination

    def create_dispute_id_folders(self):
        self.dispute_ids_path = os.path.join(self.destination, 'Dispute IDs')
        if not os.path.exists(self.dispute_ids_path):
            os.makedirs(self.dispute_ids_path)
        return self.dispute_ids_path


    def extract_number(self, dispute_id):
        if dispute_id != None or dispute_id != "":
            dispute_number = re.findall(r'\d+', dispute_id)
            return dispute_number[0].strip() if dispute_number else None

    def create_dispute_no_folder(self, dispute_id):
        if dispute_id != None or dispute_id != "":
            dispute_number = self.extract_number(dispute_id)
            dispute_no_path = os.path.join(self.dispute_ids_path, dispute_number)
            if not os.path.exists(dispute_no_path):
                os.makedirs(dispute_no_path)
            return dispute_no_path  


    def place_dispute_file(self, dispute_id):
        disputes_folder = os.path.join(self.destination, 'Disputes')
        if os.path.exists(disputes_folder):
            dispute_number = re.findall(r'\d+', dispute_id)
            pdf_files = [f for f in os.listdir(disputes_folder) if f.endswith('.pdf')]
            for file in pdf_files:
                number = re.findall(r'\d+', file)
                if number and dispute_number and number[0] == dispute_number[0]:
                    source_dispute_file = os.path.join(disputes_folder, file)
                    destination_dispute = os.path.join(self.dispute_ids_path, dispute_number[0])
                    shutil.copy(source_dispute_file, destination_dispute)

    def create_subfolders(self, dispute_no_path, parts):
        subfolders = [f for f in os.listdir(dispute_no_path) 
              if os.path.isdir(os.path.join(dispute_no_path, f))]
        
        subfolders.sort()
        subfolder_name = os.path.basename(dispute_no_path)
        all_present = all(name in folder_name for name in parts)

        for idx, folder_name in enumerate(subfolders):
            if all_present:
                letter = string.ascii_uppercase[idx]
                old_path = os.path.join(dispute_no_path, folder_name)
                new_name = f"{subfolder_name}-{letter}"
                new_path = os.path.join(dispute_no_path, new_name)
                os.rename(old_path, new_path)


  