from rename_data import RenameData
from excel_data_handler import ExcelDataHandler
from create_folders import CreateFolders
from search_era import SearchERA
import os
import string
import shutil

class RepeatedDisputes:
    def __init__(self, repeated_disputes, excel_path, destination):
        self.repeated_disputes = repeated_disputes
        self.excel_path = excel_path
        self.destination = destination
        excel_new_path = self.excel_path

    def find_disputes_excel(self):
        excel_new_path = self.excel_path
        excel_data_handler = ExcelDataHandler(self.excel_path)
        create_folders = CreateFolders(self.destination)
        
        dispute_ids_path = create_folders.create_dispute_id_folders()
        df = excel_data_handler.load_workbook(excel_new_path)
        suffixes = [f"-{char}" for char in string.ascii_uppercase]

        for dispute_id in self.repeated_disputes:
            if df is not None:
                matched_rows = df[df['Dispute No.'] == dispute_id]
                repeated_dispute_no_path = os.path.join(dispute_ids_path, 
                                                        create_folders.extract_number(dispute_id))
                suffix_counter = 0
                suffix = ""
                print(repeated_dispute_no_path)
                for index, row in matched_rows.iterrows():
                    type_id = row['Type']
                    ICN = row['ICN']
                    service_date = row['Service Date']
                                       
                    patient_name = row['Patient Name']
                    parts = self.split_patient_names(patient_name)
                    for disputes in os.listdir(repeated_dispute_no_path):
                        if all(part.upper() in disputes.upper() for part in parts):
                            if suffix_counter < len(suffixes):
                                suffix = suffixes[suffix_counter]
                            else:
                                suffix = f"-X{suffix_counter}"

                            original_path = os.path.join(repeated_dispute_no_path, disputes)
                            new_folder_name = create_folders.extract_number(dispute_id) + suffix
                            new_path = os.path.join(repeated_dispute_no_path, new_folder_name)
                            
                            if not os.path.exists(new_path):
                                os.rename(original_path, new_path)
                                rename_data = RenameData(repeated_dispute_no_path, service_date)
                                cleaned_service_date = rename_data.process_service_date()
                                # if not os.path.exists(os.path.join(repeated_dispute_no_path, 'NURSE & PHYSICIAN NOTES')):
                                #     os.makedirs(os.path.join(repeated_dispute_no_path, 'NURSE & PHYSICIAN NOTES'))
                                for file in os.listdir(new_path):
                                    file_path = os.path.join(new_path, file)
                                    print (file_path)
                                    print (new_folder_name)
                                    rename_data.pdf_rename(file, new_folder_name, file_path, cleaned_service_date)
                                    
                            elif os.path.exists(original_path) and os.path.isdir(original_path):
                                shutil.rmtree(original_path)

                    suffix = suffixes[suffix_counter]
                    era_folder_path = os.path.join(repeated_dispute_no_path, 'ERAS')
                    self.rename_eras(parts, era_folder_path, suffix, dispute_id, ICN, type_id)             
                    suffix_counter += 1              
                                
    def split_patient_names(self, patient_name):
        parts = patient_name.strip().split()
        count = len(parts)

        first = parts[0] if count >=1 else ""
        last = parts[1] if count >=2 else ""
        middles = parts[1:-1]

        if count == 4:
            middle1, middle2 = middles
            return first, middle1, middle2, last
        elif count == 3:
            middle1 = middles[0]
            return first, middle1, last
        elif count == 2:
            return first, last

    def rename_eras(self, parts, era_folder_path, suffix, dispute_id, ICN, type_id):
        create_folders = CreateFolders(self.destination)
        search_era = SearchERA(self.destination, type_id)
        if os.path.exists(era_folder_path):
            for file_name in os.listdir(era_folder_path):
                file_path = os.path.join(era_folder_path, file_name)
                if all(part.upper() in file_name.upper() for part in parts):
                    if search_era.match_text(file_path, ICN):
                        dispute_number = create_folders.extract_number(dispute_id)
                        new_file_name = 'DISP-' + dispute_number + suffix + ' ERA.pdf'
                        new_path = os.path.join(era_folder_path, new_file_name)
                    
                        if not os.path.exists(new_path):
                            os.rename(os.path.join(era_folder_path, file_name), new_path)
                        elif os.path.exists(os.path.join(era_folder_path, file_name)):
                            os.remove(os.path.join(era_folder_path, file_name))

    
        