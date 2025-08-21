from csv import excel
import tkinter as tk
from excel_data_handler import ExcelDataHandler
import os
from place_holder_entry import PlaceholderEntry
from tkinter import filedialog
from create_folders import CreateFolders
from pop_up_messages import UserInterface
from rename_data import RenameData
from search_era import SearchERA
from rename_word_files import RenameWordFiles
from place_data import PlaceData
from rename_data import RenameData
from repeated_disputes import RepeatedDisputes

class DataArrangement:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Collection Software")
        self.root.state('zoomed')
        self.df = None
        self.dispute_no_path = None
        self.visible_rows = None
        self.parts = None
        self.repeated_dispute = []
        

        self.make_folders_var = tk.BooleanVar()
        self.filter_excel_files_var = tk.BooleanVar()
        self.search_for_ERAS_var = tk.BooleanVar()
        self.rename_word_files_var = tk.BooleanVar()
        self.place_data_var = tk.BooleanVar()
        self.rename_data_var = tk.BooleanVar()

        tk.Label(root, text="Excel File:").grid(row=0, column=0, sticky='w', padx = 10)
        self.excel_entry = PlaceholderEntry(root, "Select Excel File (.xlsx)", width=50)
        self.excel_entry.grid(row=0, column=1, padx=20, pady=5)
        tk.Button(root, text="Browse", command=self.browse_excel).grid(row=0, column=2)

        tk.Label(root, text="Source Folder:").grid(row=1, column=0, sticky='w', padx = 10)
        self.source_entry = PlaceholderEntry(root, "Select Source Folder", width=50)
        self.source_entry.grid(row=1, column=1, padx=20, pady=5)
        tk.Button(root, text="Browse", command=self.browse_source).grid(row=1, column=2)

        tk.Label(root, text="Destination Folder:").grid(row=2, column=0, sticky='w', padx = 10)
        self.dest_entry = PlaceholderEntry(root, "Select Destination Folder", width=50)
        self.dest_entry.grid(row=2, column=1, padx=20, pady=5)
        tk.Button(root, text="Browse", command=self.browse_dest).grid(row=2, column=2)

        tk.Label(root, text="Template Folder"). grid(row=3, column=0, sticky='w', padx=10)
        self.template_entry = PlaceholderEntry(root, "Select Template Path", width = 50)
        self.template_entry.grid(row=3, column=1, padx=20, pady=5)
        tk.Button(root, text="Browse", command=self.browse_template).grid(row=3, column=2)

        
        make_folders_cb = tk.Checkbutton(root, text="Make Folders", variable=self.make_folders_var)
        make_folders_cb.grid(row=4, column=0, padx= 10, pady=5, sticky = 'w')


        filter_excel_file_cb = tk.Checkbutton(root, text="Filter Excel Files", 
                                              variable=self.filter_excel_files_var)
        filter_excel_file_cb.grid(row=5, column=0,padx = 10, pady=5, sticky = 'w')


        search_for_ERAS_cb = tk.Checkbutton(root, text="Search for ERAS",
                                            variable=self.search_for_ERAS_var)
        search_for_ERAS_cb.grid(row=6, column=0, padx=10, pady=5, sticky='w')
        

        rename_word_files_cb = tk.Checkbutton(root, text="Rename Word Files",
                                              variable=self.rename_word_files_var)
        rename_word_files_cb.grid(row=7, column=0, padx=10, pady=5, sticky='w')

        place_data_cb = tk.Checkbutton(root, text="Place Data Folders",
                                              variable=self.place_data_var)
        place_data_cb.grid(row=8, column=0, padx=10, pady=5, sticky='w')

        rename_data_cb = tk.Checkbutton(root, text="Rename Data",
                                        variable = self.rename_data_var)
        rename_data_cb.grid(row=9, column=0, padx=10, pady=5, sticky='w')


        self.process_button = tk.Button(root, text="Process Files", command=self.process_files)
        self.process_button.grid(row=10, column=0, padx = 10, pady=5)

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, path)

    def browse_source(self):
        path = filedialog.askdirectory()
        if path:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, path)   

    def browse_dest(self):
        path = filedialog.askdirectory()
        if path:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, path)

    def browse_template(self):
        path = filedialog.askdirectory()
        if path:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, path)



    def on_make_folders_toggle(self, dispute_id, dest_path):
        if self.make_folders_var.get():
            create_folders = CreateFolders(dest_path)
            create_folders.create_dispute_id_folders()
            self.dispute_no_path = create_folders.create_dispute_no_folder(dispute_id)
            create_folders.place_dispute_file(dispute_id)
            

    def on_filter_excel_files(self, excel_data_handler, dispute_id):
        if self.filter_excel_files_var.get():
            self.visible_rows = excel_data_handler.filter_excel_files(dispute_id, self.dispute_no_path)
            
    def on_search_for_ERAS(self, patient_name, dest_path, ICN, dispute_id, type_id):
        
        if self.search_for_ERAS_var.get():
            search_era = SearchERA(dest_path, type_id)
            search_era.search_era_file(patient_name, self.dispute_no_path, ICN,dispute_id
                                       , self.visible_rows)
            

    def on_rename_word_files(self, type_id, facility, dispute_id, template_path):
        if self.rename_word_files_var.get():
            
            rename_word_files = RenameWordFiles(type_id, facility, template_path)
            rename_word_files.get_file()
            rename_word_files.change_header(dispute_id, self.dispute_no_path)

    def on_place_data_files(self, source_path, patient_name, service_date, 
                            dest_path, facility, claim_no, dispute_id):
        if self.place_data_var.get():
            
            place_data = PlaceData(source_path, patient_name, service_date, dest_path, facility, 
                                   self.dispute_no_path, claim_no, dispute_id)
            place_data.search_folders(facility, patient_name, service_date)
            self.parts = place_data.split_patient_names

    def on_rename_data(self, service_date):
        if self.rename_data_var.get():
           rename_data_files = RenameData(self.dispute_no_path, service_date)
           rename_data_files.rename_data_files()

    def process_files(self):
        excel_path = self.excel_entry.get().strip()
        source_path = self.source_entry.get().strip()
        dest_path = self.dest_entry.get().strip()
        template_path = self.template_entry.get().strip()

        if any(field in ("", "Select", None) for field in [excel_path, source_path, dest_path, template_path]):
            text = "Please Fill in all Fields"
            user_interface = UserInterface(text)
            user_interface.create_message_info_pop_up()
            

        elif all(os.path.exists(path) for path in [excel_path, source_path, dest_path, template_path]):
            
            excel_data_handler = ExcelDataHandler(excel_path)
            
            df = excel_data_handler.load_workbook(excel_path)
           
            if df is not None:

                for index, row in df.iterrows():
                    type_id = row['Type']
                    patient_name = row['Patient Name']
                    service_date = row['Service Date']
                    dispute_id = row['Dispute No.']
                    claim_no = row['CLAIM NO']
                    ICN = row['ICN']
                    facility = row['Facility']

                    if dispute_id != None or dispute_id != "":

                        if type_id == 'Institutional':
                        
                            self.on_make_folders_toggle(dispute_id, dest_path)
                            self.on_filter_excel_files(excel_data_handler, dispute_id)
                            self.on_search_for_ERAS( patient_name, dest_path, ICN, dispute_id, type_id)
                            self.on_rename_word_files(type_id, facility, dispute_id, template_path)
                            self.on_place_data_files(source_path, patient_name, service_date, 
                                                     dest_path, facility, claim_no, dispute_id)
                            self.on_rename_data(service_date)

                        elif type_id == 'Professional':
                            self.visible_rows = 0
                            self.on_make_folders_toggle(dispute_id, dest_path)
                            self.on_filter_excel_files(excel_data_handler, dispute_id)
                            self.on_place_data_files(source_path, patient_name, service_date, dest_path, facility,
                                                     claim_no, dispute_id)
                            self.on_rename_word_files(type_id, facility, dispute_id, template_path)
                        
                            self.on_search_for_ERAS( patient_name, dest_path, ICN, dispute_id, type_id)
                            if self.visible_rows > 1:
                                if dispute_id not in self.repeated_dispute:
                                    self.repeated_dispute.append(dispute_id)

                            else:
                                self.on_rename_data(service_date)
        
        repeated_disputes = RepeatedDisputes(self.repeated_dispute, excel_path, dest_path)
        repeated_disputes.find_disputes_excel()
        print(self.repeated_dispute)
                        
if __name__ == "__main__":
    root = tk.Tk()
    app = DataArrangement(root)
    root.mainloop()
            