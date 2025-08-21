from datetime import datetime
import os
import PyPDF2
import fitz 

class RenameData:
    def __init__(self, dispute_no_path, service_date):
        self.dispute_no_path = dispute_no_path
        self.service_date = service_date

    def find_folder(self):
        return [
            os.path.join(self.dispute_no_path, name)
            for name in os.listdir(self.dispute_no_path)
            if os.path.isdir(os.path.join(self.dispute_no_path , name))]

    def rename_files(self, new_file_name, file_path):
        if os.path.exists(file_path):
            if not os.path.exists(os.path.join(self.dispute_no_path, new_file_name)):
                os.rename(file_path, os.path.join(self.dispute_no_path, new_file_name))

    def read_file(self, file, pattern1, pattern2, file_path, folder_name):
        with open (file_path, "rb") as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text = page.extract_text()
                if pattern1 in text:
                    pdf_file.close()
                    self.rename_files(f'DISP-{folder_name} NURSE AND PHYSICIAN NOTES.pdf', file_path)
                    return
                elif pattern2 in text:
                    pdf_file.close()
                    self.rename_files(f'DISP-{folder_name} NURSE AND PHYSICIAN NOTES.pdf', file_path)
                    return

    def process_service_date(self):
        if self.service_date:
            return self.service_date.strftime('%m%d%Y')
        return None

    def save_new_pdf(self, file_path, page_number, folder_name):
        
        old_pdf_document = fitz.open(file_path)
        new_pdf_document = fitz.open()
        
        
        for new_pdf_document_page_number in range(page_number):
            new_pdf_document.insert_pdf(old_pdf_document, from_page = new_pdf_document_page_number, to_page = new_pdf_document_page_number)
        
        count = new_pdf_document.page_count

        if count > 0:
        
            new_file_path = os.path.join(self.dispute_no_path, 'DISP-' + folder_name + ' NURSE AND PHYSICANS NOTES.pdf')

            new_pdf_document.save(new_file_path)
            new_pdf_document.close()
            old_pdf_document.close()
            
            os.remove(file_path)
            return new_file_path

    def detect_image(self, file_path, folder_name):
        
        pdf_document = fitz.open(file_path)
        for page_number in range(pdf_document.page_count):
            page = pdf_document.load_page(page_number)
            
            if page.get_images(full = True):
                pdf_document.close()
                new_file_path = self.save_new_pdf( file_path, page_number, folder_name)
                return new_file_path
        return 
                 

    def rename_data_files(self):
        
        folder_data = self.find_folder()
        for folder in folder_data:
                      
            for file in os.listdir(folder):
                
                file_path = os.path.join( folder, file)
                cleaned_service_date = self.process_service_date()
                folder_name = os.path.basename(self.dispute_no_path)
                self.pdf_rename(file, folder_name, file_path, cleaned_service_date)


    def pdf_rename(self, file, folder_name, file_path, cleaned_service_date):

        if isinstance(file, str) and file.endswith("SHEET.pdf"):
            self.rename_files(f'DISP-{folder_name} ORDER SHEET.pdf', file_path)
                
        elif isinstance(file, str) and file.endswith('NURSE NOTES.pdf'):
            self.rename_files(f'DISP-{folder_name} NURSE NOTE.pdf', file_path)
                
        elif isinstance(file, str) and file.endswith('PHYSICIAN NOTES.pdf'):
            self.rename_files(f'DISP-{folder_name} PHYSICIAN NOTE.pdf', file_path)
                
        elif isinstance(file, str) and file.endswith('REPORT.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('VIEWER.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('VIEW.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('RECORDS OBS.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('Record.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('RECORDS.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('BULK PRINT.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('MEDICAL BULK.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('Rcds.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('Records.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('Report.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith(cleaned_service_date + '.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)

        elif isinstance(file, str) and file.endswith('REPORTS.pdf'):
            self.read_file(file, 'Physician Clinical', 'Milestone Report', file_path, folder_name)
        else:
            self.detect_image(file_path, folder_name)
