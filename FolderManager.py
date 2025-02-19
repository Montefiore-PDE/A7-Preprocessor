import os

class FolderManager:
    def __init__(self, 
                 manufacturer: str, 
                 contract: str):
        self.current_directory = os.getcwd()
        self.default_directory = os.path.join(self.current_directory, 'Preprocessor_Data')
        self.input_shared_folder = os.path.join(self.current_directory, 'SHARED_DATA')
        
        self.base_directory = os.path.join(self.default_directory, f'{manufacturer}')
        self.input_ccx_folder = os.path.join(self.base_directory,'ccx')
        self.input_to_process_folder = os.path.join(self.base_directory, contract, 'to_process')
        self.output_folder = os.path.join(self.base_directory, contract, 'output')
        self.temp_folder = os.path.join(self.base_directory, contract, 'temp')
        self.processed_folder = os.path.join(self.base_directory, contract, 'processed')

        self.manufacturer = manufacturer
        self.contract = contract

    def create_folders(self):
        for folder in [self.input_ccx_folder, 
                       self.input_shared_folder, 
                       self.input_to_process_folder,
                       self.output_folder, 
                       self.temp_folder,
                       self.processed_folder]:
            if not os.path.exists(folder):
                os.makedirs(folder, exist_ok=True)
        print(f'project folders created successfully. please find project folder and subfolders here at {self.base_directory}')
        message = f'project folder created successfully at {self.base_directory}'
        return message

    def get_folder_path(self, 
                         folder_type: str):
        if folder_type == 'input_ccx':
            return self.input_ccx_folder
        elif folder_type == 'input_shared':
            return self.input_shared_folder
        elif folder_type == 'input_to_process':
            return self.input_to_process_folder
        elif folder_type == 'output':
            return self.output_folder
        elif folder_type == 'temp':
            return self.temp_folder
        elif folder_type == 'processed':
            return self.processed_folder
        else:
            message = "Error in FolderManager.get_folder_path"
            print("the folder type is not valid, we only have 'input', 'input_ccx', 'input_infor', 'input_to_process', 'output' and 'temp' folders")
            return message
        
    def get_base_folder(self):
        print(f"here is the base directory we are using for this manufacturer: {self.base_directory}")
        return self.base_directory



    

        