from FileProcessor import FileProcessor
from FolderManager import FolderManager
from TypesDefinition import CheckMode, ProcessType
# import pip_system_certs #dummy import to ensure the package is installed

try:
    import pip_system_certs.wrapt_requests
except ImportError:
    import os
    import ssl
    # Set up SSL context manually if the module fails to import
    ssl_context = ssl.create_default_context()
    ssl_context.load_default_certs()
    # Make this the default context
    ssl._create_default_https_context = lambda: ssl_context
    print("Using system certificates directly via SSL context")

if __name__ == '__main__':
    print("Initiating .......")
    exit_program = False

    process_type_map = {
                    '0': ('Full Pre-processor (including process 1 through 6)', 'v1.0', ProcessType.full_process),
                    '1': ('Pre-checking', '1.0', ProcessType.pre_check),
                    '2': ('Scoping', 'v1.0', ProcessType.scoping),
                    '3': ('All Source Data Standardization and Stacking', 'v1.0', ProcessType.standardize_all_and_stack),
                    '4': ('Duplication Search Report', 'v1.0', ProcessType.dup_search_and_compare),
                    '5': ('Item Master Match Report', 'v1.0', ProcessType.itemmast_search_and_compare),
                    '6': ('Replacement Contract Check (only support 1:1 replacement) Report', 'v1.0', ProcessType.replacement_contract_pair_check),
                    '7': ('Run All Reports (process 3 through 6, use this option if you have run through process 1 and 2 previously)', 'v1.0', ProcessType.ccx_dup_search_and_itemmast_match),
                    '8': ('residue_distribution', 'TBI', ProcessType.residue_distribution)
                    }
   
    while not exit_program:
        print("=========================================================")
        print("Let's set up the project folders for this pre-processor")
        manufacturer_name = input('Please enter the manufacturer name (Recommended: copy and paste CCX manufacturer name as it is, remove trailing "." if there is any) of the contract(s) to be pre-processed: ').strip()
        contract_name = input('Please enter the contract number to be pre-processed (if mulitple contracts are to be processed or we do not know the contract number, put a unique descriptive phrase or the Wrike task number to help locate your project): ').strip()
        folder_manager = FolderManager(manufacturer_name, contract_name)
        folder_manager.create_folders()
        print(f"All sub-folders are created.")
        print("==========================================================")
    
        # folder_manager = FolderManager('Bard Medical Division', 
        #                                'L0000000000052')y
        # folder_manager.create_folders()
    
        for key, vals in process_type_map.items():
            if process_type_map[key][1] != 'TBI':
                print(f"Process {key}: {vals[0]}")
        process_to_run = input('key in the process number we want to proceed with: ')

        if (process_to_run in process_type_map) and (process_type_map[process_to_run][1] != 'TBI'):
            process_name, version, process_type = process_type_map[process_to_run]
            if process_type == '1':
                preprocessor = FileProcessor(folder_manager, check_mode = CheckMode.MFN_RF, data_caching = False)
            else:
                shared_data_ready = FileProcessor.check_shared_data()
                while shared_data_ready == False:
                    print("Please check the 'SHARED_DATA' folder and ensure we have all data collections ready before running the program.")
                    shared_data_refreshed = input("All data collections are ready/refreshed? (Y/N):")
                    if shared_data_refreshed.lower() in ['y', 'yes', 'ready', 'refreshed']:
                        shared_data_ready = FileProcessor.check_shared_data()
                    else:
                        input("Press any key to exit.")
                        exit_program = True
                        break
                print("Loading Infor contract data, this will take a while ...")
                preprocessor = FileProcessor(folder_manager, check_mode = CheckMode.MFN_RF)
            preprocessor.process_files(process_type = process_type)
        else:
            print("module under construction, currenty not supported.")
        
        exit_program = input('Do you have more contract to run? (Y/N):')
        if exit_program.lower() in ['n', 'no', 'exit', 'quit']:
            exit_program = True
            print("Enjoy your day, bye.")
        else:
            new_project = input('Do you want to start a new project? (Y/N):')
            if new_project.lower() == 'y' or new_project.lower() == 'yes':
                exit_program = False
                # reset the FileProcessor object
                preprocessor = None
                folder_manager = None
            else:
                exit_program = True
                preprocessor.reset_cache()
                print("Enjoy your day, bye.")

input("Press any key to exit the Preprocessor program.")
