from FileProcessor import FileProcessor
from FolderManager import FolderManager
from TypesDefinition import CheckMode, ProcessType

if __name__ == '__main__':
    print("Initiating .......")
    exit_program = False

    process_type_map = {'1': ('pre_check', 'v1.0', ProcessType.pre_check),
                    '2': ('scoping', 'v1.0', ProcessType.scoping),
                    '3': ('standardize_all_and_stack', 'v1.0', ProcessType.standardize_all_and_stack),
                    '4': ('dup_search_and_compare', 'v1.0', ProcessType.dup_search_and_compare),
                    '5': ('itemmast_search_and_compare', 'v1.0', ProcessType.itemmast_search_and_compare),
                    '6': ('replacement_contract_pair_check', 'v1.0', ProcessType.replacement_contract_pair_check),
                    '7': ('ccx_dup_search_and_itemmast_match_and_replacement_check', 'v1.0', ProcessType.ccx_dup_search_and_itemmast_match),
                    '8': ('residue_distribution', 'TBI', ProcessType.residue_distribution),
                    '0': ('full_process', 'v1.0', ProcessType.full_process)}
   
    while not exit_program:
        print("=======================================================")
        print("Let's set up the project folders for the pre-processor")
        manufacturer_name = input('Please enter the manufacturer name (copy and paste CCX string as it is) of the contract(s) to be pre-processed: ').strip()
        contract_name = input('Please enter the contract number to be pre-processed (if mulitple contracts are to be pre-processed, key in "multiple", "cleanup" or any other descriptive phrase that help you to identify your input: ').strip()
        folder_manager = FolderManager(manufacturer_name, contract_name)
        folder_manager.create_folders()
        print("All folders created.")
        print("=========================================================")
    
        # folder_manager = FolderManager('Bard Medical Division', 
        #                                'L0000000000052')
        # folder_manager.create_folders()
    
        for key, vals in process_type_map.items():
            print(f"Process {key}: {vals[0]}")
        process_to_run = input('key in the process number we want to proceed with: ')

        if (process_to_run in process_type_map) and (process_type_map[process_to_run][1] != 'TBI'):
            process_name, version, process_type = process_type_map[process_to_run]
            if process_type == '1':
                preprocessor = FileProcessor(folder_manager, check_mode = CheckMode.MFN_RF, data_caching = False)
            else:
                print("Loading Infor contract data, this will take a while ...")
                preprocessor = FileProcessor(folder_manager, check_mode = CheckMode.MFN_RF)

            preprocessor.process_files(process_type = process_type)
        else:
            print("module under construction, currenty not supported.")
        
        exit_program = input('Do you want to exit the program? (Y/N):')
        if exit_program.lower() == 'y' or exit_program.lower() == 'yes' or exit_program.lower() == 'exit':
            exit_program = True
            print("Bye.")
        else:
            new_project = input('Do you want to start a new project? (Y/N):')
            if new_project.lower() == 'y' or new_project.lower() == 'yes':
                exit_program = False
                # reset the FileProcessor object
                preprocessor = None
                folder_manager = None
            else:
                exit_program = True
                print("Bye.")
