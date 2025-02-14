import FileProcessor
import FolderManager
from TypesDefinition import CheckMode, ProcessType

if __name__ == '__main__':
    # folder_manager = FolderManager('Zimmer', 'cleanup')
    # # pre_checker = FileProcessor(folder_manager)
    # pre_checker = FileProcessor(folder_manager, check_mode = CheckMode.MFN_RF)
    # pre_checker.set_scope()
    # pre_checker.standardize_all_and_stack()

    print("Initiating .......")
    print("=======================================================")
    print("Let's set up the folders first")
    manufacturer_name = input('Please enter the manufacturer name (copy and paste CCX string as it is) of the contract(s) to be pre-processed: ').strip()
    contract_name = input('Please enter the contract number to be pre-processed (if mulitple contracts are to be pre-processed, key in "multiple", "cleanup" or any other descriptive phrase that help you to identify your input: ').strip()
    folder_manager = FolderManager(manufacturer_name, contract_name)
    folder_manager.create_folders()
    print("all folders created, please move files to their respective locations.")
    print(f"Infor contract exports: {folder_manager.get_folder_path('input_shared')}")
    print(f"CCX contract exports under manufactuer {manufacturer_name}: {folder_manager.get_folder_path('input_ccx')}")
    print(f"files to be processed: {folder_manager.get_folder_path('input_to_process')}")
    print("=========================================================")
    

    # folder_manager = FolderManager('Fisher Scientific Healthcare', 
    #                                'PP-LA-508')
    # folder_manager.create_folders()
    

    process_type_map = {'1': ('pre_check', 'v1.0', ProcessType.pre_check),
                        '2': ('scoping', 'v1.0', ProcessType.scoping),
                        '3': ('standardize_all_and_stack', 'v1.0', ProcessType.standardize_all_and_stack),
                        '4': ('dup_search_and_compare', 'v1.0', ProcessType.dup_search_and_compare),
                        '5': ('itemmast_search_and_compare', 'v1.0', ProcessType.itemmast_search_and_compare),
                        '6': ('replacement_contract_pair_check', 'v1.0', ProcessType.replacement_contract_pair_check),
                        '7': ('ccx_dup_search_and_itemmast_match_and_replacement_check', 'v1.0', ProcessType.ccx_dup_search_and_itemmast_match),
                        '8': ('residue_distribution', 'TBI', ProcessType.residue_distribution)
                        }
    
    for key, vals in process_type_map.items():
        print(f"Process {key}: {vals[0]}")
    process_to_run = input('key in the process number we want to proceed with: ')
    if (process_to_run in process_type_map) and (process_type_map[process_to_run][1] != 'TBI'):
        process_name, version, process_type = process_type_map[process_to_run]
        if process_to_run in ['1', '4', '5']:
            check_mode = input("Please select the check mode for the pre-check process, key in 'MFN' or 'MFN RF': ")
            if check_mode.upper() == 'MFN':
                pre_checker = FileProcessor(folder_manager, check_mode = CheckMode.MFN)
                pre_checker.process_files(process_type = process_type)
            elif check_mode.upper() == 'MFN RF':
                pre_checker = FileProcessor(folder_manager, check_mode = CheckMode.MFN_RF)
                pre_checker.process_files(process_type = process_type)
            else:
                print("check mode not recognized, we only accept 'MFN' or 'MFN RF'")
        else:
            pre_checker = FileProcessor(folder_manager)
            pre_checker.process_files(process_type = process_type)
    else:
        print("module under construction, currenty not supported")
