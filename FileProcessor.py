import os
import pandas as pd
import numpy as np
from sentence_transformers import SentenceTransformer
from scipy.spatial import distance
from datetime import datetime
from colorama import init, Fore, Style
from FolderManager import FolderManager
from TypesDefinition import ProcessType, CheckMode, StandardizeTarget, Status
from ReportFurnishing import ReportFurnishing
import warnings

# turn on colorama
init(autoreset=True)

# no show certian warning
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=DeprecationWarning)

# set pandas dispaly options
pd.set_option('display.max_columns', 20)
pd.set_option('display.max_rows', 30)

class FileProcessor:

    _infor_std_cache = None
    _import_std_cache = None

    def __init__(self, 
                folder_manager: FolderManager,
                check_mode: CheckMode = CheckMode.MFN_RF,
                data_caching = True):
        self.folder_manager = folder_manager
        self.check_mode = check_mode
        self.datesig = datetime.today().strftime('%Y%m%d')
        self.today = datetime.today().strftime('%Y-%m-%d')
        self.infor_contract_line_file_name = "ContractLine.csv"
        self.infor_contract_line_import_file_name = "ContractLineImport.csv"
        self.manufacturer_map_file = "Manufacturers.csv"
        self.vendor_map_file = "Suppliers.csv"
        self.itemUOM_file = "ItemUOM.csv"
        self.contract_organization_file = "ContractOrganization.xlsx"
        self.manufacturer = folder_manager.manufacturer
        self.contract = folder_manager.contract
        self.shared_file_path = folder_manager.get_folder_path('input_shared')
        self.ccx_file_path = folder_manager.get_folder_path('input_ccx')
        self.tp_file_path = folder_manager.get_folder_path('input_to_process')
        self.output_file_path = folder_manager.get_folder_path('output')
        self.processed_file_path = folder_manager.get_folder_path('processed')
        self.temp_file_path = folder_manager.get_folder_path('temp')
        self.search_scope = []
        self.tp_std = None
        self.ccx_std = None
        self.stacked_std = None
        self.model = None

        if data_caching == False:
            self.infor_std = None
            self.import_std = None
        else:
            if FileProcessor._infor_std_cache is None:
                FileProcessor._infor_std_cache = self.standardize(StandardizeTarget.INFOR)
            self.infor_std = FileProcessor._infor_std_cache

            if FileProcessor._import_std_cache is None:
                FileProcessor._import_std_cache = self.standardize(StandardizeTarget.IMPORT)
            self.import_std = FileProcessor._import_std_cache


    @classmethod
    def reset_cache(cls):
        cls._infor_std_cache = None
        cls._import_std_cache = None

    def set_check_mode(self):
        check_mode = input("Please select the check mode for the pre-check process, key in 'MFN' or 'MFN RF': ")
        self.check_mode = CheckMode.MFN if check_mode.upper() == 'MFN' else CheckMode.MFN_RF

    
    def MFN_reformat(self, MFN: str):
        """
        Reformat manufacturer part number to a reduced version by 
        1. removing leading zeros for input that only contains digits
        2. removing dashes
        3. removing leading/trailing spaces
        ** note if we see input as like a decimal number 123.100, we will NOT convert it to 123.1, but retain its origianl form
        ** similarly, if input is like 089.800.988, we will NOT remove the leading/trailing zeros
        """
        MFN = MFN.replace('-', '').strip()
        if MFN.isdigit():
            return str(int(MFN))
        return MFN

    def UOM_helper(self, 
                   UOM: str):
        """
        Helper function to standardize UOMs
        """
        # read UOM.csv from infor folder and dup that into the dictionay called translation
        uom_translation = pd.read_csv(os.path.join(self.shared_file_path, 'UOM.csv')).\
                          set_index('see UOM').to_dict()['use UOM']
        
        # apply the translation to the provided UOM
        return uom_translation[UOM.upper().strip()] if UOM.upper().strip() in uom_translation else 'TBD'
    
    
    # this is the main logic loop for all different processes
    # Ideally I should move all the process input here as parameters
    def process_files(self,
                      process_type: ProcessType
                      ):
        if process_type == ProcessType.pre_check:
            print('Initiating pre-check process .......')
            self.pre_check(check_mode = self.check_mode)
        elif process_type == ProcessType.scoping:
            print('Initiating scoping process .......')
            self.scoping()
        elif process_type == ProcessType.standardize_all_and_stack:
            print('Initiating standardize all and stack process .......')
            self.set_scope(search_term = self.manufacturer)
            self.standardize_all_and_stack()
        elif process_type == ProcessType.dup_search_and_compare:
            print('Initializing search and compare process .......')
            self.set_scope(search_term = self.manufacturer)
            self.standardize_all_and_stack()
            self.set_model()
            standard_run = 'yes'
            base_set, search_set_input = 'TP', 'CCX'
            standard_run = input('do we want to run a standard dup search between your to-be processed items Vs. CCX contract items? (yes/no)')
            if standard_run.lower() == 'no' or standard_run.lower() == 'n':
                base_set = input('which set you would like to run as base set? (type one of TP/Infor/Import/CCX)')
                search_set_input = input('which set you would like to run as search set? (type in TP/Infor/Import/CCX, for multiple search sets, seperate by ",")')
                print(f'proceding with base set as {base_set}, search set as {search_set_input}')
            else:
                print('proceeding with standard dup search sets .......')
            self.dup_search_and_compare(check_mode = self.check_mode,
                                        base_set = base_set,
                                        search_set_input = search_set_input)
        elif process_type == ProcessType.itemmast_search_and_compare:
            print('Initiating itemmast search and compare process .......')
            self.set_scope(search_term = self.manufacturer)
            self.standardize_all_and_stack()
            self.set_model()
            self.itemmast_search_and_compare(check_mode = self.check_mode)
        elif process_type == ProcessType.ccx_dup_search_and_itemmast_match:
            print('Initiating pre-processor reporting process .......')
            self.set_scope(search_term = self.manufacturer)
            self.standardize_all_and_stack()
            self.set_model()
            self.dup_search_and_compare(check_mode = self.check_mode,
                            base_set = 'TP',
                            search_set_input = 'CCX')
            self.itemmast_search_and_compare(check_mode = self.check_mode)
            self.replacement_contract_pair_check(check_mode = CheckMode.MFN)
        elif process_type == ProcessType.replacement_contract_pair_check:
            print('Initiating replacement contract pair check process .......')
            self.set_scope(search_term = self.manufacturer)
            self.standardize_all_and_stack()
            self.replacement_contract_pair_check(check_mode = CheckMode.MFN)
        elif process_type == ProcessType.full_process:
            print('Initiating full process for preprocessor .......')
            s_pre_check, s_scoping, s_set_scope, s_std = Status.FAILED, Status.FAILED, Status.FAILED, Status.FAILED
            s_dup_run, s_im_match, s_replace = Status.FAILED, Status.FAILED, Status.FAILED
            file_ready = input(f"please put your file to be processed to {self.tp_file_path}, ready to run (Y/N)?: ")
            if file_ready.lower() == 'yes' or file_ready.lower() == 'y' or file_ready.lower() == 'ready': 
                s_pre_check = self.pre_check(check_mode = self.check_mode)
                while s_pre_check == Status.FAILED: 
                    pre_check_retry = input('Exit or Retry? (E/R)')
                    if pre_check_retry.lower() == 'r' or pre_check_retry.lower() == 'retry':
                        s_pre_check = self.pre_check(check_mode = self.check_mode)
                    else:
                        print("Exit preprocessor, bye.")
                        break
            else:
                print("Exit preprocessor, bye.")
                return None
            
            s_scoping = self.scoping()
            while s_scoping == Status.FAILED: 
                scoping_retry = input('Exit or Retry? (E/R)')
                if scoping_retry.lower() == 'r' or scoping_retry.lower() == 'retry':
                    s_scoping = self.scoping()
                else:
                    print("Exit preprocessor, bye.")
                    break
            
            s_set_scope = self.set_scope()
            if s_set_scope == Status.FAILED:
                print("Exit preprocessor, bye.")
                return None

            s_std = self.standardize_all_and_stack()
            if s_std == Status.FAILED:
                print("Exit preprocessor, bye.")
                return None
            
            self.set_model()
            s_dup_run = self.dup_search_and_compare(check_mode = self.check_mode,
                                                    base_set = 'TP',
                                                    search_set_input = 'CCX')
            if s_dup_run == Status.FAILED:
                print("Duplication search failed.")
                
            s_im_match = self.itemmast_search_and_compare(check_mode = self.check_mode)
            if s_im_match == Status.FAILED:
                print("Itemmast search failed.")
            
            s_replace = self.replacement_contract_pair_check(check_mode = CheckMode.MFN)
            if s_replace == Status.FAILED:
                print("Replacement contract pair check failed.")

        else:
            print(f'Invalid process type: {process_type}')
    
        return None
    
    def pre_check(self, check_mode:CheckMode = CheckMode.MFN_RF):
        """
        Perform pre-check operations on all files for a folder, usualy for the 'to_process' folder.
        We will check for the following:
        - File existence (empty folder will raise message 'no files were found')
        - File extention (only .xlsx files will be processed)
        - File follows a specific template and have all necessary columns
        - Does the imported file(s) contains duplicated items (by manufacturer part number or manufacturer part number reduced)
        """
        print(f'reading input files from {self.tp_file_path}')
        if not os.path.exists(self.tp_file_path):
            print('folder not found')
            return Status.FAILED
        if not os.listdir(self.tp_file_path):
            print('no files were found')
            return Status.FAILED
        
        dfs = []
        checker_null_value, checker_dup_value, checker_unknwon_uom, checker_EA_QOE = False, False, False, False
        for file in os.listdir(self.tp_file_path):
            if not file.endswith('.xlsx'):
                print(f'ignoring {file} because it is not a .xlsx file')
                continue
            file_path = os.path.join(self.tp_file_path, file)
            try:
                df_all = pd.read_excel(file_path, sheet_name = None, dtype = str)
            except PermissionError as e:
                print(f'error: {file} is open, please close and try again.')
                return Status.FAILED
            for tab, df in df_all.items():
                try:
                    df = df[['Mfg Part Num', 
                             'Vendor Part Num', 
                             'Buyer Part Num', 
                             'Description', 
                             'Contract Price', 
                             'UOM', 
                             'QOE', 
                             'Effective Date', 
                             'Expiration Date']].copy()
                    assert(all(df.columns == ['Mfg Part Num',
                                            'Vendor Part Num', 
                                            'Buyer Part Num',
                                            'Description',
                                            'Contract Price',
                                            'UOM',
                                            'QOE',
                                            'Effective Date',
                                            'Expiration Date']))
                except ValueError as e:
                    print(f'error: {file}_{tab} does not have the correct columns, please check and use template and try again.')
                    return Status.FAILED
                df.loc[:, 'Contract Number'] = tab.strip()
                df.loc[:, 'File Name'] = file
                dfs.append(df)
       
        # formamtting
        # now everything should already get combined into one dataframe
        df_combined = pd.concat(dfs, ignore_index = True)
        df_combined.loc[:, 'MFN RF'] = df_combined['Mfg Part Num'].apply(lambda x: self.MFN_reformat(x))
        for col in ['Mfg Part Num', 'Vendor Part Num', 'UOM', 'QOE', 'Description', 'Effective Date', 'Expiration Date']:
            df_combined.loc[:, col] = df_combined[col].apply(lambda x: np.nan 
                                                            if (str(x).strip() == '' or pd.isnull(x)) 
                                                            else x)
        for col in ['Effective Date', 'Expiration Date']:
            df_combined.loc[:, col] = pd.to_datetime(df_combined[col], errors = 'coerce')
        df_combined.loc[:, 'Contract Price'] = df_combined['Contract Price'].apply(lambda x: np.nan
                                                                                   if (x == '' or pd.isnull(x))
                                                                                   else float(x.replace('$', '').replace(',','')))
        df_combined.loc[:, 'UOM STD'] = df_combined['UOM'].apply(lambda x: np.nan if (x == '' or pd.isnull(x))
                                                             else self.UOM_helper(x))
        df_combined.loc[:, 'Buyer Part Num'] = df_combined['Buyer Part Num'].fillna('').str.strip()
        df_combined.loc[:, 'seq'] = df_combined.groupby(['Contract Number']).cumcount('Contract Number') + 1
        
        
        # input data checks
        # only output a few columns for pre-view purpose
        cols_to_display = ['Mfg Part Num', 'Contract Number', 'seq', 'Contract Price', 'UOM', 'QOE']
        # display values that need to be filled
        rows_with_nulls = df_combined[df_combined.isnull().any(axis=1)]
        rows_with_nulls_index = rows_with_nulls.index
        if len(rows_with_nulls) > 0:
            print("see below for rows with missing values, fill those in and try again")
            print(rows_with_nulls[cols_to_display])
            checker_null_value = False
            # rows_with_nulls.to_csv(os.path.join(temp_folder_path, 'rows_with_nulls.csv'), index = False)
        else:
            checker_null_value = True

        # check for duplications
        check_mode = self.check_mode
        df_combined.loc[:, 'dup count (by MFN RF)'] = df_combined.groupby(['MFN RF'])['seq'].transform('count')
        df_combined.loc[:, 'dup count (by MFN)'] = df_combined.groupby(['Mfg Part Num'])['seq'].transform('count')
        df_combined.sort_values(by = ['dup count (by MFN RF)', 
                                      'dup count (by MFN)',
                                      'MFN RF', 
                                      'Mfg Part Num', 
                                      'Contract Number'], 
                                ascending = [False, False, True, True, True], inplace = True)
        if check_mode == CheckMode.MFN_RF:
            potential_duplicates = df_combined[df_combined['dup count (by MFN RF)'] > 1]
            potential_duplicates_index = potential_duplicates.index
            if len(potential_duplicates) > 0:
                print("see below for potential duplicated items in your input files, if those are true duplicates, remove from input files and try again with check_mode set up to 'MFN'")
                print(potential_duplicates[cols_to_display])
                checker_dup_value = False
                # potential_duplicates.to_csv(os.path.join(temp_folder_path, 'potential_duplicates.csv'), index = False)
            else:
                checker_dup_value = True
        elif check_mode == CheckMode.MFN:
            duplicates = df_combined[df_combined['dup count (by MFN)'] > 1]
            duplicates_index = duplicates.index
            if len(duplicates) > 0:
                print("see below for duplicated items in your input files, remove from input files and try again")
                print(duplicates[cols_to_display])
                checker_dup_value = False
                # duplicates.to_csv(os.path.join(temp_folder_path, 'duplicates.csv'), index = False)
            else:
                checker_dup_value = True

        # check for UOMs
        unknown_uom = df_combined[df_combined['UOM STD'] == 'TBD']
        unknown_uom_index = unknown_uom.index
        if len(unknown_uom) > 0:
            print("see below for UOMs that need to be standardized, please confirm those are correct UOMs and try again")
            print(unknown_uom[cols_to_display])
            checker_unknwon_uom = False
            # potential_duplicates.to_csv(os.path.join(temp_folder_path, 'unknown_uom.csv'), index = False)
        else:
            checker_unknwon_uom = True

        # check for EA and its respective QOE is 1
        EA_QOE_1 = df_combined[(df_combined['UOM STD'] == 'EA') & (df_combined['QOE'] != '1')]
        EA_QOE_1_index = EA_QOE_1.index
        if len(EA_QOE_1) > 0:
            print("see below for items with UOM EA but QOE not equal to 1, please confirm and try again")
            print(EA_QOE_1[cols_to_display])
            checker_EA_QOE = False 
        else:
            checker_EA_QOE = True
        

        # output the pre_checked - deduped file
        # if all checks passsed
        if checker_null_value and checker_dup_value and checker_unknwon_uom and checker_EA_QOE:
            print(f"Pre-checking {Fore.LIGHTGREEN_EX}***PASSED***{Style.RESET_ALL}, preparing the combined file ......")
            all_items = df_combined.drop_duplicates(subset = ['Mfg Part Num', 
                                                            'Contract Number',
                                                            'Contract Price',
                                                            'UOM',
                                                            'QOE'],
                                                    keep = 'first')
            for col in ['Effective Date', 'Expiration Date']:
                all_items.loc[:, col] = pd.to_datetime(all_items[col], errors = 'coerce')
                all_items.loc[:, col] = all_items[col].apply(lambda x: x.to_pydatetime().strftime('%Y-%m-%d') if not pd.isnull(x) else '1900-01-01')
            
            # archive the old input file, shift the combined input file to be the new round of input
            for file in os.listdir(self.tp_file_path):
                if file.endswith('.xlsx'):
                    os.replace(os.path.join(self.tp_file_path, file), 
                              os.path.join(self.processed_file_path, f'archived_{self.datesig}_{file}'))
            print("Old input file(s) are archived under 'processed' folder")
            
            # output the pre_checked file and name it as TP_INPUT_prechecked.xlsx
            all_items.to_excel(os.path.join(self.tp_file_path, 
                              f'TP_INPUT_prechecked.xlsx'), 
                              index = False)
            print("All items to pre-process are saved as 'TP_INPUT_prechcked.xlsx' in the 'to_process' folder for further processing")
            return Status.SUCCESS
        else:
            # output the pre_checked df_combined to temp folder, directly mark problems on the combined file
            df_combined.loc[rows_with_nulls_index, 'Missing Value'] = 'Check missing value'
            if check_mode == CheckMode.MFN_RF:
                df_combined.loc[potential_duplicates_index, 'Potential Duplicate'] = 'Check potential duplicate'
            elif check_mode == CheckMode.MFN:
                df_combined.loc[duplicates_index, 'Duplicate'] = 'Check duplicate'
            df_combined.loc[unknown_uom_index, 'Unknown UOM'] = 'Check UOM'
            df_combined.loc[EA_QOE_1_index, 'EA QOE not 1'] = 'Check EA QOE'
            df_combined.to_excel(os.path.join(self.temp_file_path, 
                                f'failed_prechecking_{self.manufacturer}_{self.contract}_{self.datesig}.xlsx'), 
                                index = False)
            # do not archive the input file(s), keep the file as it is for user to check and make necessary changes
            print(f"Pre-check {Fore.LIGHTRED_EX}***FAILED***{Style.RESET_ALL}, please carefully review console message and check the temp folder report for more details, once problems fixed, try again.")
        
        return Status.FAILED
    
    def standardize_helper(self, 
                           std_df: pd.DataFrame):
        """
        Standardize the std_df by stripping, type-changing, and filing missing values
        """
        # run stripping
        for col in std_df.columns:
            std_df.loc[:, col] = std_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            std_df.loc[:, col] = std_df[col].apply(lambda x: '' if (isinstance(x, str) and pd.isnull(x)) else x)
        # take care of nan values
        std_df.loc[:, 'OnHold'] = std_df['OnHold'].fillna('No')
        std_df.loc[:, 'ActiveLine'] = std_df['ActiveLine'].fillna('Yes')
        std_df.loc[:, 'ContractLineState'] = std_df['ContractLineState'].fillna('Active')
        std_df.loc[:, 'Contract.ContractStatus'] = std_df['Contract.ContractStatus'].fillna('Active')
        std_df.loc[:, 'FileName'] = std_df['FileName'].fillna('Not Applicable')
        std_df.loc[:, 'Contract'] = std_df['Contract'].fillna('Not Applicable')
        std_df.loc[:, 'ContractImport'] = std_df['ContractImport'].fillna('Not Applicable')
        std_df.loc[:, 'ItemType'] = std_df['ItemType'].fillna('')
        std_df.loc[:, 'IN'] = std_df['IN'].fillna('')
        std_df.loc[:, 'VN'] = std_df['VN'].fillna(std_df['MFN'])
        # take care of column type
        for col in ['UnitCost', 'QOE']:
            std_df.loc[:, col] = std_df[col].apply(lambda x: float(str(x).replace('$','').replace(',','')))
        for col in ['Effective Date', 'Expiration Date']:
            std_df.loc[:, col] = pd.to_datetime(std_df[col], errors = 'coerce')
            std_df.loc[:, col] = std_df[col].apply(lambda x: x.strftime('%Y-%m-%d') if not pd.isnull(x) else '1900-01-01')
        std_df.loc[:, 'Contract Line'] = std_df['Contract Line'].astype(str)
        # make upper case for all contract number
        std_df.loc[:, 'Contract Number'] = std_df['Contract Number'].apply(lambda x: str(x).upper() if not pd.isnull(x) else '')
        # add reduced manufacturer number
        std_df.loc[:, 'MFN RF'] = std_df['MFN'].apply(lambda x: self.MFN_reformat(x) if isinstance(x, str) else '')
        # add countS
        std_df.loc[:, 'count'] = 1
        # compute expiration flag
        std_df.loc[:, 'ExpiredFlag'] = std_df['Expiration Date'].apply(lambda x: 'Expired' if x < self.today
                                                                       else 'Non Expired')
        # compute overall active rank (1 as active overall, 2 hit some type of deactivation flag)
        ind_active = \
        std_df[(std_df['OnHold'] == 'No') &
                (std_df['ActiveLine'] == 'Yes') &
                (std_df['ContractLineState'] == 'Active') &
                (std_df['Contract.ContractStatus'] == 'Active') &
                (std_df['ExpiredFlag'] == 'Non Expired')].index
        std_df.loc[ind_active, 'Active Rank'] = '1'
        std_df.loc[:, 'Active Rank'] = std_df['Active Rank'].fillna('2')
                
        return std_df

    def split_manufacturerinformation(self, 
                                      import_df: pd.DataFrame):
        """
        Split ManufacturerInformation column into Manufacturer and ManufacturerNumber
        """
        import_df.loc[:, 'Manufacturer'] = import_df['ManufacturerInformation'].apply(lambda x: str(x)[:4])
        import_df.loc[:, 'ManufacturerNumber'] = import_df['ManufacturerInformation'].apply(lambda x: str(x)[4:])
        return import_df
    
    def standardize(self, 
                    target_file: StandardizeTarget,
                    vendor_name: str = 'TBD'):
        """
        Standardize all files for a specific input type, we offer the following options:
        1. Infor: standardize Infor files
        2. Import: standardize Import files
        3. CCX: standardize CCX files
        4. TP: standardize files to pre-process
        """
        std_cols = ['Contract Number', 'MFN', 'VN', 'IN', 'Description', 'UnitCost', 'UOM', 'QOE',
            'Effective Date', 'Expiration Date', 'Contract Line', 'Manufacturer', 'Vendor',
            'Contract', 'ItemType', 'OnHold', 'ActiveLine', 'ContractLineState', 'Contract.ContractStatus',
            'ContractImport', 'FileName']
        
        
        if target_file == StandardizeTarget.INFOR:
            print("standardizing Infor files ......")
            file_path = self.shared_file_path
            file_name = self.infor_contract_line_file_name
            if not os.path.exists(file_path):
                print(f"Folder '{file_path}' does not exist, please check the folder path")
                return Status.FAILED
            try:
                infor_df = pd.read_csv(os.path.join(file_path, file_name), dtype = str)
            except FileNotFoundError as e:
                print(f"File '{file_name}' does not exist, check infor contract line download and try again")
                return Status.FAILED
            for col in ['Vendor','Manufacturer', 'ManufacturerNumber', 'VendorItem']:
                infor_df.loc[:, col] = infor_df[col].fillna('unknown')
            for col in ['Contract Import', 'File Name']:
                infor_df.loc[:, col] = np.nan
            infor_cols_to_take = ['Contract.WorkingContractID',
                                  'ManufacturerNumber', 'VendorItem', 'ItemNumber',
                                  'ItemDescription', 'BaseCost', 'UOM', 'DerivedUOMConversion',
                                  'EffectiveDate', 'ExpirationDate', 'ContractLine',
                                  'Manufacturer', 'Vendor',
                                  'Contract', 'ItemType', 'OnHold', 'ActiveLine', 
                                  'ContractLineState', 'Contract.ContractStatus',
                                  'Contract Import', 'File Name']
            infor_std = infor_df[infor_cols_to_take].copy()
            infor_std.columns = std_cols
            infor_std.loc[:, 'Source System'] = 'Infor'
            infor_std.loc[:, 'seq'] = ['i' + str(i) for i in range(len(infor_std))]
            infor_std = self.standardize_helper(infor_std)
            print('Infor files standardized successfully')
            return infor_std

        elif target_file == StandardizeTarget.IMPORT:
            print("standardizing Import files ......")
            file_path = self.shared_file_path
            file_name = self.infor_contract_line_import_file_name
            if not os.path.exists(file_path):
                print(f"Folder '{file_path}' does not exist, please check the folder path")
                return Status.FAILED
            try:
                import_df = pd.read_csv(os.path.join(file_path, file_name), dtype = str)
            except FileNotFoundError as e:
                print(f"File '{file_name}' does not exist, check infor contract line import download and try again")
                return Status.FAILED
            import_df = self.split_manufacturerinformation(import_df)
            for col in ['ContractImport.Vendor', 'VendorItem']:
                import_df.loc[:, col] = import_df[col].fillna('unknown')
            for col in ['ItemType', 'OnHold', 'ActiveLine', 'ContractLineState', 'Contract.ContractStatus',
                        'File Name']:
                import_df.loc[:, col] = np.nan
            import_cols_to_take = ['ContractImport.WorkingContractID', 
                                   'ManufacturerNumber', 'VendorItem', 'ItemNumber',
                                   'ItemDescription', 'BaseCost', 'UOM', 'UOMConversion',
                                   'EffectiveDate', 'ExpirationDate', 'ContractLineImport',
                                   'Manufacturer', 'ContractImport.Vendor',
                                   'ContractRel.Contract', 'ItemType', 'OnHold', 'ActiveLine',
                                   'ContractLineState', 'Contract.ContractStatus',
                                   'ContractImport', 'File Name']
            import_std = import_df[import_cols_to_take].copy()
            import_std.columns = std_cols
            import_std.loc[:, 'Source System'] = 'Import'
            import_std.loc[:, 'seq'] = ['imp' + str(i) for i in range(len(import_std))]
            import_std = self.standardize_helper(import_std)
            print('Import files standardized successfully')
            return import_std
            
        elif target_file == StandardizeTarget.CCX:
            print("standardizing CCX files ......")
            file_path = self.ccx_file_path
            ccx_dfs = []
            if not os.path.exists(file_path):
                print(f"Folder '{file_path}' does not exist, please check the folder path")
                return Status.FAILED
            for file in os.listdir(file_path):
                if not file.endswith('.xlsx'):
                    print(f'ignoring {file} because it is not a .xlsx file')
                    continue
                try:
                    ccx_df = pd.read_excel(os.path.join(file_path, file), dtype = str)
                    ccx_df.loc[:, 'ContractLine'] = [str(i+1) for i in range(len(ccx_df))]
                    ccx_dfs.append(ccx_df)
                except FileNotFoundError as e:
                    print(f"File '{file}' does not exist, check CCX contract download and try again")
                    return Status.FAILED
            ccx_df = pd.concat(ccx_dfs, ignore_index = True)
            for cols in ['Contract', 'ItemType', 'OnHold', 'ActiveLine', 
                         'ContractLineState', 'Contract.ContractStatus','ContractImport',
                         'File Name']:
                ccx_df.loc[:, cols] = np.nan
            ccx_cols_to_take = ['Contract Number', 
                                'Mfg Part Num', 'Vendor Part Num', 'Buyer Part Num', 
                                'Description', 'Contract Price', 'UOM', 'QOE',
                                'Effective Date', 'Expiration Date', 'ContractLine',
                                'Manufacturer', 'Vendor', 
                                'Contract', 'ItemType', 'OnHold', 'ActiveLine', 
                                'ContractLineState', 'Contract.ContractStatus', 
                                'ContractImport', 'File Name']
            ccx_std = ccx_df[ccx_cols_to_take].copy()
            ccx_std.columns = std_cols
            ccx_std.loc[:, 'Source System'] = 'CCX'
            ccx_std.loc[:, 'seq'] = ['ccx' + str(i) for i in range(len(ccx_std))]
            ccx_std = self.standardize_helper(ccx_std)
            print('CCX files standardized successfully')
            return ccx_std
        
        elif target_file == StandardizeTarget.TP:
            print("standardizing files to pre-process ......")
            file_path = self.tp_file_path
            file_name = f"TP_INPUT_prechecked.xlsx"
            if not os.path.exists(file_path):
                print(f"Folder '{file_path}' does not exist, please check the folder path")
                return Status.FAILED
            try:
                tp_df = pd.read_excel(os.path.join(file_path, file_name), dtype = str)
            except FileNotFoundError as e:
                print(f"File '{file_name}' does not exist, check the output from pre_check process and try again")
                return Status.FAILED
            tp_df.loc[:, 'Vendor'] = vendor_name
            tp_df.loc[:, 'Manufacturer'] = self.manufacturer
            tp_df.loc[:, 'ContractLine'] = tp_df.groupby(['Contract Number']).cumcount() + 1
            for col in ['Contract', 'ItemType', 'OnHold', 'ActiveLine', 
                        'ContractLineState', 'Contract.ContractStatus','ContractImport']:
                tp_df.loc[:, col] = np.nan
            tp_cols_to_take = ['Contract Number', 
                               'Mfg Part Num', 'Vendor Part Num', 'Buyer Part Num', 
                               'Description', 'Contract Price', 'UOM STD', 'QOE', 
                               'Effective Date', 'Expiration Date', 'ContractLine',
                               'Manufacturer', 'Vendor',
                               'Contract', 'ItemType', 'OnHold', 'ActiveLine',
                               'ContractLineState', 'Contract.ContractStatus',
                               'ContractImport', 'File Name']
            tp_std = tp_df[tp_cols_to_take].copy()
            tp_std.columns = std_cols
            tp_std.loc[:, 'Source System'] = 'TP'
            tp_std.loc[:, 'seq'] = ['tp' + str(i) for i in range(len(tp_std))]
            tp_std = self.standardize_helper(tp_std)
            print('TP files standardized successfully')
            return tp_std

        else:
            print("standardize target not recognized, please check the target file and try again")
        
        return Status.FAILED
    
    def manufacturer_map(self, mfn: str):
        """
        Import the manufacturer map file and return a DataFrame containing manufacturer and manufacturer name
        """
        mfn_map_df = pd.read_csv(os.path.join(self.shared_file_path, self.manufacturer_map_file),
                                 dtype = str)
        mfn_map_df = mfn_map_df[['Manufacturer', 'Description']].copy()
        mfn_map = mfn_map_df.set_index('Manufacturer')['Description'].to_dict()
        
        return mfn_map.get(mfn, 'TBD')
    
    def vendor_map(self, vendor: str):
        """
        Import the vendor map file and return a DataFrame containing vendor and vendor name
        """
        vendor_map_df = pd.read_csv(os.path.join(self.shared_file_path, self.vendor_map_file),
                                      dtype = str)
        # change the import to dictionary so that the key is "Vendor", values are ["Vendor.VendorName", "Representative Text"]
        vendor_map_df = vendor_map_df[['Vendor', 'Vendor.VendorName', 'RepresentativeText']].copy()
        vendor_map_df.loc[:, 'Vendor'] = vendor_map_df['Vendor'].fillna('unknown_vendorID')
        vendor_map = {k: [v, r] for k, v, r in zip(vendor_map_df['Vendor'], 
                                                   vendor_map_df['Vendor.VendorName'],
                                                   vendor_map_df['RepresentativeText'])}
        
        return vendor_map.get(vendor, ['TBD', 'TBD'])
    
    def scoping(self):
        """
        Define the searching space for downstream contract line comparesions.
        Steps:
        1. Import all the the TP items (files to pre-process)
        2. Import infor contract lines
        3. Run the to-be pre-processed items against full infor contract lines to determine the potential overlaps
        4. Export the scoping_manual_review.xlsx file so we review the items by contract and constrain downstream searches
        on only the contracts belong to the same manufacturer OR has overlaps between 1&2
        
        standardized columns are
        std_cols = ['Contract Number', 'MFN', 'VN', 'IN', 'Description', 'UnitCost', 'UOM', 'QOE',
        'Effective Date', 'Expiration Date', 'Contract Line', 'Manufacturer', 'Vendor',
        'Contract', 'ItemType', 'OnHold', 'ActiveLine', 'ContractLineState', 'Contract.ContractStatus',
        'ContractImport', 'FileName', 'ExpirationFlag', 'Active Rank']

        the scoping method will always take the reduced format of manufacturer as the join key
        """
        if self.tp_std is None:
            tp_std = self.standardize(StandardizeTarget.TP)
        else:
            tp_std = self.tp_std
        
        if self.infor_std is None:
            infor_std = self.standardize(StandardizeTarget.INFOR)
        else:
            infor_std = self.infor_std
        
        tp_mini = tp_std[['Contract Number', 'seq', 'MFN', 'VN', 'Description', 'UnitCost', 'MFN RF']].copy()
        mfn_to_check = set(tp_mini['MFN'])
        mfn_rf_to_check = set(tp_mini['MFN RF'])
        infor_cols_to_take = ['Description', 'UnitCost', 'MFN', 'MFN RF', 'VN', 'Contract Number',
                              'Manufacturer', 'Vendor']
        infor_interferring = infor_std[(infor_std['MFN'].isin(mfn_to_check) | 
                                       infor_std['MFN RF'].isin(mfn_rf_to_check) |
                                       infor_std['VN'].isin(mfn_to_check) |
                                       infor_std['VN'].isin(mfn_rf_to_check)) & 
                                       (infor_std['Active Rank'] == '1')][infor_cols_to_take].copy()
        
        infor_interferring.loc[:, 'VendorName'] = infor_interferring['Vendor'].apply(lambda x: self.vendor_map(x)[0])
        infor_interferring.loc[:, 'Supplier'] = infor_interferring['Vendor'].apply(lambda x: self.vendor_map(x)[-1])
        infor_interferring.loc[:, 'ManufacturerName'] = infor_interferring['Manufacturer'].apply(lambda x: self.manufacturer_map(x))
        tp_mini.loc[:, 'Take'] = ''
        scoping_df = tp_mini.merge(infor_interferring,
                                   on = ['MFN RF'], 
                                   how = 'inner',
                                   suffixes=('_ccx', '_infor'))
        scoping_df.loc[:, 'Same MFN'] = scoping_df['MFN_ccx'] == scoping_df['MFN_infor']
        scoping_df_colarrg = ['Contract Number_ccx', 'seq',
                              'MFN_ccx', 'VN_ccx', 'Description_ccx', 'UnitCost_ccx', 
                              'Description_infor', 'UnitCost_infor',
                              'Take', 'MFN RF', 'Same MFN', 
                              'MFN_infor', 'VN_infor', 'Contract Number_infor',
                              'Supplier',
                              'Manufacturer', 'ManufacturerName',
                              'Vendor', 'VendorName']
        scoping_df = scoping_df[scoping_df_colarrg].copy()
        scoping_df.sort_values(by = ['Supplier', 'Contract Number_infor'], inplace = True)
        
        with pd.ExcelWriter(os.path.join(self.output_file_path,
                                         f'scoping_manual_review_{self.datesig}.xlsx')) as writer:
            scoping_df.to_excel(writer, sheet_name = 'ToReview', index = False)
            headers = ['Contract Number_infor', 'Supplier']
            empty_df = pd.DataFrame(columns = headers)
            empty_df.to_excel(writer, sheet_name = 'ContractToTake', index = False)
        print(f"""scoping file is saved as 
{Fore.LIGHTGREEN_EX}'scoping_manual_review_{self.datesig}.xlsx'{Style.RESET_ALL} in the 
output folder for manual review. Once reviewed, copy and paste infor contract(s) 
contain duplicates to tab 'ContractToTake', and rename the file 
to {Fore.LIGHTGREEN_EX}'scoping_manual_reviewed.xlsx'{Style.RESET_ALL}""".replace("\n", ""))
        
        scoping_reviewed = input("Have you reviewed the scoping file and identified the contract we want to include in subsequent steps? (yes/no)").lower()    
        if scoping_reviewed == 'yes' or scoping_reviewed == 'y':
            try:
                scoping_df = pd.read_excel(os.path.join(self.output_file_path, 
                                                        f'scoping_manual_reviewed.xlsx'),
                                           dtype = str)
            except FileNotFoundError as e:
                print("file not found, please make sure the file is saved under name 'scoping_manual_reviewed.xlsx' and you have listed the contract needed in 'Sheet2' under column 'Contract Number_y'." )
                return Status.FAILED
            return Status.SUCCESS
        
        return Status.FAILED
    
    
    def set_scope(self):
        """True function to set scope
        1. use the set_scope_helper to get and display the contract we will download from CCX
        2. if we don't like the results we could retry until we are satisfied
        the function will eventually set the searching scope to the object"""
        
        search_scope = self.set_scope_helper()
        
        # if user is not satisfied with the search scope, we will retry differernt search term
        # until the re_scoping input is set to 'N'
        re_scoping = 'Y'
        re_scoping = input("Want to try different search term(s)? (Y/N)")
        while re_scoping.lower() in ['yes', 'y']:
            search_scope = self.set_scope_helper()
            re_scoping = input("Want to try different search term(s)? (Y/N)")
            if re_scoping.lower() in ['no', 'n']:
                break
      
        if len(search_scope) > 0 and search_scope is not None:  
            self.search_scope = search_scope
            ccx_downlaod = input("Contract(s) listed above are downloaded from CCX? (Y/N)")
            if ccx_downlaod.lower() in ['yes', 'y']:
                print(f"searching scope is set to {search_scope}")
                return Status.SUCCESS          
        
        return Status.FAILED
    
    
    def set_scope_helper(self, search_term: str = None):
        """Set searching space by combining
        1. the contract numbers screened after scoping manual review
        2. the contract numbers registered under the targeted manufacturer on ccx
        when the searching space is computed, go CCX to get all the relevarnt contracts downloaded
        then we have the data ready to do duplication search"""
        scoping_reviewed, contract_inscope_ccx = [], []
        # from manual review
        scoping_manual_reviewed = pd.read_excel(os.path.join(self.output_file_path, 
                                                            f'scoping_manual_reviewed.xlsx'),
                                                dtype = str,
                                                sheet_name = 'ContractToTake')
        if len(scoping_manual_reviewed) == 0:
            pass
        else:
            scoping_reviewed = scoping_manual_reviewed['Contract Number_infor'].tolist()

        # from CCX search
        contract_organization_df = pd.read_excel(os.path.join(self.shared_file_path,
                                                              self.contract_organization_file),
                                                dtype = str)
        for col in ['Manufacturer', 'Vendor', 'ERP Vendor Number']:
            contract_organization_df.loc[:, col] = contract_organization_df[col].fillna('')
        
        # make selections based on our search criteria
        search_term_set_up = input(f"Default scope to current manufacturer '{self.manufacturer}'? (yes/no)")
        if search_term_set_up.lower() == 'yes' or search_term_set_up.lower() == 'y':
            search_term = self.manufacturer
        elif search_term_set_up.lower() == 'no' or search_term_set_up.lower() == 'n':
            search_term = input("Enter the manufacturer name (if more than one, using pipe '|' to separate):")
        else:
            print(f"invalid input. Defaulted to current manufacturer '{self.manufacturer}'.")
            search_term = self.manufacturer

        manufacturer_to_look = contract_organization_df['Manufacturer'].str.contains(search_term, 
                                                                                     case = False)
        vendor_to_look = contract_organization_df['Vendor'].str.contains(search_term, 
                                                                         case = False)
        ERP_linked = contract_organization_df['ERP Vendor Number'] != ''
        contract_inscope_ccx = contract_organization_df[(manufacturer_to_look | vendor_to_look) & 
                                                        ERP_linked]['Contract Number'].tolist()
        
        all_contracts_to_look = sorted(list(set(contract_inscope_ccx + scoping_reviewed)))
        all_contracts_to_look = sorted(list(set([i.upper().strip() for i in all_contracts_to_look])))
        # fetch the manufacturer attached to these contract
        manufacturer_contract_map = contract_organization_df[['Contract Number', 'Manufacturer']].copy()
        manufacturer_contract_map.drop_duplicates(subset = ['Contract Number'], 
                                                  keep = 'first', inplace = True)
        manufacturer_contract_dict = {k: v for k, v in zip(manufacturer_contract_map['Contract Number'].apply(lambda x: x.upper()), 
                                                           manufacturer_contract_map['Manufacturer'])}
        print("Contract(s) to be screened:")
        for i, contract in enumerate(all_contracts_to_look):
            print(f"{i+1:<3} {contract:<40}, {manufacturer_contract_dict.get(contract.upper(), 'Unknown')}")
        
        return all_contracts_to_look
    
    def standardize_all_and_stack(self):
        """Standardize all four major sources of input tables
        1. to process (the user submission)
        2. Infor ContractLine daily output
        3. Infor ContratLineImport daily output
        4. CCX downloaded contract
        when all files output to temp foloder, many subsequent comparisons can be made in 
        various flexible ways"""
        if self.tp_std is None:
            tp_std = self.standardize(StandardizeTarget.TP)
        else:
            tp_std = self.tp_std
        if self.infor_std is None:
            infor_std = self.standardize(StandardizeTarget.INFOR)
        else:
            infor_std = self.infor_std
        if self.import_std is None:
            import_std = self.standardize(StandardizeTarget.IMPORT)
        else:
            import_std = self.import_std
        if self.ccx_std is None:
            ccx_std = self.standardize(StandardizeTarget.CCX)
        else:
            ccx_std = self.ccx_std
        
        if self.search_scope is None:
            print("searching scope not set, please run set_scope() function first and try again")
            return Status.FAILED
        
        print("current searching scope set as: ", self.search_scope)

        search_scope = [i.upper() for i in self.search_scope]
        infor_std = infor_std[infor_std['Contract Number'].isin(search_scope)].copy()
        import_std = import_std[import_std['Contract Number'].isin(search_scope)].copy()
        ccx_std = ccx_std[ccx_std['Contract Number'].isin(search_scope)].copy()

        stacked_std = pd.concat([ccx_std, infor_std, import_std, tp_std], ignore_index = True)
        self.stacked_std = stacked_std
        print(f"all file sources standardized, we have {len(stacked_std)} records in total.")
        print(stacked_std.groupby(['Source System']).size().reset_index(name = 'counts'))

        proof = input("do you want to create a data dump for the standardized data used in the project? (Y/N)")
        if proof.lower() == 'yes' or proof.lower() == 'y':
            print(f"""file sources are standardized and stacked togather, 
                  a hard copy will be created and stored under {self.temp_file_path}. 
                  This will take a while.""".replace("\n", ""))
            stacked_std.to_excel(os.path.join(self.temp_file_path, 
                                            f'stacked_std_{self.manufacturer}_{self.contract}_{self.datesig}.xlsx'),
                                            index = False)
        else:
            pass
        return Status.SUCCESS
    
    def set_model(self, model_name:str = 'all-MiniLM-L6-v2'):
        self.model = SentenceTransformer(model_name)
        return Status.SUCCESS
    
    def calc_similarity(self, 
                        desc1: str, 
                        desc2: str):
        similarity = 0
        embeddings = self.model.encode([desc1, desc2])
        cosine_distance = distance.cosine(embeddings[0], embeddings[1])
        similarity = 1 - cosine_distance
        return similarity
    
    def compute_sims_df(self, 
                        to_emb: pd.DataFrame):
        """compute cosine similarity between descriptions in the dataframe
        and create a new column to store the results"""
        if len(to_emb) == 0:
            print("there is no similarity to compute")
            return Status.FAILED
        sims = []
        to_emb.loc[:, 'ref'] = [str(i) for i in range(len(to_emb))]
        total_records_to_process = len(to_emb)
        for i, val in enumerate(to_emb.values):
            s = self.calc_similarity(val[0], val[1])
            sims.append([val[2], s])
            if i%100 == 0:
                print(f'processed {i}/{total_records_to_process} records.', end = "\r")
        print(f'processed {i+1}/{total_records_to_process} records for description similarity.')
        sims_df = pd.DataFrame(sims, columns = ['ref', 'Description Similarity'])
        sims_calc_df = to_emb.merge(sims_df, on = ['ref'], how = 'left')
        sims_calc_df.drop('ref', axis = 1, inplace = True)
        return sims_calc_df

    def dup_search_and_compare(self, 
                               check_mode: CheckMode = CheckMode.MFN_RF,
                               base_set: str = 'TP', 
                               search_set_input: str = 'CCX'):
        """a function compares the left and right dataframe and match the items
        based on match mode.
        1. left_df usually should be the tp_std (to process file)
        2. right_df usually will be ccx_std (ccx files to compare and find duplicates)
        But since the function is general, it can be adapted to any combination of interest
        """
        # if we run to here, we need to make sure we already have everything run up to scope
        # the search set can be have more than one searching data group, just separate the input by comma
        search_set = search_set_input.split(',')
        left_df = self.stacked_std[self.stacked_std['Source System'] == base_set].copy()
        right_df = self.stacked_std[self.stacked_std['Source System'].isin(search_set)].copy()

        left_cols = ['Source System', 'Contract Number', 
                     'MFN', 'VN', 'IN', 'Description',
                     'UnitCost', 'UOM', 'QOE', 
                     'Effective Date', 'Expiration Date', 
                     'seq', 'MFN RF']
        left_df = left_df[left_cols].copy()

        if self.check_mode == CheckMode.MFN_RF:
            dup_found = left_df.merge(right_df, on = ['MFN RF'])
        elif self.check_mode == CheckMode.MFN:
            dup_found = left_df.merge(right_df, on = ['MFN'])
        else:
            print("Invalid check mode. Please use CheckMode.MFN_RF or CheckMode.MFN.")
            return Status.FAILED
        
        if len(dup_found) == 0:
            print("no duplication found in the search set. All good now.")
            return Status.PASS
       
        dup_found.loc[:, 'Same QOE'] = dup_found['QOE_x'] == dup_found['QOE_y']
        dup_found.loc[:, 'Same UOM'] = dup_found['UOM_x'] == dup_found['UOM_y']
        dup_found.loc[:, 'EACostDiff'] = dup_found.apply(lambda x: (x['UnitCost_x']/x['QOE_x'])/(x['UnitCost_y']/x['QOE_y'])
                                                           if ((x['UnitCost_y']/x['QOE_y']) != 0 and x['QOE_x'] != 0)
                                                           else -1, axis = 1)

        # the long way of compute text similarity
        # this can be optimized by only computing the text similarity for contracts that are
        # not exactly coming from the target manufactuer
        # can save a lot of time if the joined dup_found becomes too long
        efficient_mode = False
        if efficient_mode is True:
            # to be implemented
            pass
        else:
            to_emb = dup_found[['Description_x', 'Description_y']].drop_duplicates()
        
        sims_calc_df = self.compute_sims_df(to_emb)
        if sims_calc_df is None:
            dup_found_m = dup_found.copy()
            dup_found_m.loc[:, 'Description Similarity'] = np.nan
        else:
            dup_found_m = dup_found.merge(sims_calc_df, 
                                            on = ['Description_x', 'Description_y'], 
                                            how = 'left')
        dup_found_m.loc[:, 'Description Similarity'] = dup_found_m['Description Similarity'].fillna(1)
        dup_found_m.loc[:, 'Drop'] = ''
        dup_found_m.sort_values(by = ['Description Similarity'], inplace = True)
        dup_found_m.to_excel(os.path.join(self.output_file_path, 
                                          f'dup_search_review_{self.datesig}.xlsx'),
                                          index = False)
        print(f"""Initial search for duplication items completed, results are saved as 
{Fore.LIGHTGREEN_EX}'dup_search_review_{self.datesig}.xlsx'{Style.RESET_ALL} in the temp folder. Please review 
and mark false positive matches under columns 'Drop' with 'x' and rename the 
reviewed file to {Fore.LIGHTGREEN_EX}'dup_search_reviewed.xlsx{Style.RESET_ALL}'""".replace("\n", ""))
        
        duplication_review_completed = "no"
        duplication_review_completed = input("Have we reviewed the duplication search results and rename the file? (yes/no): ")
        if duplication_review_completed.lower() == "yes" or duplication_review_completed.lower() == "y":
            dup_found_reviewed = pd.read_excel(os.path.join(self.output_file_path, 
                                                            'dup_search_reviewed.xlsx'),
                                                             dtype = str)
            dup_found_clean = dup_found_reviewed[dup_found_reviewed['Drop'].fillna('') == ''].copy()
            dup_found_clean.drop(columns = ['Drop'], inplace = True)
            dup_found_clean = dup_found_clean.reset_index(drop = True)
            # mark the lines for review
            for col in ['EACostDiff', 'UnitCost_x', 'UnitCost_y', 'Description Similarity']:
                dup_found_clean.loc[:, col] = dup_found_clean[col].apply(lambda x: round(float(x),2) if not pd.isnull(x) else np.nan)
            for col in ['QOE_x', 'QOE_y']:
                dup_found_clean.loc[:, col] = dup_found_clean[col].apply(lambda x: int(x) if not pd.isnull(x) else 0)
            
            dups_review1 = dup_found_clean['Same UOM'] == "False"
            dups_review2 = dup_found_clean['Same QOE'] == "False"
            dups_review3 = ((dup_found_clean['EACostDiff'] > 2) | (dup_found_clean['EACostDiff'] < 0.5)) & \
                           (dup_found_clean['Manufacturer'] != self.manufacturer)
            dups_review4 = ((dup_found_clean['EACostDiff'] > 1.5) | (dup_found_clean['EACostDiff'] < 0.65)) & \
                           (dup_found_clean['Manufacturer'] == self.manufacturer)
            dups_review5 = (dup_found_clean['UOM_x'] == 'EA') & (dup_found_clean['QOE_x'] != 1)

            dup_review_ind = dup_found_clean[dups_review1 | dups_review2 | dups_review3 | dups_review4 | dups_review5].index
            dup_found_clean.loc[:, 'Action'] = 'Deactivate'
            dup_found_clean.loc[dup_review_ind, 'Action'] = 'Review'

            dup_found_clean.sort_values(by = ['Action'], ascending = [False], inplace = True)
            pre_output = \
            dup_found_clean[dup_found_clean['Active Rank'] == '1'][['MFN_y', 'VN_y',
                            'IN_y', 'Description_y',
                            'UnitCost_y', 'UOM_y', 'QOE_y',
                            'Effective Date_y', 'Expiration Date_y', 
                            'Contract Number_y', 'Contract Line', 'Source System_y',
                            'Action', 
                            'MFN_x', 'VN_x',
                            'Description_x', 'UnitCost_x', 'UOM_x', 'QOE_x',
                            'Contract Number_x',
                            'Same UOM', 'Same QOE', 'EACostDiff', 'Description Similarity']].copy()
            
            # in theory, if later we decide to output in same tab, we can use this format
            # dup_found_clean.to_excel(os.path.join(self.temp_file_path, "just a dummy file.xlsx"))
            print(pre_output.shape)

            # prepare to output the data into different tabs
            to_output = {}
            for contract in pre_output['Contract Number_y'].unique():
                print(contract)
                to_output[contract] = pre_output[pre_output['Contract Number_y'] == contract]
            # output summary information so we know initially how many items on per contract inscope
            to_count_df = self.stacked_std[self.stacked_std['Source System'].isin(search_set + [base_set])].copy()
            
            count_summary = to_count_df[to_count_df['Active Rank'] == '1'].groupby(['Source System', 
                                                                                    'Contract Number',
                                                                                    'Manufacturer']).\
                                                                           agg(line_count = ('seq', 'count')).\
                                                                           reset_index()
            count_intersection = pre_output.groupby(['Source System_y',
                                                     'Contract Number_y']).\
                                            agg(overlap_count = ('Contract Line', 'count')).\
                                            reset_index()
            count_intersection.rename(columns = {'Contract Number_y': 'Contract Number',
                                                 'Source System_y': 'Source System'}, inplace = True)
            count_summary_to_output = \
            count_summary.merge(count_intersection, 
                                on = ['Source System', 'Contract Number'],
                                how = 'left')
            
            # Create an ExcelWriter object
            # with pd.ExcelWriter(os.path.join(self.output_file_path,
            #                                  f'dedup_output_{self.manufacturer}_{self.contract}_{self.datesig}.xlsx')) as writer:
            #     # Write each DataFrame to a separate sheet
            #     count_summary_to_output.to_excel(writer, sheet_name = 'quick_line_count', index = False)
            #     for k, v in to_output.items():
            #         v.to_excel(writer, sheet_name = k, index = False)
            #     dup_found_clean.to_excel(writer, sheet_name = 'all_dup_raw', index = False)
            # print(f"duplication search completed, report will be saved to {self.output_file_path}")
            report_to_write = ReportFurnishing(self.folder_manager)
            report_to_write.make_dedup_report(to_output,
                                              count_summary_to_output,
                                              dup_found_clean)
            return Status.SUCCESS

        else:
            print("Take your time and once the review is done, rerun the process and try again")

        return Status.FAILED
    
    def itemmast_search_and_compare(self,
                                    check_mode: CheckMode = CheckMode.MFN_RF):
        im_df = self.stacked_std[(self.stacked_std['Source System'] == 'Infor') & 
                                 (self.stacked_std['ItemType'] == 'Itemmast') &
                                 (self.stacked_std['Active Rank'] == '1')].copy()
        tp_df = self.stacked_std[self.stacked_std['Source System'] == 'TP']
        tp_cols_to_take = ['Contract Number', 'MFN', 'VN', 'IN', 'Description', 'UnitCost', 'UOM', 'QOE', 
                           'Effective Date', 'Expiration Date', 'seq', 'MFN RF']
        infor_cols_to_take = ['MFN RF', 'Contract Number', 'MFN', 'VN', 'IN', 'Description', 'UnitCost', 'UOM', 'QOE', 'ItemType']
        tp_side = tp_df[tp_cols_to_take].copy()
        im_side = im_df[infor_cols_to_take].copy()

        tp_im = tp_side.merge(im_side, on = ['MFN RF'], how = 'inner')

        # need implementation -- we also have a subset of item master items that are not in im_side (not backed up by contract)
        # in this case, we need to import the VendorItem class and select on items with no contract references and do the
        # second pass of screening
        
        # since im item usually should be relatively small, directly call similarity function instead of 
        # generating the sims_df (this will help us monitor the progress)
        print(f'found {tp_im.shape[0]} of potential item master item.')
        if len(tp_im) == 0:
            print("No Item master item hit, nothing need to be done.")
            return Status.SUCCESS
        tp_im.loc[:, 'Description Similarity'] = tp_im.apply(lambda x: 
                                                             self.calc_similarity(x['Description_x'], 
                                                                                  x['Description_y']),
                                                                                  axis = 1)
        tp_im.rename(columns = {'UOM_x': 'UOM',
                                'IN_y': 'Item'}, inplace = True)
        # read in ItemUOM
        itemUOM = pd.read_csv(os.path.join(self.shared_file_path, self.itemUOM_file), dtype = str)
        itemUOM_cols_to_take = ['Item', 'UnitOfMeasure', 'UOMConversion', 'ValidForBuying', 'Item.Active']
        itemUOM = itemUOM[itemUOM_cols_to_take].copy()
        itemUOM.loc[:, 'UOMConversion'] = itemUOM['UOMConversion'].apply(lambda x: int(float(x.replace(',',''))) if not pd.isnull(x) else 0)
        itemUOM.rename(columns = {'UnitOfMeasure': 'UOM'}, inplace = True)
        valid_buyuom = itemUOM[itemUOM['ValidForBuying'] != 'Not Valid'].copy()
        valid_buyuom.loc[:, 'AllValidBuyUOMandCF'] = valid_buyuom['UOM'].astype(str) + '*' + valid_buyuom['UOMConversion'].astype(int).astype(str)
        all_buyuom = valid_buyuom.groupby(['Item'])['AllValidBuyUOMandCF'].apply(lambda x: ','.join(x)).to_frame().reset_index()
        # merge to tp_im
        im_label = tp_im.merge(valid_buyuom, on = ['Item', 'UOM'], how = 'left').\
                         merge(all_buyuom, on = ['Item'], how = 'left')
        im_label.loc[:, 'IM_check'] = im_label['ValidForBuying'].apply(lambda x: 'Failed' if pd.isnull(x) else 'Passed')
        im_label.loc[:, 'IM_check'] = im_label.apply(lambda x: 'Failed' 
                                                     if x['QOE_x'] != x['UOMConversion'] 
                                                     else x['IM_check'], axis = 1)
        # drop by seq_x and contract to make the output simple
        im_label_simple = \
        im_label[['Contract Number_x', 'MFN_x', 'VN_x', 
                  'Description_x', 'UnitCost_x', 'UOM', 'QOE_x', 'seq',
                  'Effective Date', 'Expiration Date', 'MFN RF', 'Description_y', 'Description Similarity',
                  'MFN_y', 'Item', 'ItemType', 'UOMConversion', 'ValidForBuying',
                  'AllValidBuyUOMandCF_y', 'IM_check']].copy()
        im_label_simple.sort_values(by = ['IM_check'], ascending = [True], inplace = True)

        im_label_simple = im_label_simple.drop_duplicates(subset = ['seq', 'Item'])
        im_label_simple.loc[:, 'Numbers of Item Matched'] = im_label_simple.groupby(['seq'])['Item'].transform('count')
        im_label_simple.to_excel(os.path.join(self.output_file_path,
                                                f'itemmast_matched_{self.manufacturer}_{self.contract}_{self.datesig}.xlsx'),
                                                index = False)
        print('item master matching completed successfully, results are saved to output folder.')
        
        return Status.SUCCESS
    
    def replacement_contract_pair_check(self,
                                        check_mode: CheckMode = CheckMode.MFN_RF):
        """compare two contracts: TP - new, replacement ccx - old contract to be replaced to identify
        1. items only lives on old contract
        2. items only lives on old contract and marked as itemmast on Infor"""
        replaced_contract = input("please enter the replacement contract number: ")
        replaced_contract = replaced_contract.strip().upper()
        replacement_ccx_df = self.stacked_std[(self.stacked_std['Source System'] == 'CCX') & 
                                              (self.stacked_std['Contract Number'] == replaced_contract)].copy()
        replacement_infor_df = self.stacked_std[(self.stacked_std['Source System'] == 'Infor') &
                                                (self.stacked_std['Contract Number'] == replaced_contract) &
                                                (self.stacked_std['Active Rank'] == '1')][[check_mode, 'ItemType']].copy()
        tp_df = self.stacked_std[self.stacked_std['Source System'] == 'TP'].copy()
        on_infor_flag = True
        if len(replacement_ccx_df) == 0:
            print(f"Contract {replaced_contract} not found in CCX, please check the contract number and try again.")
            return Status.SUCCESS
        if len(replacement_infor_df) == 0:
            print(f"Contract {replaced_contract} not found in Infor, please check the contract number and try again.")
            on_infor_flag = False
        
        replacement_cols_to_take = ['Contract Number', 'MFN', 'VN', 'IN', 'Description', 'UnitCost', 'UOM', 'QOE', 
                                    'Effective Date', 'Expiration Date', 'seq']
        replaced_contract_mfn = set(replacement_ccx_df[check_mode])
        tp_contract_mfn = set(tp_df[check_mode])
        replaced_leftover = replaced_contract_mfn.difference(tp_contract_mfn)
        if len(replaced_leftover) == 0:
            print("full coverage using replacement contract, no leftover items found.")
            return Status.SUCCESS
        replacement_leftover_df = replacement_ccx_df[replacement_cols_to_take].copy()
        replacement_leftover_df = replacement_leftover_df[replacement_leftover_df[check_mode].isin(replaced_leftover)].copy()
        replacement_leftover_df.loc[:, 'On Replacement Contract'] = 'No'

        if on_infor_flag is True:
            replacement_leftover_df = replacement_leftover_df.merge(replacement_infor_df, on = [check_mode], how = 'left')
            replacement_leftover_df.loc[:, 'ItemType'] = replacement_leftover_df['ItemType'].fillna('Special')
        else:
            replacement_leftover_df.loc[:, 'ItemType'] = 'Special'

        if len(replacement_leftover_df) == 0:
            print("full coverage using replacement contract, no leftover items found.")
            return Status.SUCCESS
        else:
            replacement_leftover_df.sort_values(by = ['ItemType'], ascending = [True], inplace = True)
            replacement_leftover_df.to_excel(os.path.join(self.output_file_path,
                                                          f'replacement_leftover_{replaced_contract}_{self.datesig}.xlsx'),
                                                          index = False)
            print(f"replacement contract pair check completed, results are saved to output folder.")
        return Status.SUCCESS
    
    def make_ccx_upload_file(self):
        print('make_ccx_upload_file')
        return None
    
    def make_infor_upload_file_multiple(self):
        print('make_infor_upload_file_multiple')
        return None

