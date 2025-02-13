class ProcessType:
    pre_check = "pre_check"
    scoping = "scoping"
    standardize_all_and_stack = "standardize_all_and_stack"
    dup_search_and_compare = "dup_search"
    itemmast_search_and_compare = "itemmast_search"
    ccx_dup_search_and_itemmast_match = "ccx_dup_search_and_itemmast_match"
    residue_distribution = "residue_item_redistribution"
    replacement_contract_pair_check = "replacement_contractS_pair_check"
    dissolve = "dissolve"
    post_check = "post_check"

class CheckMode:
    MFN = "MFN"
    MFN_RF = "MFN RF"

class StandardizeTarget:
    INFOR = "Infor"
    IMPORT = "Import"
    CCX = "CCX"
    TP = "TP" # to pre-process