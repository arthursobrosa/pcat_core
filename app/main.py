from modules import process_workbooks, process_data_base, merge_last_dbs
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

if __name__ == "__main__":  
    # process_workbooks("Permissionária")
    # process_data_base("Concessionária")
    merge_last_dbs()
