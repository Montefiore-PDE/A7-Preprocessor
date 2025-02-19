import tkinter as tk
from tkinter import messagebox
from FileProcessor import FileProcessor
from FolderManager import FolderManager
from TypesDefinition import CheckMode, ProcessType

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Pre-processor v1.0")

        # Manufacturer Name
        tk.Label(root, text="Manufacturer Name:").grid(row=0, column=0, padx=10, pady=10)
        self.manufacturer_name_entry = tk.Entry(root)
        self.manufacturer_name_entry.grid(row=0, column=1, padx=10, pady=10)

        # Contract Name
        tk.Label(root, text="Contract Name:").grid(row=1, column=0, padx=10, pady=10)
        self.contract_name_entry = tk.Entry(root)
        self.contract_name_entry.grid(row=1, column=1, padx=10, pady=10)

        # Process Type
        tk.Label(root, text="Process Type:").grid(row=2, column=0, padx=10, pady=10)
        self.process_type_var = tk.StringVar(root)
        self.process_type_var.set("full_process")  # default value
        process_types = ["pre_check", "scoping", "standardize_all_and_stack", "dup_search_and_compare", 
                         "itemmast_search_and_compare", "replacement_contract_pair_check", 
                         "ccx_dup_search_and_itemmast_match", "full_process"]
        self.process_type_menu = tk.OptionMenu(root, self.process_type_var, *process_types)
        self.process_type_menu.grid(row=2, column=1, padx=10, pady=10)

        # Run Button
        self.run_button = tk.Button(root, text="Run", command=self.run_process)
        self.run_button.grid(row=3, column=0, columnspan=2, pady=20)

    def run_process(self):
        manufacturer_name = self.manufacturer_name_entry.get().strip()
        contract_name = self.contract_name_entry.get().strip()
        process_type_str = self.process_type_var.get()

        if not manufacturer_name or not contract_name:
            messagebox.showerror("Input Error", "Please enter both manufacturer name and contract name.")
            return

        folder_manager = FolderManager(manufacturer_name, contract_name)
        folder_manager.create_folders()

        process_type_map = {
            "pre_check": ProcessType.pre_check,
            "scoping": ProcessType.scoping,
            "standardize_all_and_stack": ProcessType.standardize_all_and_stack,
            "dup_search_and_compare": ProcessType.dup_search_and_compare,
            "itemmast_search_and_compare": ProcessType.itemmast_search_and_compare,
            "replacement_contract_pair_check": ProcessType.replacement_contract_pair_check,
            "ccx_dup_search_and_itemmast_match": ProcessType.ccx_dup_search_and_itemmast_match,
            "full_process": ProcessType.full_process
        }

        process_type = process_type_map.get(process_type_str)
        if process_type is None:
            messagebox.showerror("Process Error", "Invalid process type selected.")
            return

        file_processor = FileProcessor(folder_manager)
        file_processor.process_files(process_type)
        messagebox.showinfo("Success", f"Process '{process_type_str}' completed successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    print("Initiating .......")
    app = App(root)
    root.mainloop()