import os
import sys

def delete_excel_files(file_paths):
    for file_path in file_paths:
        try:
            # Check if the file exists
            if os.path.exists(file_path):
                # Delete the file
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
            else:
                print(f"File not found: {file_path}")
        except Exception as e:
            print(f"Error deleting file {file_path}: {e}")

def main_delete():   
    if getattr(sys, 'frozen', False):
        # When running as a bundled executable (e.g., PyInstaller)
        script_dir = os.path.dirname(sys.executable)
    else:
        # When running as a script
        script_dir = os.path.dirname(os.path.abspath(__file__))

    files_to_delete = [
        os.path.join(script_dir, "consolidated_&_formatted_data (new).xlsx"),
        os.path.join(script_dir, "formatted_sales_dashboard (new).xlsx"),
        os.path.join(script_dir, "sales_dashboard (new).xlsx"),
        os.path.join(script_dir, "consignment_data.xlsx"),
        os.path.join(script_dir, "new_data.xlsx"),
        os.path.join(script_dir, "quotation_data.xlsx"),
        os.path.join(script_dir, "scrapexport_data.xlsx"),
        os.path.join(script_dir, "sold_data.xlsx"),
        os.path.join(script_dir, "void_data.xlsx")
    ]
    
    delete_excel_files(files_to_delete)

if __name__ == "__main__":
    main_delete()
