import os
import sys
import shutil
import hashlib
import traceback
from typing import Any
from pandas import DataFrame
from xml.sax.saxutils import escape as xml_sax_escape
from datetime import datetime,timedelta
from xml.etree import ElementTree as ET

class XML_Logger:
    def __init__(self,log_file:str,archive_folder="archive",log_retention_days=30,base_dir=os.path.dirname(os.path.abspath(__file__))):
        self.log_file = log_file
        self.archive_folder = archive_folder
        self.log_retention_days = log_retention_days
        self.base_dir = base_dir

    def save_variable_info(self,globals_dict:dict[str,Any],locals_dict:dict[str,Any],variable_save_path:str) -> None:
        try:
            # Combine them, prioritizing locals (to avoid duplicates)
            all_vars:dict[str,Any] = {**globals_dict, **locals_dict}
            
            # Filter out modules, functions, and built-ins
            variable_info:list[dict[str,str|int|float|list|set|dict|bytes]] = []
            for name, value in all_vars.items():
                # Skip special variables, modules, and callables
                if name.startswith('__') and name.endswith('__'):
                    continue
                if callable(value):
                    continue
                if isinstance(value, type(sys)):  # Skip modules
                    continue
                    
                # Get variable details
                var_type:str = type(value).__name__
                try:
                    var_hash:str = hashlib.sha256(str(value).encode('utf-8')).hexdigest()
                except Exception:
                    var_hash:str = "Unhashable"
                
                var_size:int = sys.getsizeof(value)
                
                variable_info.append({
                    "Variable Name": name,
                    "Type": var_type,
                    "Hash": var_hash,
                    "Size (bytes)": var_size
                })
            
            # Convert to a DataFrame for nice tabular output
            df:DataFrame = DataFrame(variable_info)
            df.to_json(os.path.join(self.base_dir,variable_save_path),orient='table',indent=4)
        except Exception as e:
            traceback.print_exc()
            print(str(e))

    def get_current_log_filename(self,basepath:str) -> str:
        """Generates a log filename based on the current date."""
        return f"{basepath}/{self.log_file}_{datetime.now().strftime('%Y%m%d')}.xml"

    def rotate_logs(self):
        """Checks if the log file date has changed and archives it if needed."""
        current_date = datetime.now().strftime("%Y%m%d")
        
        # Get the last modified date of the existing log file
        if os.path.exists(self.log_file):
            modified_time = datetime.fromtimestamp(os.path.getmtime(self.log_file)).strftime("%Y%m%d")

            if modified_time != current_date:  # If the log is from a previous day, archive it
                # Ensure archive folder exists
                if not os.path.exists(self.archive_folder):
                    os.makedirs(self.archive_folder)

                # Move the old log file to archive with a date-based name
                archive_filename = f"{self.archive_folder}/{self.log_file}_{modified_time}.xml"
                shutil.move(self.log_file, archive_filename)

                # Perform cleanup of old logs
                self.delete_old_logs()

    def delete_old_logs(self):
        """Deletes logs that are older than LOG_RETENTION_DAYS."""
        cutoff_date:datetime = datetime.now() - timedelta(days=self.log_retention_days)

        if not os.path.exists(self.archive_folder):
            return  # No logs to delete

        for filename in os.listdir(self.archive_folder):
            if filename.startswith(f"{self.log_file}") and filename.endswith(".xml"):
                try:
                    # Extract date from filename
                    date_str:str = filename.replace(f"{self.log_file}", "").replace(".xml", "")
                    log_date:datetime = datetime.strptime(date_str, "%Y%m%d")

                    # Delete files older than retention period
                    if log_date < cutoff_date:
                        file_path:str = os.path.join(self.archive_folder, filename)
                        os.remove(file_path)

                except ValueError:
                    # Ignore files that don't match the expected date format
                    pass

    def log_to_xml(self, message:str, basepath:str, status="INFO"):
        """
        Logs a message to an XML file, ensuring daily log rotation and old log cleanup.
        """
        self.rotate_logs()  # Check if the date has changed and archive if necessary

        # Get the correct filename for today's log
        current_log_file = self.get_current_log_filename(basepath=basepath)

        # Create file if it does not exist
        if not os.path.exists(current_log_file):
            root = ET.Element("logs")
            tree = ET.ElementTree(root)
            tree.write(current_log_file)

        # Load existing XML file
        tree = ET.parse(current_log_file)
        root = tree.getroot()

        # Create log entry
        log_entry = ET.SubElement(root, "log")
        log_entry.set("timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f"))
        log_entry.set("status", status)

        message_element = ET.SubElement(log_entry, "message")
        message_element.text = xml_sax_escape(message)

        # Write changes back to the file
        tree.write(current_log_file)

    def __str__(self):
        return f"XML Logger saves to {os.path.join(self.base_dir,self.log_file)}. Archives to {self.archive_folder}."
    
    def __repr__(self):
        return f"Class XML_Logger, log_file: {self.log_file}, archive_folder: {self.archive_folder}, log_retention_days: {self.log_retention_days}, base_dir: {self.base_dir}"