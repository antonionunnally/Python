import os
import csv
import copy
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from typing import Dict, List, Optional, Any, Tuple

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("arw_processor.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("ARWProcessor")

# Dictionary mapping full state names to abbreviations
STATE_ABBREVIATIONS = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID",
    "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS",
    "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM", "new york": "NY",
    "north carolina": "NC", "north dakota": "ND", "ohio": "OH", "oklahoma": "OK",
    "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT",
    "vermont": "VT", "virginia": "VA", "washington": "WA", "west virginia": "WV",
    "wisconsin": "WI", "wyoming": "WY", "district of columbia": "DC"
}

# Fields that should be converted from "0" to empty string
ZERO_TO_EMPTY_FIELDS = [
    'Cancellation_Date', 'Cancel_Reason_Code', 'Business_Name', 'Customer_Address_2',
    'Customer_Phone', 'Customer_Email', 'Sales_Ticket_Number', 'Manufacturer_Name',
    'Model_Number', 'Model_Name', 'Serial_Number', 'Product_Condition',
    'Contract_Note', 'Renewal_Contract_Number', 'Change_Flag', 'Original_Contract_Number'
]

# Transaction reason codes that require Contract_Refund_Amount to be set to 0
ZERO_REFUND_REASON_CODES = ['1', '2', '5']


class FileHandler:
    """Handles file operations for reading and writing CSV files."""
    
    @staticmethod
    def read_csv(file_path: str) -> List[Dict[str, str]]:
        """
        Read and load CSV file data.
        
        Args:
            file_path: Path to the CSV file
            
        Returns:
            List of dictionaries representing CSV records
            
        Raises:
            FileNotFoundError: If the file doesn't exist
            Exception: For other reading errors
        """
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
                
            records = []
            with open(file_path, 'r', encoding='utf-8-sig') as data_file:
                for record in csv.DictReader(data_file, delimiter=",", quoting=csv.QUOTE_MINIMAL):
                    records.append(copy.deepcopy(record))
            
            logger.info(f"Successfully read {len(records)} records from {file_path}")
            return records
            
        except FileNotFoundError as e:
            logger.error(f"File not found: {file_path}")
            raise
        except Exception as e:
            logger.error(f"Error reading file: {str(e)}")
            raise Exception(f"Error reading file: {str(e)}")
    
    @staticmethod
    def write_csv(file_path: str, records: List[Dict[str, str]]) -> None:
        """
        Write processed records to CSV file.
        
        Args:
            file_path: Path to write the output file
            records: List of dictionaries to write
            
        Raises:
            Exception: If writing fails
        """
        if not records:
            logger.warning("No records to write")
            raise Exception("No records to write to file")
            
        try:
            keys = records[0].keys()
            with open(file_path, 'w', newline='', encoding='utf-8') as output_file:
                dict_writer = csv.DictWriter(output_file, keys, delimiter=",", quoting=csv.QUOTE_MINIMAL)
                dict_writer.writeheader()
                dict_writer.writerows(records)
            
            logger.info(f"Successfully wrote {len(records)} records to {file_path}")
        except Exception as e:
            logger.error(f"Error writing output file: {str(e)}")
            raise Exception(f"Error writing output file: {str(e)}")


class RecordProcessor:
    """Responsible for all record transformation operations."""
    
    @staticmethod
    def convert_zeros_to_empty(record: Dict[str, str], fields: List[str]) -> Dict[str, str]:
        """
        Replace "0" values with empty strings for specified fields.
        
        Args:
            record: Dictionary containing record data
            fields: List of field names to process
            
        Returns:
            Processed record
        """
        for field in fields:
            if field in record and record[field] == "0":
                record[field] = ""
        return record
    
    @staticmethod
    def standardize_state(record: Dict[str, str]) -> Dict[str, str]:
        """
        Capitalize state values and convert to standard abbreviations.
        
        Args:
            record: Dictionary containing record data
            
        Returns:
            Record with standardized state field
        """
        if 'Customer_State' in record and record['Customer_State']:
            state_value = record['Customer_State'].lower().strip()
            
            # Check if already a valid abbreviation (2 uppercase letters)
            if len(state_value) == 2 and state_value.upper() == state_value:
                record['Customer_State'] = state_value.upper()
            # Check if it's a full state name that needs conversion
            elif state_value in STATE_ABBREVIATIONS:
                record['Customer_State'] = STATE_ABBREVIATIONS[state_value]
            # Otherwise capitalize whatever is there
            else:
                record['Customer_State'] = state_value.upper()
                
        return record
    
    @staticmethod
    def fix_contract_price(record: Dict[str, str]) -> Dict[str, str]:
        """
        Change 0 values in Contract_Price_Retail_Cost to 1.
        
        Args:
            record: Dictionary containing record data
            
        Returns:
            Record with fixed contract price
        """
        if 'Contract_Price_Retail_Cost' in record:
            try:
                # Handle various zero representations (0, 0.0, etc.)
                price_value = float(record['Contract_Price_Retail_Cost'])
                if price_value == 0:
                    record['Contract_Price_Retail_Cost'] = "1"
            except (ValueError, TypeError):
                # If value can't be converted to float, leave unchanged
                pass
                
        return record
    
    @staticmethod
    def handle_refund_amounts(record: Dict[str, str]) -> Dict[str, str]:
        """
        Set Contract_Refund_Amount to 0 for Transaction_Reason values 1, 2, or 5.
        
        Args:
            record: Dictionary containing record data
            
        Returns:
            Record with properly handled refund amounts
        """
        if ('Transaction_Reason' in record and 
            'Contract_Refund_Amount' in record and 
            record['Transaction_Reason'] in ZERO_REFUND_REASON_CODES):
            record['Contract_Refund_Amount'] = "0"
                
        return record


class ARWFileProcessor:
    """Main processor class for ARW files."""
    
    def __init__(self):
        """Initialize the processor with empty data structures."""
        self.file_handler = FileHandler()
        self.record_processor = RecordProcessor()
        self.input_file_path: Optional[str] = None
        self.output_file_path: Optional[str] = None
        self.data: List[Dict[str, str]] = []
        self.processed_records: List[Dict[str, str]] = []
        
    def read_file(self, csv_file_path: str) -> int:
        """
        Read and load the CSV file data.
        
        Args:
            csv_file_path: Path to the CSV file
            
        Returns:
            Number of records read
            
        Raises:
            Exception: If file reading fails
        """
        self.input_file_path = csv_file_path
        self.data = self.file_handler.read_csv(csv_file_path)
        return len(self.data)
            
    def process(self) -> int:
        """
        Process the data with all required transformations.
        
        Returns:
            Number of processed records
            
        Raises:
            Exception: If no data loaded or processing fails
        """
        if not self.data:
            logger.error("No data loaded for processing")
            raise Exception("No data loaded. Please select a file first.")
            
        try:
            self.processed_records = []
            
            for record in self.data:
                processed_record = copy.deepcopy(record)
                
                # Apply all transformations
                processed_record = self.record_processor.convert_zeros_to_empty(
                    processed_record, ZERO_TO_EMPTY_FIELDS)
                processed_record = self.record_processor.standardize_state(processed_record)
                processed_record = self.record_processor.fix_contract_price(processed_record)
                processed_record = self.record_processor.handle_refund_amounts(processed_record)
                
                self.processed_records.append(processed_record)
            
            logger.info(f"Successfully processed {len(self.processed_records)} records")
            return len(self.processed_records)
            
        except Exception as e:
            logger.error(f"Error during processing: {str(e)}")
            raise Exception(f"Error during processing: {str(e)}")
        
    def write_output_file(self, output_path: str) -> bool:
        """
        Write processed data to the output file.
        
        Args:
            output_path: Path to write the output file
            
        Returns:
            True if successful
            
        Raises:
            Exception: If no processed data or writing fails
        """
        if not self.processed_records:
            logger.error("No processed data available for writing")
            raise Exception("No processed data available. Please process the data first.")
            
        self.output_file_path = output_path
        self.file_handler.write_csv(output_path, self.processed_records)
        return True


class ARWFileProcessorGUI:
    """GUI for the ARW File Processor application."""
    
    def __init__(self, root):
        """
        Initialize the GUI.
        
        Args:
            root: Root Tkinter window
        """
        self.root = root
        self.processor = ARWFileProcessor()
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the GUI components."""
        self.root.title("ARW File Processor")
        self.root.geometry("800x600")
        
        # Configure grid
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(4, weight=1)
        
        # Create frames
        self.input_frame = ctk.CTkFrame(self.root)
        self.input_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        
        self.output_frame = ctk.CTkFrame(self.root)
        self.output_frame.grid(row=1, column=0, padx=20, pady=20, sticky="ew")
        
        self.button_frame = ctk.CTkFrame(self.root)
        self.button_frame.grid(row=2, column=0, padx=20, pady=20, sticky="ew")
        
        self.status_frame = ctk.CTkFrame(self.root)
        self.status_frame.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        
        self.log_frame = ctk.CTkFrame(self.root)
        self.log_frame.grid(row=4, column=0, padx=20, pady=20, sticky="nsew")
        
        # Configure frames grid
        for frame in [self.input_frame, self.output_frame, self.button_frame, self.status_frame, self.log_frame]:
            frame.grid_columnconfigure(1, weight=1)
        
        # Input file section
        ctk.CTkLabel(self.input_frame, text="Input File:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.input_path_var = tk.StringVar()
        self.input_entry = ctk.CTkEntry(self.input_frame, textvariable=self.input_path_var, width=500)
        self.input_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(self.input_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=2, padx=10, pady=10)
        
        # Output file section
        ctk.CTkLabel(self.output_frame, text="Output File:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.output_path_var = tk.StringVar()
        self.output_entry = ctk.CTkEntry(self.output_frame, textvariable=self.output_path_var, width=500)
        self.output_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(self.output_frame, text="Browse", command=self.browse_output_file).grid(row=0, column=2, padx=10, pady=10)
        
        # Buttons
        self.process_button = ctk.CTkButton(self.button_frame, text="Process File", command=self.process_file)
        self.process_button.grid(row=0, column=0, padx=10, pady=10)
        
        self.clear_button = ctk.CTkButton(self.button_frame, text="Clear All", command=self.clear_all)
        self.clear_button.grid(row=0, column=1, padx=10, pady=10)

        # Updated button text and function
        self.file_location_button = ctk.CTkButton(
            self.button_frame, 
            text="Processed File Location", 
            command=self.open_file_location,
            #tooltip="Opens the directory containing the processed file"
        )
        self.file_location_button.grid(row=0, column=2, padx=10, pady=10)
        
        # Status
        ctk.CTkLabel(self.status_frame, text="Status:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.status_var = tk.StringVar(value="Ready")
        ctk.CTkLabel(self.status_frame, textvariable=self.status_var).grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # Log
        ctk.CTkLabel(self.log_frame, text="Log:").grid(row=0, column=0, padx=10, pady=10, sticky="nw")
        self.log_text = ctk.CTkTextbox(self.log_frame, width=700, height=200)
        self.log_text.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        
    def browse_input_file(self):
        """Open file dialog to select input CSV file."""
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if not file_path:
            return
            
        self.input_path_var.set(file_path)
        self.log(f"Selected input file: {file_path}")
        
        # Automatically set output file name based on input file
        dir_name, file_name = os.path.split(file_path)
        base_name, ext = os.path.splitext(file_name)
        output_file = os.path.join(dir_name, f"{base_name}_Fix{ext}")
        self.output_path_var.set(output_file)
        self.log(f"Default output file set to: {output_file}")
            
    def browse_output_file(self):
        """Open file dialog to select output CSV file."""
        file_path = filedialog.asksaveasfilename(
            title="Save Output File",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_path_var.set(file_path)
            self.log(f"Selected output file: {file_path}")
            
    def process_file(self):
        """Process the input file and save the results."""
        input_path = self.input_path_var.get()
        output_path = self.output_path_var.get()
        
        if not input_path:
            messagebox.showerror("Error", "Please select an input file")
            return
            
        if not output_path:
            messagebox.showerror("Error", "Please select an output file")
            return
            
        try:
            self.status_var.set("Processing...")
            self.log("Starting file processing...")
            
            # Read input file
            record_count = self.processor.read_file(input_path)
            self.log(f"Read {record_count} records from input file")
            
            # Process data
            processed_count = self.processor.process()
            self.log(f"Processed {processed_count} records")
            
            # Write output file
            self.processor.write_output_file(output_path)
            self.log(f"Successfully wrote output to: {output_path}")
            
            self.status_var.set("Completed")
            messagebox.showinfo("Success", f"File processed successfully. {processed_count} records written to output file.")
            
        except Exception as e:
            self.status_var.set("Error")
            self.log(f"ERROR: {str(e)}")
            messagebox.showerror("Error", str(e))

    def open_file_location(self):
        """Open the directory containing the processed output file using the default file explorer."""
        output_path = self.output_path_var.get()
        
        if not output_path:
            messagebox.showerror("Error", "No output file path specified")
            return
            
        # Get the directory path
        dir_path = os.path.dirname(os.path.abspath(output_path))
        
        if not os.path.exists(dir_path):
            messagebox.showerror("Error", "Directory does not exist")
            return
            
        try:
            # Open file explorer to the directory
            if os.name == 'nt':  # Windows
                os.startfile(dir_path)
            elif os.name == 'posix':  # macOS, Linux
                import subprocess
                if os.path.exists('/usr/bin/xdg-open'):  # Linux
                    subprocess.call(['xdg-open', dir_path])
                else:  # macOS
                    subprocess.call(['open', dir_path])
            self.log(f"Opened directory containing output file: {dir_path}")
        except Exception as e:
            self.log(f"Error opening directory: {str(e)}")
            messagebox.showerror("Error", f"Could not open directory: {str(e)}")
            
    def clear_all(self):
        """Clear all entries and reset the processor."""
        self.input_path_var.set("")
        self.output_path_var.set("")
        self.status_var.set("Ready")
        self.log_text.delete("1.0", tk.END)
        self.processor = ARWFileProcessor()
        self.log("All fields cleared and reset")
        
    def log(self, message):
        """Add a message to the log text box."""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        logger.info(message)


def main():
    """Main entry point for the application."""
    # Set appearance mode and default color theme
    ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
    ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
    
    root = ctk.CTk()
    app = ARWFileProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()