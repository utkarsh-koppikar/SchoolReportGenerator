import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import time
import win32com.client
from docxtpl import DocxTemplate
from docx2pdf import convert
import traceback

class FileManager:
    """Handles file and directory operations for the report card generator."""
    
    @staticmethod
    def ensure_directory_exists(directory):
        """Create directory if it doesn't exist."""
        if not os.path.exists(directory):
            os.makedirs(directory)
    
    @staticmethod
    def get_absolute_path(base_dir, *path_parts):
        """Construct an absolute path from base directory and path parts."""
        return os.path.join(base_dir, *path_parts)

class DataProcessor:
    """Processes Excel data for report card generation."""
    
    def __init__(self, excel_path, mapping_path):
        """
        Initialize data processor with Excel file and mapping.
        
        :param excel_path: Path to the Excel file
        :param mapping_path: Path to the JSON mapping file
        """
        self.df = self._load_dataframe(excel_path)
        self.column_map = self._load_mapping(mapping_path)
    
    def _load_dataframe(self, excel_path):
        """
        Load and preprocess Excel dataframe.
        
        :param excel_path: Path to the Excel file
        :return: Processed DataFrame
        """
        try:
            df = pd.read_excel(excel_path)
            return df[0:]
        except Exception as e:
            raise ValueError(f"Error loading Excel file: {e}")
    
    def _load_mapping(self, mapping_path):
        """
        Load column mapping from JSON file.
        
        :param mapping_path: Path to the mapping JSON
        :return: Column mapping dictionary
        """
        with open(mapping_path, "r") as file:
            return dict(json.load(file))
    
    def process_student_data(self, row, class_name):
        """
        Process individual student data for report card generation.
        
        :param row: DataFrame row for a student
        :param class_name: Name of the class
        :return: Processed student data dictionary
        """
        try:
            print(f"Processing student data for class: {class_name}")
            print(f"Row data: {dict(row)}")
            print(f"Column map: {self.column_map}")

            field_dict = {}
            print(f"Initial field_dict: {field_dict}")
            for key in self.column_map.keys():
                if key not in ['percentage', 'remark', 'class']:
                    field_dict[key] = row[self.column_map[key]]
                elif key == 'percentage':
                    field_dict['percentage'] = f"{float(row[self.column_map['percentage']]):.2f}%"
                elif key == 'remark':
                    field_dict['remark'] = row[self.column_map['remark']] + "!"
            
            field_dict['class'] = class_name.replace("_", " ")
            
            # Handle null/empty values
            nan_values = ['NAN', 'NONE', 'NA']
            field_dict = {
                key: value if value is not None and 
                str(value).upper().replace(" ", "") not in nan_values 
                else "---" 
                for key, value in field_dict.items()
            }
            
            print(f"Processed field dictionary: {field_dict}")
            return field_dict

        except Exception as e:
            print(f"Error processing student data: {e}")
            traceback.print_exc()
            return None

class ReportCardGenerator:
    """Manages the generation of report cards."""
    
    def __init__(self):
        """
        Initialize report card generator.
        
        :param base_directory: Base directory for project files
        """
        self.input_folder = 'input_files'
        self.mappings_folder = 'mappings'
    
    def generate_report_cards(self, excel_path, template_path, mapping_path, class_name):
        """
        Generate report cards for a given class.
        
        :param excel_filename: Name of the Excel file
        :param class_name: Name of the class
        """
        print(f"excel_path: {excel_path}")
        print(f"template_path: {template_path}")
        print(f"mapping_path: {mapping_path}")

        # Setup directories
        FileManager.ensure_directory_exists('word')
        report_cards_dir = f'{class_name} report_cards'
        FileManager.ensure_directory_exists(report_cards_dir)
        
        # Process data
        data_processor = DataProcessor(excel_path, mapping_path)
        
        print(f"data_processor: {data_processor}")

        print(f"{data_processor.df.head()}")
        # Generate Word documents
        for index,row in data_processor.df.iterrows():
            print(f"row: {row}")
            student_data = data_processor.process_student_data(
                row, class_name
            )
            
            if student_data['name'] == '---':
                break
            
            self._create_word_document(
                template_path, student_data, 'word'
            )
        
        # Convert to PDF
        self._convert_to_pdf(data_processor.df, data_processor.column_map, 
                              class_name, report_cards_dir)
    
    def _create_word_document(self, template_path, student_data, output_dir):
        """
        Create a Word document for a student.
        
        :param template_path: Path to the template document
        :param student_data: Processed student data
        :param output_dir: Output directory for Word files
        """
        try:
            template = DocxTemplate(template_path)
            template.render(student_data)
            
            filename = f"{student_data['name']}.docx"
            filled_path = os.path.join(os.getcwd(), output_dir, filename)
            template.save(filled_path)
        except Exception as e:
            print(f"Error creating document: {e}")
            template.save("dummy.docx")
    
    def _convert_to_pdf(self, df, column_map, class_name, output_dir):
        """
        Convert Word documents to PDF.
        
        :param df: DataFrame with student data
        :param column_map: Column mapping dictionary
        :param class_name: Name of the class
        :param output_dir: Output directory for PDF files
        """
        for row in range(len(df)):
            student = df.iloc[row][column_map['name']]
            word_filename = f'{student}.docx'
            word_path = os.path.join(os.getcwd(), 'word', word_filename)
            pdf_path = os.path.join(os.getcwd(), output_dir, f'{student}.pdf')
            
            print(f'Generating Report card for {student}')
            convert(word_path, pdf_path)

class ReportCardGeneratorApp:
    """Tkinter GUI for Report Card Generator."""
    
    def __init__(self, root):
        """
        Initialize the application GUI.
        
        :param root: Tkinter root window
        """
        self.root = root
        self.root.title("Report Card Generator")
        
        # Base project directory (modify as needed)
        self.base_directory = os.path.dirname(__file__)
        
        # Variables
        self.excel_file_path = tk.StringVar()
        self.word_file_path = tk.StringVar()
        self.mapping_file_path = tk.StringVar()
        self.class_name = tk.StringVar()
        
        # Create UI
        self._create_ui()
    def _create_ui(self):
        """Create the user interface components."""
        notebook = ttk.Notebook(self.root)
        tab_run_script = ttk.Frame(notebook)
        notebook.add(tab_run_script, text="Run Script")
        notebook.pack(expand=1, fill="both")
        
        # File Selection Frame
        file_frame = ttk.LabelFrame(tab_run_script, text="File Selection", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        # Excel File Selection
        ttk.Label(file_frame, text="Excel File:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(
            file_frame, 
            textvariable=self.excel_file_path, 
            state="readonly",
            width=50
        ).grid(row=0, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(
            file_frame, 
            text="Browse", 
            command=lambda: self._choose_file([("Excel files", "*.xlsx")],self.excel_file_path)
        ).grid(row=0, column=2, padx=5, pady=5)
        
        # Word File Selection
        ttk.Label(file_frame, text="Word Template:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(
            file_frame, 
            textvariable=self.word_file_path, 
            state="readonly",
            width=50
        ).grid(row=1, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(
            file_frame, 
            text="Browse", 
            command=lambda: self._choose_file([("Word files", "*.docx")],self.word_file_path)
        ).grid(row=1, column=2, padx=5, pady=5)

        # Mapping File Selection
        ttk.Label(file_frame, text="Mapping File:").grid(
            row=2, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(
            file_frame, 
            textvariable=self.mapping_file_path, 
            state="readonly",
            width=50
        ).grid(row=2, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(
            file_frame, 
            text="Browse", 
            command=lambda: self._choose_file([("JSON files", "*.json")],self.mapping_file_path)
        ).grid(row=2, column=2, padx=5, pady=5)

        print
        # Configure grid column weights
        file_frame.columnconfigure(1, weight=1)

        # Class Name Frame
        class_frame = ttk.LabelFrame(tab_run_script, text="Class Details", padding=10)
        class_frame.pack(fill="x", padx=10, pady=5)
        
        # Class Name Entry
        ttk.Label(class_frame, text="Class Name:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(
            class_frame, 
            textvariable=self.class_name,
            width=50
        ).grid(row=0, column=1, padx=5, pady=5, sticky="we")
        
        # Configure grid column weights for class frame
        class_frame.columnconfigure(1, weight=1)
        
        # Run Button Frame
        button_frame = ttk.Frame(tab_run_script)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        # Run Script Button
        ttk.Button(
            button_frame, 
            text="Run Script",
            command=self._run_script,
            style="Accent.TButton"  # Add this if you want to style the button differently
        ).pack(pady=5)

    def _choose_file(self, file_types,target_var):
        """
        Open file dialog to choose files based on specified types.
        Args:
            file_types: List of tuples with file type descriptions and extensions.
                    Example: [("Excel files", "*.xlsx"), ("JSON files", "*.json")]
        """
        if file_types is None:
            file_types = [
                ("Excel files", "*.xlsx"),
                ("Word files", "*.docx"),
                ("JSON files", "*.json")
            ]
        
        file_path = filedialog.askopenfilename(filetypes=file_types)
        if file_path:
            target_var.set(file_path)
    
    def _run_script(self):
        """Execute the report card generation script."""
        excel_path = self.excel_file_path.get()
        word_path = self.word_file_path.get()
        mapping_path = self.mapping_file_path.get()
        class_name = self.class_name.get()
        
        if not excel_path or not word_path or not mapping_path or not class_name:
            messagebox.showerror("Error", "Please select files and enter class name")
            return
        
        try:
            generator = ReportCardGenerator()
            generator.generate_report_cards(excel_path,word_path,mapping_path, class_name)
            messagebox.showinfo("Success", "Report cards generated successfully!")

        except Exception as e:
            print("Error occurred:")
            traceback.print_exc()  # Print full stack trace
            messagebox.showerror("Error", str(e))

def quit_word_application():
    """Quit Microsoft Word application if open."""
    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Quit()
        print("Microsoft Word application quit successfully.")
    except Exception as e:
        print(f"Error: {e}")

def main():
    """Main entry point for the application."""
    root = tk.Tk()
    ReportCardGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()