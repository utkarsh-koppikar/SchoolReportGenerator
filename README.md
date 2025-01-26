# Report Card Generator

## Project Overview
A Python-based application for generating report cards from Excel files using a Word template.

## Prerequisites
- Python 3.8+
- Windows Operating System (due to Win32 dependencies)

## Project Structure
```
report_generator/
│
├── input_files/
│   ├── template-word1A.docx
│   └── student_marks.xlsx
│
├── mappings/
│   ├── I_A_mapping.json
│   ├── II_B_mapping.json
│   └── ... (other class mappings)
│
├── requirements.txt
├── main.py
└── setup.py
```

## Setup Instructions

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/report-card-generator.git
cd report-card-generator
```

### 2. Create Virtual Environment
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

## Input File Preparations

### Excel File Requirements
- Located in `input_files/`
- Filename format: `{class_name}_marks.xlsx`
- Columns must match the corresponding mapping JSON

### Mapping JSON Files
- Located in `mappings/`
- Filename format: `{Class_Name}_mapping.json`
- Maps Excel column names to report card template fields

### Word Template
- Located in `input_files/template-word1A.docx`
- Must have placeholders matching the mapping JSON

## Configuration

### Mapping JSON Example
```json
{
    "name": "Student Name Column",
    "percentage": "Total Percentage Column",
    "remark": "Remarks Column"
}
```

## Running the Application

### Direct Python Execution
```bash
python main.py
```

### Build Executable

#### Using PyInstaller
```bash
# Install PyInstaller
pip install pyinstaller

# Create executable
pyinstaller --onefile --windowed --add-data "input_files;input_files" --add-data "mappings;mappings" main.py
```

## Troubleshooting

### Common Issues
- Ensure all required dependencies are installed
- Check Excel file format matches expected structure
- Verify mapping JSON is correctly configured

### Dependencies
Detailed dependencies are listed in `requirements.txt`. Key libraries include:
- pandas
- python-docx
- docxtpl
- pywin32
- tkinter

## License
[Your License Here - e.g., MIT]

## Contributing
1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Contact
[Your Contact Information]
```

## Additional Notes for Setup

### Requirements.txt Content
```
pandas==1.3.5
python-docx==0.8.11
docxtpl==0.16.4
pywin32==303
openpyxl==3.0.9
tk==0.1.0
```

Would you like me to elaborate on any section of the README or provide additional details about the project setup?
