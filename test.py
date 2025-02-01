import traceback
from ReportGenV2 import *
import sys

def read_config(file_path):
    config = {}
    with open(file_path, 'r') as file:
        for line in file:
            key, value = line.strip().split('=')
            config[key] = value
    return config

def main(arg):
    try:
        config = read_config('test_config.txt')
        excel_path = config['excel_path']
        word_path = config['word_path']
        mapping_path = config['mapping_path']
        class_name = config['class_name']

        generator = ReportCardGenerator()
        if arg is None or arg == "run":
            generator.generate_report_cards(excel_path, word_path, mapping_path, class_name)
            print("Report generation successful")
        elif arg == "read_excel":
            data_processor = DataProcessor(excel_path, mapping_path)
            df = data_processor._load_dataframe(excel_path)
            print(df.to_markdown())
        else:
            print("No such test")

    except Exception as e:
        print(f"Error generating report cards: {e}")
        traceback.print_exc()

main(sys.argv[1])