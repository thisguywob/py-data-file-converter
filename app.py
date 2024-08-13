import pandas as pd
import json
import os
import time
from tqdm import tqdm
from termcolor import colored, cprint
import tkinter as tk
from tkinter import filedialog

# Global Variables
INPUT_PATH = None
OUTPUT_DIR = None
FILETYPES = None
EXTENSION = None

def print_banner():
    banner = r"""
███╗   ███╗███████╗██████╗ ██╗ ██████╗ █████╗ ██████╗ ██████╗ 
████╗ ████║██╔════╝██╔══██╗██║██╔════╝██╔══██╗██╔══██╗██╔══██╗
██╔████╔██║█████╗  ██║  ██║██║██║     ███████║██████╔╝██║  ██║
██║╚██╔╝██║██╔══╝  ██║  ██║██║██║     ██╔══██║██╔══██╗██║  ██║
██║ ╚═╝ ██║███████╗██████╔╝██║╚██████╗██║  ██║██║  ██║██████╔╝
╚═╝     ╚═╝╚══════╝╚═════╝ ╚═╝ ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝ 
 ___  ___  _     ___    ___  ___   _  _ __   __ ___  ___  _____  ___  ___ 
| __||_ _|| |   | __|  / __|/ _ \ | \| |\ \ / /| __|| _ \|_   _|| __|| _ \
| _|  | | | |__ | _|  | (__| (_) || .` | \ V / | _| |   /  | |  | _| |   /
|_|  |___||____||___|  \___|\___/ |_|\_|  \_/  |___||_|_\  |_|  |___||_|_\ 

Developed by: Application Maintenance and Support L3
    """
    cprint(banner, 'cyan')

def print_menu():
    cprint("========== MENU ==========", 'yellow', attrs=['bold'])
    
    for key, item in PROGRAM_MENU.items():
        print(colored(f"[{key}]", item['color']), item['label'])

    cprint("==========================", 'yellow', attrs=['bold'])

def show_error(message):
    cprint("\n[ERROR] ", 'red', attrs=['bold'], end='')
    cprint(message, 'red')

def show_progress(desc, total_steps=100):
    with tqdm(total=total_steps, desc=desc, bar_format="{l_bar}{bar} [Time elapsed: {elapsed}]", ncols=100) as pbar:
        for i in range(total_steps):
            time.sleep(0.02)
            pbar.update(1)

def trim_multiline_string(cell):
    if isinstance(cell, str):
        return '\n'.join(line.strip() for line in cell.splitlines())
    return cell

def load_json(file_path):
    with open(file_path, encoding='utf-8') as json_file:
        return json.load(json_file)

def save_excel(df, file_path):
    df.to_excel(file_path, index=False, engine='openpyxl')

def save_csv(df, file_path):
    df.to_csv(file_path, index=False, encoding='utf-8')

def load_excel(file_path):
    return pd.read_excel(file_path, engine='openpyxl')

def load_csv(file_path):
    return pd.read_csv(file_path, encoding='utf-8')

def save_json(df, file_path):
    with open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(df.to_dict(orient='records'), json_file, indent=4, ensure_ascii=False)

def convert_json_to_xlsx(input_path, output_path):
    data = load_json(input_path)
    df = pd.json_normalize(data)
    df = df.apply(lambda col: col.map(trim_multiline_string))
    save_excel(df, output_path)

def convert_json_to_csv(input_path, output_path):
    data = load_json(input_path)
    df = pd.json_normalize(data)
    df = df.apply(lambda col: col.map(trim_multiline_string))
    save_csv(df, output_path)

def convert_xlsx_to_json(input_path, output_path):
    df = load_excel(input_path)
    save_json(df, output_path)

def convert_xlsx_to_csv(input_path, output_path):
    df = load_excel(input_path)
    save_csv(df, output_path)

def convert_csv_to_json(input_path, output_path):
    df = load_csv(input_path)
    save_json(df, output_path)

def convert_csv_to_xlsx(input_path, output_path):
    df = load_csv(input_path)
    save_excel(df, output_path)

def select_file(title, fileTypes):
    try:
        root = tk.Tk()
        root.withdraw()
        global INPUT_PATH
        INPUT_PATH = filedialog.askopenfilename(filetypes=fileTypes, title=title)
        return INPUT_PATH
    except Exception as e:
        show_error(f"Can't select file: {e}\n")

def select_directory(title):
    root = tk.Tk()
    root.withdraw()
    global OUTPUT_DIR
    OUTPUT_DIR = filedialog.askdirectory(title=title)
    return OUTPUT_DIR

PROGRAM_MENU = {
    '1': {
        'label': "JSON to XLSX",
        'color': 'green',
        'inputFileType': [("JSON files", "*.json")],
        'outputExtension': '.xlsx',
        'callback': convert_json_to_xlsx
    },
    '2': {
        'label': "JSON to CSV",
        'color': 'green',
        'inputFileType': [("JSON files", "*.json")],
        'outputExtension': '.csv',
        'callback': convert_json_to_csv
    },
    '3': {
        'label': "XLSX to JSON",
        'color': 'green',
        'inputFileType': [("XLSX files", "*.xlsx")],
        'outputExtension': '.json',
        'callback': convert_xlsx_to_json
    },
    '4': {
        'label': "XLSX to CSV",
        'color': 'green',
        'inputFileType': [("XLSX files", "*.xlsx")],
        'outputExtension': '.csv',
        'callback': convert_xlsx_to_csv
    },
    '5': {
        'label': "CSV to JSON",
        'color': 'green',
        'fileType': [("CSV files", "*.csv")],
        'extension': '.json',
        'callback': convert_csv_to_json
    },
    '6': {
        'label': "CSV to XLSX",
        'color': 'green',
        'inputFileType': [("CSV files", "*.csv")],
        'outputExtension': '.xlsx',
        'callback': convert_csv_to_xlsx
    },
    '7': {
        'label': "Exit",
        'color': 'red',
    },
}

def main():
    print_banner()
    while True:
        print_menu()
        
        choice = input(colored("Your choice: ", 'blue'))

        if choice == '7':
            cprint("Exiting... Thank you for using MEDICARD FILE CONVERTER!", 'yellow')
            break

        if choice not in ['1', '2', '3', '4', '5', '6']:
            show_error("Invalid choice. Please try again.\n")
            continue

        selectedFeature = PROGRAM_MENU[choice]

        global FILETYPES, EXTENSION
        FILETYPES = selectedFeature['inputFileType']
        EXTENSION = selectedFeature['outputExtension']
                
        print('Select your file...')
        input_path = select_file('Select File', FILETYPES)
        if not input_path:
            show_error("No file selected. Returning to menu...\n")
            continue

        print('Choose the destination folder...')
        output_dir = select_directory('Select Output Directory')
        if not output_dir:
            show_error("No directory selected. Returning to menu...\n")
            continue

        output_name = (input(colored("Enter output file name (or skip to use the original name): ", 'blue')) or os.path.splitext(os.path.basename(input_path))[0]) + EXTENSION
        output_path = os.path.normpath(os.path.join(output_dir, output_name))

        try:
            cprint(f"\nLoading '{os.path.basename(input_path)}'... ", 'yellow', end="")
            if choice in ['1', '2']:
                data = load_json(input_path)
                item_count = len(data)
                print(colored(f"{item_count:,} items found ({os.path.getsize(input_path) / (1024*1024):.2f} MB)", 'cyan'))
            else:
                df = load_csv(input_path)
                print(colored(f"{len(df):,} rows found ({os.path.getsize(input_path) / (1024*1024):.2f} MB)", 'cyan'))
        except Exception as e:
            show_error(f"Error loading file: {e}")
            continue

        try:
            show_progress(f"Converting {selectedFeature['label']}")
            selectedFeature['callback'](input_path, output_path)
            cprint(f"Conversion complete! Output saved to: {output_path} \n", 'green')
        except Exception as e:
            show_error(f"An error occurred during the conversion: {e}")

if __name__ == "__main__": 
    main()
