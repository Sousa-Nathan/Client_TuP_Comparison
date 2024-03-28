import pandas as pd
import numpy as np
import tkinter as tk
import os

from pathlib import Path
from tkinter import filedialog
from datetime import datetime

def filename_name_dir(wb1, wb2, out_dir):
    now = datetime.now()
    date = now.strftime("%Y_%m_%d")
    time = now.strftime("%H_%M_%S")
    
    wb1_dir_split = wb1.split("/")
    wb2_dir_split = wb2.split("/")
    
    wb1_ext_split = wb1_dir_split[-1].split(".")
    wb2_ext_split = wb2_dir_split[-1].split(".")
    
    wb1_version = wb1_ext_split[0][:19]
    wb2_version = wb2_ext_split[0][:19]
    
    comparison_tup_file = os.path.join(out_dir, f"TuP Comparison {wb1_version} v. {wb2_version} {date} {time}.xlsx")
    
    return comparison_tup_file

def tup_comparison(wb_dir1, wb_dir2, out_dir):
    workbook = pd.ExcelFile(wb_dir1)

    sheetnames = workbook.sheet_names

    no_no_sheets = ["Readme", "SAR Workbook Friendly ANT 0", "SAR Workbook Friendly ANT 1", "SAR Workbook Friendly ANT 2", "SAR Workbook Friendly ANT 5", "SAR Workbook Friendly ANT 6", "SAR Workbook Friendly ANT 7"]

    tup_data_sheets = [n for i, n in enumerate(sheetnames) if n not in no_no_sheets]

    tup_workbook_1_data = [pd.read_excel(wb_dir1, sheet_name = tup_data_sheets[i]) for i, n in enumerate(tup_data_sheets)]
    tup_workbook_2_data = [pd.read_excel(wb_dir2, sheet_name = tup_data_sheets[i]) for i, n in enumerate(tup_data_sheets)]

    for data in range(len(tup_workbook_1_data)):
        tup_workbook_1_data[data].equals(tup_workbook_2_data[data])
        
        comparison_values = tup_workbook_1_data[data].values == tup_workbook_2_data[data].values
        
        rows, cols = np.where(comparison_values == False)
        
        for item in zip(rows, cols):
            tup_workbook_1_data[data].iloc[item[0], item[1]] = f"New: {tup_workbook_1_data[data].iloc[item[0], item[1]]:.1f} | Delta: {tup_workbook_2_data[data].iloc[item[0], item[1]]:.1f} -> {tup_workbook_1_data[data].iloc[item[0], item[1]]:.1f}"
        
        try:
            abs_filepath = Path(out_dir).resolve(strict = True)

        except FileNotFoundError:
            with pd.ExcelWriter(out_dir) as writer: # pylint: disable=abstract-class-instantiated
                tup_workbook_1_data[data].to_excel(writer, sheet_name = f"{tup_data_sheets[data]}", index = False)

        else:
            with pd.ExcelWriter(f"{out_dir}", mode = "a") as writer: # pylint: disable=abstract-class-instantiated
                tup_workbook_1_data[data].to_excel(writer, sheet_name = f"{tup_data_sheets[data]}", index = False)

if __name__ == "__main__":
    
    root = tk.Tk()
    root.withdraw()
    
    tup_workbook_1 = filedialog.askopenfilename(title = "New TuP", filetypes = (("Excel File", "*.xlsx*"),))
    tup_workbook_2 = filedialog.askopenfilename(title = "Old TuP", filetypes = (("Excel File", "*.xlsx*"),))
    comparison_tup_dir = filedialog.askdirectory(title = "Output Directory")
    
    output_tup_filename = filename_name_dir(tup_workbook_1, tup_workbook_2, comparison_tup_dir)
    
    tup_comparison(tup_workbook_1, tup_workbook_2, output_tup_filename)
    
    print("\nDone\n")