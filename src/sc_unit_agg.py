import os
import pandas as pd
from typing import Dict
import pdb
import re
import datetime

def import_excel_files(path: str = None) -> Dict[str, pd.DataFrame]:
    """
    Import Excel files from a directory or a single file.
    Only files starting with "Sales Journal for " are imported.
    Returns a dictionary mapping file names to pandas DataFrames.
    The path can be provided as an argument, or selected via a GUI if not provided.
    """
    if path is None:
        try:
            import tkinter
            from tkinter import filedialog, messagebox
        except ImportError:
            raise ImportError("tkinter is required for file dialog. Please install it or provide a path.")

        root = tkinter.Tk()
        root.withdraw()
        root.update()
        msg = "Select an Excel file (must start with 'Sales Journal for ') or select cancel and choose a folder containing such files."
        messagebox.showinfo("Select File or Folder", msg)
        path = filedialog.askopenfilename(
            title=msg,
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        if not path:
            path = filedialog.askdirectory(title="Or select a folder")
        root.destroy()
        if not path:
            raise ValueError("No file or folder selected.")

    excel_data = {}

    def is_valid_file(fname: str) -> bool:
        return fname.startswith("Sales Journal for ") and fname.lower().endswith(('.xls', '.xlsx'))

    def extract_month_year(fname: str) -> str:
        # Extracts "Month Year" from "Sales Journal for Month Year.xlsx"
        match = re.search(r"Sales Journal for ([A-Za-z]+ \d{4})", fname)
        if match:
            return match.group(1)
        else:
            return fname  # fallback to original if pattern not found

    if os.path.isdir(path):
        for fname in os.listdir(path):
            if is_valid_file(fname):
                fpath = os.path.join(path, fname)
                key = extract_month_year(fname)
                excel_data[key] = pd.read_excel(fpath)
    elif os.path.isfile(path) and is_valid_file(os.path.basename(path)):
        fname = os.path.basename(path)
        key = extract_month_year(fname)
        excel_data[key] = pd.read_excel(path)
    else:
        raise ValueError('Path must be an Excel file or a directory containing Excel files, and file(s) must start with "Sales Journal for ".')
    
    if not excel_data:
        raise ValueError('No valid Excel files found starting with "Sales Journal for ".')
    
    return excel_data

def filter_units(data_dict: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    """
    For each DataFrame in data_dict, find rows where the second column contains 'G', 'H', or 'I'.
    If found, copy the second and sixth column values to a new DataFrame.
    Clean the sixth column: remove parenthesis, remove '$', and make value absolute.
    If not found, create a DataFrame with a message and 0.
    Returns a dictionary with the same keys as data_dict.
    Adds a row to each DataFrame that is the sum of all Rent.
    """
    filtered_dict = {}
    for key, df in data_dict.items():
        # Ensure there are at least 6 columns
        if df.shape[1] < 6:
            filtered_dict[key] = pd.DataFrame([["no units of conataining G, H, or I", 0]])
            continue

        col2 = df.iloc[:, 1].astype(str)
        mask = col2.str.contains(r'[GHI]', na=False)
        if mask.any():
            filtered_df = df.loc[mask, [df.columns[1], df.columns[5]]].copy()
            # Drop rows where the 6th column is NaN
            filtered_df = filtered_df[filtered_df[df.columns[5]].notna()]
            # Clean the sixth column
            filtered_df[df.columns[5]] = (
                filtered_df[df.columns[5]]
                .astype(str)
                .str.replace(r'[\$,()]', '', regex=True)
                .astype(float)
                .abs()
            )
            # Drop rows where the 6th column is 0
            filtered_df = filtered_df[filtered_df[df.columns[5]] != 0]
            if not filtered_df.empty:
                # Add sum row for Rent
                rent_sum = filtered_df[df.columns[5]].sum()
                sum_row = pd.DataFrame([[f"Total Rent", rent_sum]], columns=[df.columns[1], df.columns[5]])
                filtered_df = pd.concat([filtered_df, sum_row], ignore_index=True)
                filtered_dict[key] = filtered_df
            else:
                filtered_dict[key] = pd.DataFrame([["no units of conataining G, H, or I", 0]], columns=[df.columns[1], df.columns[5]])
        else:
            filtered_dict[key] = pd.DataFrame([["no units of conataining G, H, or I", 0]], columns=[df.columns[1], df.columns[5]])
    return filtered_dict

def write_unit_aggregation_report(filtered_dict: Dict[str, pd.DataFrame]):
    """
    Write each DataFrame in filtered_dict to a separate sheet in an Excel file.
    The user selects the destination and filename via a GUI.
    The default filename is today's date + ' unit aggregation report.xlsx'.
    """
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    default_filename = f"{today_str} unit aggregation report.xlsx"

    try:
        import tkinter
        from tkinter import filedialog
    except ImportError:
        raise ImportError("tkinter is required for file dialog. Please install it or provide a path.")

    root = tkinter.Tk()
    root.withdraw()
    root.update()
    file_path = filedialog.asksaveasfilename(
        title="Save Excel Report As...",
        defaultextension=".xlsx",
        initialfile=default_filename,
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.destroy()

    if not file_path:
        print("No file selected. Report not saved.")
        return

    with pd.ExcelWriter(file_path) as writer:
        for sheet_name, df in filtered_dict.items():
            # Sheet names in Excel have a max length of 31 characters
            safe_sheet_name = str(sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
    print(f"Report written to {file_path}")


data_dict = import_excel_files()
filtered_dict = filter_units(data_dict)
write_unit_aggregation_report(filtered_dict)