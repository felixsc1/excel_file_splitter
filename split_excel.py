import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook


def set_integer_format(workbook, dataframe):
    """
    Apply integer format to columns in the given dataframe that have integer values.
    """
    for col_idx, (col, col_data) in enumerate(dataframe.items(), start=1):
        try:
            # Check if all non-NaN values are integers
            if (
                col_data.dropna()
                .apply(lambda x: isinstance(x, (int, float)) and float(x).is_integer())
                .all()
            ):
                col_letter = get_column_letter(col_idx)
                for row in workbook[col_letter]:
                    row.number_format = "#"
        except ValueError:
            continue


def split_excel(input_file, rows_per_file, suffix_template):
    # Read all sheets from the excel file
    with pd.ExcelFile(input_file) as xls:
        sheet_names = xls.sheet_names  # Capture the order of sheets
        all_sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in sheet_names}

    # Identify the sheet with the maximum rows
    max_rows_sheet_name = max(all_sheets, key=lambda sheet: len(all_sheets[sheet]))
    max_rows_sheet = all_sheets[max_rows_sheet_name]

    # Calculate number of smaller files required for the max_rows_sheet
    no_of_files = len(max_rows_sheet) // rows_per_file + (
        1 if len(max_rows_sheet) % rows_per_file else 0
    )

    # Determine the output directory and original filename from the input file's path
    output_dir = os.path.dirname(input_file)
    original_filename = os.path.splitext(os.path.basename(input_file))[
        0
    ]  # get filename without extension

    for i in range(no_of_files):
        # Extract subset of rows for the main sheet
        start_row = i * rows_per_file
        end_row = (i + 1) * rows_per_file
        subset_main_sheet = max_rows_sheet[start_row:end_row]

        # Create a new excel writer for this subset
        formatted_number = str(i + 1).zfill(3)  # this will produce '001', '002', etc.
        output_suffix = suffix_template.replace("{number}", formatted_number)
        output_file = os.path.join(
            output_dir, f"{original_filename}_{output_suffix}.xlsx"
        )

        with pd.ExcelWriter(
            output_file, engine="openpyxl"
        ) as writer:  # Use openpyxl engine
            # Write sheets in their original order
            for sheet_name in sheet_names:
                if sheet_name == max_rows_sheet_name:
                    subset_main_sheet.to_excel(
                        writer, sheet_name=sheet_name, index=False
                    )
                else:
                    all_sheets[sheet_name].to_excel(
                        writer, sheet_name=sheet_name, index=False
                    )

            # Post-process formatting
            set_integer_format(writer.sheets[max_rows_sheet_name], subset_main_sheet)
            for sheet_name in sheet_names:
                if sheet_name != max_rows_sheet_name:
                    set_integer_format(
                        writer.sheets[sheet_name], all_sheets[sheet_name]
                    )

        print(f"Wrote file {output_file}")


def main_gui():
    root = tk.Tk()
    root.title("Excel Splitter")

    def select_input_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        input_file_var.set(file_path)

    def split_and_save():
        input_file = input_file_var.get()
        rows = int(rows_per_file_var.get())
        suff_template = suffix_var.get()

        split_excel(input_file, rows, suff_template)
        messagebox.showinfo("Success", "Excel file split successfully!")

    # GUI Elements
    input_file_var = tk.StringVar()
    rows_per_file_var = tk.StringVar(value="1000")  # Default value set to 1000
    suffix_var = tk.StringVar(value="{number}")  # Default value

    tk.Label(root, text="Select Excel File:").pack(pady=20)
    tk.Entry(root, textvariable=input_file_var, width=50).pack(pady=5)
    tk.Button(root, text="Browse", command=select_input_file).pack(pady=10)

    tk.Label(root, text="Rows per file:").pack(pady=20)
    tk.Entry(root, textvariable=rows_per_file_var, width=50).pack(pady=5)

    tk.Label(root, text="Suffix for split files (use {number} for numbering):").pack(
        pady=20
    )
    tk.Entry(root, textvariable=suffix_var, width=50).pack(pady=5)

    tk.Button(root, text="Split and Save", command=split_and_save).pack(pady=20)

    root.mainloop()


if __name__ == "__main__":
    main_gui()
