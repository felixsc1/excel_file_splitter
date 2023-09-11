import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


def split_excel(input_file, rows_per_file, suffix_template):
    # Determine the output directory from the input file's path and its filename
    output_dir = os.path.dirname(input_file)
    original_filename = os.path.splitext(os.path.basename(input_file))[
        0
    ]  # get filename without extension

    # Read the large excel file
    df = pd.read_excel(input_file)

    # Calculate number of smaller files required
    no_of_files = len(df) // rows_per_file + (1 if len(df) % rows_per_file else 0)

    for i in range(no_of_files):
        # Extract subset of rows for this file
        start_row = i * rows_per_file
        end_row = (i + 1) * rows_per_file
        subset = df[start_row:end_row]

        # Format the output file name
        formatted_number = str(i + 1).zfill(3)  # this will produce '001', '002', etc.
        output_suffix = suffix_template.replace("{number}", formatted_number)
        output_file = os.path.join(
            output_dir, f"{original_filename}_{output_suffix}.xlsx"
        )

        # Write the subset to a new excel file
        subset.to_excel(output_file, index=False)
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
