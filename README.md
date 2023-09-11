# Excel File Splitter

Simple GUI to split large excel files.


## Usage
- Select the input .xlsx file
- Specify number of rows in each output file (not including header row)
- Specify a suffix for the output files (e.g. "{number}_split" will result in "original_filename_001_split.xlsx" etc.)

Note that:
Output files are generated in the same folder as the original file.
The original structure (sheet names and order) is maintained in the output, only the sheet with the greatest number of rows is split, the remaining sheets are copied into every output file.
