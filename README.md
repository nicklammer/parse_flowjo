# FlowJo output parser
Short script that can filter FlowJo output (.xls) based on depth and fluor. Cell IDs within sample names are used to format the statistics and cell count data into a 96-well plate format. Both statistics and cell count tables are saved as a .xlsx file for package compatibility reasons.

## Prerequisites
This was written to work in python 3.12. See environment.yml for package requirements.

## Usage
Run parse_flowjo.py. You will then be prompted for four things: FlowJo output file, depth filter (input as the number of >), fluor filter, and location of the output file.

## Compiling
I compile into an executable using pyinstaller. This command should work if you've activated your conda environment:
```
pyinstaller --onefile --windowed --hidden-import openpyxl.cell._writer parse_flowjo.py
```