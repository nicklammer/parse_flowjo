# parse and transpose flowjo output (.xls)
# testing out flowjo parsing
import pandas as pd
import xlrd
import openpyxl
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox

def excelOpen(file_raw):
    df = pd.read_excel(file_raw, engine = 'xlrd')
    df = df.fillna('')

    return(df)

def excelFilter(df, depth, fluor):
    depth_str = "> " * depth
    df_depth = df[df['Depth'] == depth_str]
    df_filter = pd.DataFrame(columns=df_depth.columns)
    for i, row in df_depth.iterrows():
        # plate levels are separated by "/"
        # get the last part and see if it matches fluor
        name = row['Name']
        name_end = name.split("/")[-1]
        if fluor == name_end:
            df_filter.loc[i] = row
    return(df_filter)

def transformData(df):
    wells = []
    stats = []
    cellcounts = []
    # get well ID, stats, and cell counts in separate lists
    for _, row in df.iterrows():
        name = row['Name']
        well = name[0:3].strip(".")
        wells.append(well)

        stats.append(row['Statistic'])
        cellcounts.append(row['#Cells'])

    # hard coded plate row and column names
    letters = ["A", "B", "C", "D", "E", "F", "G", "H"]
    numbers = ["1", "2", "3", "4", "5", "6", "7", "8",
               "9", "10", "11", "12"]

    # column format dataframe for saving
    df_columns = pd.DataFrame(
        {"Well": wells,
         "Statistic": stats,
         "#Cells": cellcounts
         }
    )
    df_columns = df_columns.set_index("Well")
    # create new index to sort by plate column (number)
    well_index_new = []
    for number in numbers:
        for letter in letters:
            well_index_new.append(letter+number)
    df_columns_sorted = df_columns.reindex(well_index_new)

    # this long thing makes a dictionary assigning stats and cell counts
    # based on the well ID in the sample name
    stats_dict = {}
    cellcounts_dict = {}
    for letter in letters:
        # I wanted to accomodate empty wells so start with empty lists
        # consituting max number of values in a row
        stats_row_values = [''] * 12
        cellcounts_row_values = [''] * 12
        wells_in_row = []
        stats_in_row = []
        cellcounts_in_row = []
        # get lists of values in 1 row at a time
        for i, well in enumerate(wells):
            if well[0] == letter:
                wells_in_row.append(well)
                stats_in_row.append(stats[i])
                cellcounts_in_row.append(cellcounts[i])
        # assign values to the empty row lists based on the well ID column #
        for i, well in enumerate(wells_in_row):
            row_idx = int(well[1:]) - 1
            stats_row_values[row_idx] = stats_in_row[i]
            cellcounts_row_values[row_idx] = cellcounts_in_row[i]
        
        stats_dict[letter] = stats_row_values
        cellcounts_dict[letter] = cellcounts_row_values

    # make dataframes from the dictionaries 
    # make the keys into indices
    # column names are well numbers
    df_stats = pd.DataFrame.from_dict(stats_dict, orient = "index", columns = numbers)
    df_cellcounts = pd.DataFrame.from_dict(cellcounts_dict, orient = "index", columns = numbers)

    return(df_columns_sorted, df_stats, df_cellcounts)
        
def excelWrite(df_columns, df_stats, df_cellcounts, fluor, out_path):
    dfs = [df_stats, df_cellcounts]
    index_labels = ["Statistics", "#Cells"]
    startrow = 0
    # I got this from a help request
    # writes both dataframes to 1 sheet arranged vertically
    with pd.ExcelWriter(out_path) as writer:
        for i, df in enumerate(dfs):
            df.to_excel(writer, startrow = startrow,
                        index_label = index_labels[i],
                        sheet_name = fluor)
            startrow += (df.shape[0] + 2)
        df_columns.to_excel(writer, sheet_name = "Columns")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    file_raw = filedialog.askopenfilename(title = "Select FlowJo output",
                                        filetypes = [("Excel files", ".xlsx .xls")])
    depth = simpledialog.askinteger(title = "Enter Depth", 
                                    prompt = "Enter depth (>) as a number")
    fluor = simpledialog.askstring(title = "Enter Fluor", 
                                    prompt = "Enter fluor name to filter on")
    out_path = filedialog.asksaveasfilename(title = "Save tables as...",
                                            defaultextension=".xlsx")

    df = excelOpen(file_raw)
    df_filter = excelFilter(df, depth, fluor)
    df_columns_sorted, df_stats, df_cellcounts = transformData(df_filter)
    excelWrite(df_columns_sorted, df_stats, df_cellcounts, fluor, out_path)

    messagebox.showinfo(message = f"Done! File saved at {out_path}.")