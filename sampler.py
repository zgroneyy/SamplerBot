#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Nov  8 08:56:09 2023

@author: Ozgur
"""

import openpyxl
import random
import pandas as pd 


def create_empty_excel_file(filename):
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()
    
    # Save the workbook to the specified filename
    workbook.save(filename)

    # Close the workbook
    workbook.close()


def create_test_file(filename, column_num, row_num):
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()

    # Create a reference to the active sheet
    sheet = workbook.active

    # Fill the sheet with random numbers
    for row in range(1, row_num + 1):
        for col in range(1, column_num + 1):
            sheet.cell(row=row, column=col, value=random.randint(1, 100))

    # Save the workbook to the specified filename
    workbook.save(filename)

    # Close the workbook
    workbook.close()

def random_sampler(sampling_set_size, filename, sheet_name=None):
    # Open the Excel file
    workbook = openpyxl.load_workbook(filename)
    
    # Access the sheet by name if provided, or use the first sheet by default
    if sheet_name:
        try:
            sheet = workbook[sheet_name]
        except KeyError:
            raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file.")
    else:
        # Use the first sheet by default
        sheet = workbook.active
    
    # Get the total number of rows in the sheet
    num_rows = sheet.max_row
    
    if sampling_set_size >= num_rows:
        raise ValueError("Sampling set size cannot be greater than or equal to the number of rows.")
    
    # Generate a random sample of row indices
    sampled_rows_indices = random.sample(range(1, num_rows + 1), sampling_set_size)
    
    # Extract data from the sampled rows
    sampled_data = []
    #indices of selected rows. 
    rows_selected = []
    for row_index in sampled_rows_indices:
        row_data = [cell.value for cell in sheet[row_index]]
        sampled_data.append(row_data)
        rows_selected.append(row_index)
    return sampled_data


def stratified_sampler(filename, sheet_name, groupby_column_name, sample_size):
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(filename, sheet_name=sheet_name)

    # Check if the groupby_column_name exists in the DataFrame
    if groupby_column_name not in df.columns:
        raise ValueError(f"'{groupby_column_name}' not found in the DataFrame.")

    # Group the data by the specified column
    grouped = df.groupby(groupby_column_name)

    # Sample data from each group
    sampled_data = []
    
    for _, group_data in grouped:
        group_sample_size = int((sample_size / len(df)) * len(group_data))
        if group_sample_size > 0:
            sampled_data.extend(group_data.sample(n=group_sample_size))

    # Calculate the percentages of each group in the original data
    percentages = (df[groupby_column_name].value_counts() / len(df) * 100).to_dict()

    return sampled_data, percentages


# Example usage: create an Excel file named "sample.xlsx" with 5 columns and 10 rows
create_empty_excel_file("sample.xlsx")
create_test_file("sample.xlsx", column_num=3, row_num=100)

#stratified_sampler("sample.xlsx", sheet_name, groupby_column_name, sample_size)
