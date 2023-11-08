#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Nov  8 08:56:09 2023

@author: Ozgur
"""

import openpyxl
import random
import string
import pandas as pd 
import csv

def create_empty_excel_file(filename):
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()
    
    # Save the workbook to the specified filename
    workbook.save(filename)

    # Close the workbook
    workbook.close()

def print_into_file(data_array, file_type):
    if file_type not in ["xlsx", "csv"]:
        raise ValueError("Invalid file_type. Supported values are 'xlsx' or 'csv'.")

    if file_type == "xlsx":
        # Create a new Excel workbook
        workbook = openpyxl.Workbook()
        
        # Create a reference to the active sheet
        sheet = workbook.active

        # Print the data into the sheet
        for row_index, row_data in enumerate(data_array, start=1):
            for col_index, value in enumerate(row_data, start=1):
                sheet.cell(row=row_index, column=col_index, value=value)

        # Save the workbook as an Excel file named "results.xlsx"
        workbook.save("results.xlsx")

    elif file_type == "csv":
        # Create a CSV file named "results.csv" and print the data
        with open("results.csv", "w", newline="") as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerows(data_array)

def create_test_file(filename, num_columns, num_rows):
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()

    # Create a reference to the active sheet
    sheet = workbook.active

    # Define headers for columns with alphabet letters
    headers = list(string.ascii_uppercase)

    for col_index in range(num_columns):

        # Fill the column with header
        for row_index in range(1, num_rows + 1):
            sheet.cell(row=row_index, column=col_index + 1, value=random.choice((headers)))

        # Fill the second column with a random year between 1950 and 2023
        if col_index == 1:
            for row_index in range(1, num_rows + 1):
                sheet.cell(row=row_index, column=col_index + 1, value=random.randint(1950, 2023))

        # Fill the third column with a random city name
        elif col_index == 2:
            cities = ["Ankara", "Istanbul", "Konya", "Izmir", "Atina", "Bursa"]
            for row_index in range(1, num_rows + 1):
                sheet.cell(row=row_index, column=col_index + 1, value=random.choice(cities))

        # Fill other columns with random numbers
        elif col_index > 2:
            for row_index in range(1, num_rows + 1):
                sheet.cell(row=row_index, column=col_index + 1, value=random.randint(1, 100))

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

def systematic_sampler(sampling_set_size, filename, sheet_name=None):
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
    
    if num_rows == 0:
        raise ValueError("The sheet is empty. No data to sample.")

    # Calculate the step size to sample every nth item
    step_size = num_rows // sampling_set_size

    # Check if the sampling set size is valid
    if step_size < 1:
        raise ValueError("Sampling set size is too large for the given data.")

    # Initialize a list to store the sampled data
    sampled_data = []

    # Sample every nth item starting from the first row
    for row_index in range(1, num_rows + 1, step_size):
        row_data = [cell.value for cell in sheet[row_index]]
        sampled_data.append(row_data)

    return sampled_data

def cluster_sampler(sampling_set_size, sampling_group_column, filename, sheet_name=None):
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
    
    if num_rows == 0:
        raise ValueError("The sheet is empty. No data to sample.")

    # Check if the provided group column exists in the sheet
    if sampling_group_column not in sheet[1]:
        raise ValueError(f"Group column '{sampling_group_column}' not found in the sheet.")

    # Create a list of unique groups based on the specified group column
    groups_array = list(set(sheet.cell(row=i, column=sheet[1].index(sampling_group_column) + 1).value for i in range(2, num_rows + 1)))

    # Select a random group from the groups_array
    selected_group = random.choice(groups_array)

    # Collect all rows with the selected group
    sampled_data = []
    for row_index in range(2, num_rows + 1):
        group_value = sheet.cell(row=row_index, column=sheet[1].index(sampling_group_column) + 1).value
        if group_value == selected_group:
            row_data = [cell.value for cell in sheet[row_index]]
            sampled_data.append(row_data)

    return sampled_data

# Example usage: create an Excel file named "sample.xlsx" with 5 columns and 10 rows
create_empty_excel_file("sample.xlsx")
create_test_file("sample.xlsx",3,100)

#stratified_sampler("sample.xlsx", sheet_name, groupby_column_name, sample_size)
