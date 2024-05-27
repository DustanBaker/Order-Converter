# This order converter app imports CSV files with Kit ID's and converts the order to a SKU based CSV upload.
# Created by Dusty Baker December 2023
#Updated by Dusty Baker May 2024 to add UPS shipping options, and EAGLE file manager.
import csv
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import customtkinter
from PIL import Image
from datetime import datetime
import os

customtkinter.set_appearance_mode("system")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green

# Initialize error_added at the top level of your script
error_added = False

# Function to read a csv file to a dictionary to determine the save file name for eagle
def read_projects_csv_to_dict(input_file):
    projects_dict = {}

    with open(input_file, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        project_numbers = next(csv_reader)  # Read the first row
        project_names = next(csv_reader)  # Read the second row

        for project_num, project_name in zip(project_numbers, project_names):
            projects_dict[int(project_num)] = project_name

    return projects_dict

# Load projects.csv and define projects_dict globally
projects_file_path = 'assets/projects.csv'
projects_dict = read_projects_csv_to_dict(projects_file_path)


#funtinon to read a csv file to a dictionary to be used in the program
def read_csv_to_dict(input_file):
    data = []
    with open(input_file, 'r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for row in csv_reader:
            data.append(row)
    return data


def manipulate_USCG_data(input_data):
    manipulated_data = []
    global error_added

    for row in input_data:
        # Check if the row is the header row and add it to the manipulated data

        if row['12'] == 'Item SKU' or row['12'] == 'Kit ID':
            manipulated_data.append(row)

        elif row['12'] == 'SO-001':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWP'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)
        elif row['12'] == 'SO-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWP'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)
        elif row['12'] == 'SO-003':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWP'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)
        elif row['12'] == 'SO-004':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWP'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)
        elif row['12'] == 'SR-001':
            # Create Three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGBL'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SR-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGBL'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'SR-003':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGBL'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SR-004':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGBL'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day w/return lbl'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'AO-001':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWR'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'AO-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWR'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'AO-003':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWR'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'AO-004':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWR'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'AR-001':
            # Create Three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGBL-AR'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'AR-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGBL-AR'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'AR-003':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGBL-AR'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'AR-004':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGBL-AR'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day w/return lbl'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SRR-001':
            # Create Three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCFT'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SRR-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCFT'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'SRR-003':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCFT'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SRR-004':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCFT'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day w/return lbl'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SRRL-001':
            # Create Three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCFT-SRRL'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SRRL-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCFT-SRRL'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'SRRL-003':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCFT-SRRL'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SRRL-004':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCFT-SRRL'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day w/return lbl'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'SE-001':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-AMDT'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'SE-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-AMDT'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'SE-003':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-AMDT'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'SE-004':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-AMDT'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'AE-001':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-AMDT-AE'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'AE-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-AMDT-AE'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'AE-003':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-AMDT-AE'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'AE-004':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-AMDT-AE'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'RE-001':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'RE-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'RE-003':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'RE-004':
            # Create three new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = '210-BBBY'
            new_row_3['16'] = 'FedEx 2Day w/return lbl'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Dusty Baker'
            new_row_3['25'] = '791214637'
            manipulated_data.append(new_row_3)

        elif row['12'] == 'K-001':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWY'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'K-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWY'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'K-003':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWY'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'K-004':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BFWY'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'SRV-001':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'SRV-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'SRV-003':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)


        elif row['12'] == 'SRV-004':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-210-AZBS'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'ARV-001':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU-ARV'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'ARV-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU-ARV'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'ARV-003':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU-ARV'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-AZBS'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'ARV-004':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BGGU-ARV'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-210-AZBS'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'SVDI-001':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCXC'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'SVDI-002':
            # Create one new row with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCXC'
            new_row_1['16'] = 'Priority Overnight w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

        elif row['12'] == 'SVDI-003':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCXC'
            new_row_1['16'] = 'FedEx 2Day'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

        elif row['12'] == 'SVDI-004':
            # Create two new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = '210-BCXC'
            new_row_1['16'] = 'FedEx 2Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Dusty Baker'
            new_row_1['25'] = '791214637'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = '210-BBBY'
            new_row_2['16'] = 'FedEx 2Day w/return lbl'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Dusty Baker'
            new_row_2['25'] = '791214637'
            manipulated_data.append(new_row_2)

    return manipulated_data

def manipulate_Terminix_data(input_data):
    manipulated_data = []
    errors = []
    global error_added

    for row in input_data:
        # Check if the row is the header row and add it to the manipulated data
        if row['12'] == 'Item SKU':
            manipulated_data.append(row)

        elif row['12'] == 'standard':
            # Create six new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = 'Headset'
            new_row_1['16'] = 'FedEx Ground w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDEG'
            new_row_1['23'] = 'Kevin Mitchell'
            new_row_1['25'] = '177264750'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = 'Monitor'
            new_row_2['16'] = 'FedEx Ground w/return lbl'
            new_row_2['20'] = '1'
            new_row_2['21'] = 'FDEG'
            new_row_2['23'] = 'Kevin Mitchell'
            new_row_2['25'] = '177264750'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = 'Desktop'
            new_row_3['16'] = 'FedEx Ground w/return lbl'
            new_row_3['20'] = '1'
            new_row_3['21'] = 'FDEG'
            new_row_3['23'] = 'Kevin Mitchell'
            new_row_3['25'] = '177264750'
            manipulated_data.append(new_row_3)

            new_row_4 = row.copy()
            new_row_4['12'] = 'Webcam'
            new_row_4['16'] = 'FedEx Ground w/return lbl'
            new_row_4['20'] = '1'
            new_row_4['21'] = 'FDEG'
            new_row_4['23'] = 'Kevin Mitchell'
            new_row_4['25'] = '177264750'
            manipulated_data.append(new_row_4)

            new_row_5 = row.copy()
            new_row_5['12'] = 'NetCbl'
            new_row_5['16'] = 'FedEx Ground w/return lbl'
            new_row_5['20'] = '1'
            new_row_5['21'] = 'FDEG'
            new_row_5['23'] = 'Kevin Mitchell'
            new_row_5['25'] = '177264750'
            manipulated_data.append(new_row_5)

            new_row_6 = row.copy()
            new_row_6['12'] = 'Box'
            new_row_6['16'] = 'FedEx Ground w/return lbl'
            new_row_6['20'] = '1'
            new_row_6['21'] = 'FDEG'
            new_row_6['23'] = 'Kevin Mitchell'
            new_row_6['25'] = '177264750'
            manipulated_data.append(new_row_6)

        elif row['12'] == '2day':
            # Create six new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = 'Headset'
            new_row_1['16'] = 'FedEx 2 Day w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Kevin Mitchell'
            new_row_1['25'] = '177264750'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = 'Monitor'
            new_row_2['16'] = 'FedEx 2 Day w/return lbl'
            new_row_2['20'] = '1'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Kevin Mitchell'
            new_row_2['25'] = '177264750'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = 'Desktop'
            new_row_3['16'] = 'FedEx 2 Day w/return lbl'
            new_row_3['20'] = '1'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Kevin Mitchell'
            new_row_3['25'] = '177264750'
            manipulated_data.append(new_row_3)

            new_row_4 = row.copy()
            new_row_4['12'] = 'Webcam'
            new_row_4['16'] = 'FedEx 2 Day w/return lbl'
            new_row_4['20'] = '1'
            new_row_4['21'] = 'FDE'
            new_row_4['23'] = 'Kevin Mitchell'
            new_row_4['25'] = '177264750'
            manipulated_data.append(new_row_4)

            new_row_5 = row.copy()
            new_row_5['12'] = 'NetCbl'
            new_row_5['16'] = 'FedEx 2 Day w/return lbl'
            new_row_5['20'] = '1'
            new_row_5['21'] = 'FDE'
            new_row_5['23'] = 'Kevin Mitchell'
            new_row_5['25'] = '177264750'
            manipulated_data.append(new_row_5)

            new_row_6 = row.copy()
            new_row_6['12'] = 'Box'
            new_row_6['16'] = 'FedEx 2 Day w/return lbl'
            new_row_6['20'] = '1'
            new_row_6['21'] = 'FDE'
            new_row_6['23'] = 'Kevin Mitchell'
            new_row_6['25'] = '177264750'
            manipulated_data.append(new_row_6)

        elif row['12'] == 'overnight':
            # Create six new rows with the specified values
            new_row_1 = row.copy()
            new_row_1['12'] = 'Headset'
            new_row_1['16'] = 'Overnight Priority w/return lbl'
            new_row_1['20'] = '1'
            new_row_1['21'] = 'FDE'
            new_row_1['23'] = 'Kevin Mitchell'
            new_row_1['25'] = '177264750'
            manipulated_data.append(new_row_1)

            new_row_2 = row.copy()
            new_row_2['12'] = 'Monitor'
            new_row_2['16'] = 'Overnight Priority w/return lbl'
            new_row_2['20'] = '1'
            new_row_2['21'] = 'FDE'
            new_row_2['23'] = 'Kevin Mitchell'
            new_row_2['25'] = '177264750'
            manipulated_data.append(new_row_2)

            new_row_3 = row.copy()
            new_row_3['12'] = 'Desktop'
            new_row_3['16'] = 'Overnight Priority w/return lbl'
            new_row_3['20'] = '1'
            new_row_3['21'] = 'FDE'
            new_row_3['23'] = 'Kevin Mitchell'
            new_row_3['25'] = '177264750'
            manipulated_data.append(new_row_3)

            new_row_4 = row.copy()
            new_row_4['12'] = 'Webcam'
            new_row_4['16'] = 'Overnight Priority w/return lbl'
            new_row_4['20'] = '1'
            new_row_4['21'] = 'FDE'
            new_row_4['23'] = 'Kevin Mitchell'
            new_row_4['25'] = '177264750'
            manipulated_data.append(new_row_4)

            new_row_5 = row.copy()
            new_row_5['12'] = 'NetCbl'
            new_row_5['16'] = 'Overnight Priority w/return lbl'
            new_row_5['20'] = '1'
            new_row_5['21'] = 'FDE'
            new_row_5['23'] = 'Kevin Mitchell'
            new_row_5['25'] = '177264750'
            manipulated_data.append(new_row_5)

            new_row_6 = row.copy()
            new_row_6['12'] = 'Box'
            new_row_6['16'] = 'Overnight Priority w/return lbl'
            new_row_6['20'] = '1'
            new_row_6['21'] = 'FDE'
            new_row_6['23'] = 'Kevin Mitchell'
            new_row_6['25'] = '177264750'
            manipulated_data.append(new_row_6)

    return manipulated_data


def write_csv_from_dict(output_file, output_data, fieldnames):
    with open(output_file, 'w', newline='') as csv_file:
        csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        csv_writer.writeheader()
        csv_writer.writerows(output_data)


# create a function that reads a CSV file into a dictionary and removes additional commas
def check_and_remove_additional_commas(input_file):
    cleaned_data = []

    with open(input_file, 'r') as csv_file:
        reader = csv.reader(csv_file)
        header = next(reader)
        cleaned_data.append(header)

        for row in reader:
            cleaned_row = [value.replace(',', '') for value in row]
            cleaned_data.append(cleaned_row)

    return cleaned_data

# create a function that writes a cleaned CSV file
def write_cleaned_csv(output_file, cleaned_data):
    with open(output_file, 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerows(cleaned_data)

# create a function that checks for additional commas within strings
def clean_csv_commas(input_path, output_path):
    cleaned_data = check_and_remove_additional_commas(input_path)
    write_cleaned_csv(output_path, cleaned_data)

def USCG_Error_Handling(input_file):
    with open(input_file, 'r') as csv_file:
        reader = csv.reader(csv_file)
        header = next(reader)
        errors = []

        valid_kit_ids = [
            'SO-001', 'SO-002', 'SO-003', 'SO-004',
            'SR-001', 'SR-002', 'SR-003', 'SR-004',
            'AO-001', 'AO-002', 'AO-003', 'AO-004',
            'AR-001', 'AR-002', 'AR-003', 'AR-004',
            'SRR-001', 'SRR-002', 'SRR-003', 'SRR-004',
            'SRRL-001', 'SRRL-002', 'SRRL-003', 'SRRL-004',
            'SE-001', 'SE-002', 'SE-003', 'SE-004',
            'AE-001', 'AE-002', 'AE-003', 'AE-004',
            'RE-001', 'RE-002', 'RE-003', 'RE-004',
            'K-001', 'K-002', 'K-003', 'K-004',
            'SRV-001', 'SRV-002', 'SRV-003', 'SRV-004',
            'ARV-001', 'ARV-002', 'ARV-003', 'ARV-004',
            'SVDI-001', 'SVDI-002', 'SVDI-003', 'SVDI-004',
            'Item SKU', 'Kit ID'
        ]

        for line_num, row in enumerate(reader, start=2):  # Start from 2 to account for the header rows
            if len(row) > 11 and row[11] not in valid_kit_ids:
                errors.append(f"Error: Incorrect Kit ID found in row {line_num}")

        return errors

def Terminix_error_handling(input_file):

    with open(input_file, 'r') as csv_file:
        reader = csv.reader(csv_file)
        header = next(reader)
        errors = []

        valid_kit_ids = [
            'standard', '2day', 'overnight',
            'Item SKU', 'Kit ID'
        ]

        for line_num, row in enumerate(reader, start=2):  # Start from 2 to account for the header rows
            if len(row) > 11 and row[11] not in valid_kit_ids:
                errors.append(f"Error: Incorrect Kit ID found in row {line_num}")

        return errors


def convert_USCG_csv(input_path, output_path):
    # Read data from the input CSV file into a dictionary
    input_data = read_csv_to_dict(input_path)

    # Manipulate the data
    manipulated_data = manipulate_USCG_data(input_data)

    # Get the field names from the input data
    fieldnames = input_data[0].keys() if input_data else []

    # Write the manipulated data to a new CSV file
    write_csv_from_dict(output_path, manipulated_data, fieldnames)

def convert_Terminix_csv(input_path, output_path):
    # Read data from the input CSV file into a dictionary
    input_data = read_csv_to_dict(input_path)

    # Manipulate the data
    manipulated_data = manipulate_Terminix_data(input_data)

    # Get the field names from the input data
    fieldnames = input_data[0].keys() if input_data else []

    # Write the manipulated data to a new CSV file
    write_csv_from_dict(output_path, manipulated_data, fieldnames)


# create a function that creates a tkinter button that calls the convert_csv function
def USCG_convert_button_click():
    global error_added  # Declare error_added as a global variable
    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File for USCG", filetypes=[("CSV Files", "*.csv")])

    # Clean the input file to remove additional commas
    clean_csv_commas(input_path, input_path)

    # Check for correct USCG kit ID
    errors = USCG_Error_Handling(input_path)
    if errors:
        messagebox.showerror(title="CSV Error", message="\n".join(errors))
        return

    # Create the default output file path
    current_date = datetime.now().strftime("%m-%d-%Y")
    default_directory = r"T:\3PL Files\Shipment Template"
    default_filename = f"USCG Shipment {current_date}.csv"
    default_output_path = os.path.join(default_directory, default_filename)

    try:
        # Attempt to save to the default path
        convert_USCG_csv(input_path, default_output_path)
        messagebox.showinfo(title="Semper Paratus", message=f"File saved successfully at {default_output_path}")
    except (OSError, IOError):
        # If the default path is unavailable, ask the user for the output directory and base filename
        output_base_path = filedialog.asksaveasfilename(
            title="Select Output Directory and Base Filename",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=default_filename
        )

        if not output_base_path:
            return  # User cancelled the save dialog

        # Save to the selected path
        convert_USCG_csv(input_path, output_base_path)
        messagebox.showinfo(title="Semper Paratus", message=f"File saved successfully at {output_base_path}")


# create a function that creates a tkinter button that calls the convert_csv function for Terminix
def Terminix_convert_button_click():
    global error_added  # Declare error_added as a global variable
    # Get the input and output file paths
    input_path = filedialog.askopenfilename(title="Select Input File for Terminix", filetypes=[("CSV Files", "*.csv")])

    # Clean the input file to remove additional commas
    clean_csv_commas(input_path, input_path)

    # Check for correct Terminix kit ID
    errors = Terminix_error_handling(input_path)
    if errors:
        messagebox.showerror(title="CSV Error", message="\n".join(errors))
        return

    # Create the default output file path
    current_date = datetime.now().strftime("%m-%d-%Y")
    default_directory = r"T:\3PL Files\Shipment Template"
    default_filename = f"Terminix Shipment {current_date}.csv"
    default_output_path = os.path.join(default_directory, default_filename)

    try:
        # Attempt to save to the default path
        convert_Terminix_csv(input_path, default_output_path)
        messagebox.showinfo(title="Sweet Liberty!", message=f"File saved successfully at {default_output_path}")
    except (OSError, IOError):
        # If the default path is unavailable, ask the user for the output directory and base filename
        output_base_path = filedialog.asksaveasfilename(
            title="Select Output Directory and Base Filename",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=default_filename
        )

        if not output_base_path:
            return  # User cancelled the save dialog

        # Save to the selected path
        convert_Terminix_csv(input_path, output_base_path)
        messagebox.showinfo(title="Sweet Liberty!", message=f"File saved successfully at {output_base_path}")

##
def Eagle_button_click():
    global projects_dict

    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File", filetypes=[("CSV Files", "*.csv")])

    if not input_path:
        return  # User cancelled the file dialog

    # Clean the input file to remove additional commas
    clean_csv_commas(input_path, input_path)

    # Read the third row, first column to get the project number
    with open(input_path, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        next(csv_reader)  # Skip the first row
        next(csv_reader)  # Skip the second row
        third_row = next(csv_reader)  # Read the third row
        project_number = int(third_row[0])  # Assuming project number is an integer

    # Map the project number to the project name using projects_dict
    project_name = projects_dict.get(project_number, "Unknown_Project")

    # Create the output file path
    current_date = datetime.now().strftime("%m-%d-%Y")
    default_directory = r"T:\3PL Files\Shipment Template"
    default_filename = f"{project_name} Shipment {current_date}.csv"
    output_path = os.path.join(default_directory, default_filename)

    # Save the cleaned CSV file with the appropriate name
    try:
        os.makedirs(default_directory, exist_ok=True)
        write_cleaned_csv(output_path, check_and_remove_additional_commas(input_path))
        messagebox.showinfo(title="Success", message=f"File saved successfully at {output_path}")
    except (OSError, IOError):
        # If the default path is unavailable, ask the user for the output directory and base filename
        output_base_path = filedialog.asksaveasfilename(
            title="Select Output Directory and Base Filename",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=default_filename
        )
        if not output_base_path:
            return  # User cancelled the save dialog

        # Save to the selected path
        write_cleaned_csv(output_base_path, check_and_remove_additional_commas(input_path))
        messagebox.showinfo(title="Success", message=f"File saved successfully at {output_base_path}\n Suck it Mantis!")




## GUI_______________________________________________________________________________________________________
root = customtkinter.CTk()
root.title("Dusty's Order Converter")
root.geometry("600x600")
root.iconbitmap('assets/Lambda.ico')

# Create label
Header_Label1 = customtkinter.CTkLabel(root, text="Version 1.4, Now with EAGLE!\n UPS tab is in Beta, please ignore.")
Header_Label1.pack(pady=5)

#create a tab view with custom tkinter
My_tab = customtkinter.CTkTabview(root)
My_tab.pack(expand=1, fill="both")

#create a tab
tab_1 = My_tab.add("USCG")
tab_2 = My_tab.add("Terminix")
tab_3 = My_tab.add("UPS")
tab_4 = My_tab.add("Eagle")

## BACKGROUNDS___________________________________________________________________________________________
# background image for tab_1 / USCG
USCG_image = customtkinter.CTkImage(light_image=Image.open('assets/background.png'), dark_image=Image.open('assets/background.png'),
                                  size=(300, 300))

USCG_label = customtkinter.CTkLabel(tab_1, text="", image=USCG_image)
USCG_label.pack(side='top', pady=20)

# Background image for tab_2 / Terminix
terminix_image = customtkinter.CTkImage(light_image=Image.open('assets/terminix.jpg'), dark_image=Image.open('assets/terminix.jpg'),
                                  size=(450, 200))

terminix_label = customtkinter.CTkLabel(tab_2, text="", image=terminix_image)
terminix_label.pack(side='top', pady=20)

# Legend for tab_2 / Terminix
terminix_legend = customtkinter.CTkLabel(tab_2, text="standard = ground w/return lbl\n 2day = FedEx 2 Day w/return lbl\n"
                                                     "overnight = overnight priority w/return lbl", font=("Helvetica", 16))
terminix_legend.pack(side='top', pady=20)

## Background image for tab_3 / UPS
UPS_image = customtkinter.CTkImage(light_image=Image.open('assets/ups.png'), dark_image=Image.open('assets/ups.png'),
                                    size=(300, 175))

UPS_label = customtkinter.CTkLabel(tab_3, text="", image=UPS_image)
UPS_label.pack(side='top', pady=10)

## Background image for tab_4 / Eagle
Eagle_image = customtkinter.CTkImage(light_image=Image.open('assets/eagle.jpg'), dark_image=Image.open('assets/eagle.jpg'),
                                    size=(400, 400))

Eagle_label = customtkinter.CTkLabel(tab_4, text="", image=Eagle_image)
Eagle_label.pack(side='top', pady=10)



# Create a button that calls the convert_csv function for USCG_______________________________
convert_button = customtkinter.CTkButton(tab_1,
                                         text="Convert CSV", command=USCG_convert_button_click,
                                         border_width=2, border_color="gold")
convert_button.pack(side='bottom', pady=20)

# Create a button that calls the convert_csv function for Terminix_____________________________
Convert_button_terminix = customtkinter.CTkButton(tab_2,
                                                  text="Convert CSV", command=Terminix_convert_button_click,
                                                 border_width=2, border_color="Red", fg_color="green")
Convert_button_terminix.pack(side='bottom', pady=20)

# create a drop down menu for shipping accounts
Third_party_shipping = customtkinter.CTkComboBox(tab_3, values=["Select account", "Premier", "Strivr", "Bank Of America"])
Third_party_shipping.pack(side='top', pady=10)

# create an open text box for the user to enter the weight of the package
Weight = customtkinter.CTkEntry(tab_3,placeholder_text="Weight of packages in lbs", width=180)
Weight.pack(side='top', pady=10)

# create an open text box for the user to enter the length of the package
Length = customtkinter.CTkEntry(tab_3,placeholder_text="Length of packages in inches", width=180)
Length.pack(side='top', pady=10)

# create an open text box for the user to enter the width of the package
Width = customtkinter.CTkEntry(tab_3,placeholder_text="Width of packages in inches", width=180)
Width.pack(side='top', pady=10)

# create an open text box for the user to enter the height of the package
Height = customtkinter.CTkEntry(tab_3,placeholder_text="Height of packages in inches", width=180)
Height.pack(side='top', pady=10)

# Create a button that calls the convert_csv function for UPS batch file conversion________________________
Convert_button_UPS = customtkinter.CTkButton(tab_3,
                                                  text="Convert CSV", command=Terminix_convert_button_click,
                                                 border_width=2, border_color="#FFB500", fg_color="#351C15")
Convert_button_UPS.pack(side='bottom', pady=20)

# Create a button that calls the convert_csv function for Eagle csv filtering________________________
Convert_button_Eagle = customtkinter.CTkButton(tab_4,
                                                  text="Filter and save CSV", command=Eagle_button_click,
                                                 border_width=2)
Convert_button_Eagle.pack(side='bottom', pady=20)






def main():
    # Create a custom Tkinter window
    root.mainloop()


if __name__ == "__main__":
    main()