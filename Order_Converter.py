# This order converter app imports CSV files with Kit ID's and converts the order to a SKU based CSV upload.
# Created by Dusty Baker December 2023
#Updated by Dusty Baker March 2024 to add Terminix order converter
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


def check_additional_commas_within_strings(input_file):
    with open(input_file, 'r') as csv_file:
        reader = csv.reader(csv_file)
        header = next(reader)
        errors = []

        for line_num, row in enumerate(reader, start=2):  # Start from 2 to account for the header row
            for col_num, value in enumerate(row):
                # Check if there are commas within the string
                if ',' in value:
                    errors.append(f"Error: Comma found in row {line_num}, column {col_num + 1}")

        return errors

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


# create a function that shows an error message with an image
def show_error_with_image():
    # Create a new top-level window
    error_window = customtkinter.CTkToplevel()
    error_window.attributes('-topmost', True)
    error_window.title("Error")
    error_window.geometry("500x500")
    error_image = customtkinter.CTkImage(light_image=Image.open('images/anthonymod.jpg'), dark_image=Image.open('images/anthonymod.jpg'),
                                    size=(300, 300))
    error_label = customtkinter.CTkLabel(error_window, text="", image=error_image)
    error_label.pack(side='top', pady=20)
    error_message = customtkinter.CTkLabel(error_window, text="Womp Womp...Errors have been detected. \nPlease review the CSV file.")
    error_message.pack(side='top', pady=20)


# create a function that creates a tkinter button that calls the convert_csv function
def USCG_convert_button_click():
    global error_added  # Declare error_added as a global variable
    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File for USCG", filetypes=[("CSV Files", "*.csv")])

    # Check for additional commas within strings
    errors = check_additional_commas_within_strings(input_path)
    if errors:
        messagebox.showerror(title="CSV Error", message="\n".join(errors))
        return

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

    # Check for additional commas within strings
    errors = check_additional_commas_within_strings(input_path)
    if errors:
        messagebox.showerror(title="CSV Error", message="\n".join(errors))
        return

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
        messagebox.showinfo(title="Sweet Liberty!", message=f"File saved successfully at {output_base_path}")










## GUI_______________________________________________________________________________________________________
root = customtkinter.CTk()
root.title("Dusty's Order Converter")
root.geometry("500x500")
root.iconbitmap('images/Lambda.ico')

# Create label
Header_Label1 = customtkinter.CTkLabel(root, text="Version 1.3, Now with Shipping options!")
Header_Label1.pack(pady=5)

#create a tab view with custom tkinter
My_tab = customtkinter.CTkTabview(root)
My_tab.pack(expand=1, fill="both")

#create a tab
tab_1 = My_tab.add("USCG")
tab_2 = My_tab.add("Terminix")
tab_3 = My_tab.add("Shipping")

## BACKGROUNDS___________________________________________________________________________________________
# background image for tab_1 / USCG
USCG_image = customtkinter.CTkImage(light_image=Image.open('images/background.png'), dark_image=Image.open('images/background.png'),
                                  size=(300, 300))

USCG_label = customtkinter.CTkLabel(tab_1, text="", image=USCG_image)
USCG_label.pack(side='top', pady=20)

# Background image for tab_2 / Terminix
terminix_image = customtkinter.CTkImage(light_image=Image.open('images/terminix.jpg'), dark_image=Image.open('images/terminix.jpg'),
                                  size=(450, 200))

terminix_label = customtkinter.CTkLabel(tab_2, text="", image=terminix_image)
terminix_label.pack(side='top', pady=20)

# Legend for tab_2 / Terminix
terminix_legend = customtkinter.CTkLabel(tab_2, text="standard = ground w/return lbl\n 2day = FedEx 2 Day w/return lbl\n"
                                                     "overnight = overnight priority w/return lbl", font=("Helvetica", 16))
terminix_legend.pack(side='top', pady=20)

## Background image for tab_3 / Shipping
UPS_image = customtkinter.CTkImage(light_image=Image.open('images/ups.png'), dark_image=Image.open('images/ups.png'),
                                    size=(200, 125))

UPS_label = customtkinter.CTkLabel(tab_3, text="", image=UPS_image)
UPS_label.place(x=20, y=100)

FedEx_image = customtkinter.CTkImage(light_image=Image.open('images/fedex.png'), dark_image=Image.open('images/fedex.png'),
                                    size=(200, 100))

FedEx_label = customtkinter.CTkLabel(tab_3, text="", image=FedEx_image)
FedEx_label.place(x=250, y=100)


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

# Create a button that calls the convert_csv function for UPS batch file conversion________________________
Convert_button_UPS = customtkinter.CTkButton(tab_3,
                                                  text="Convert CSV", command=Terminix_convert_button_click,
                                                 border_width=2, border_color="#FFB500", fg_color="#351C15")
Convert_button_UPS.place(x=50, y=300)

# Create a button that calls the convert_csv function for FedEx batch file conversion________________________
Convert_button_FedEx = customtkinter.CTkButton(tab_3,
                                                  text="Convert CSV", command=Terminix_convert_button_click,
                                                 border_width=2, border_color="#FF6600", fg_color="#4D148C")
Convert_button_FedEx.place(x=280, y=300)




def main():
    # Create a custom Tkinter window
    root.mainloop()


if __name__ == "__main__":
    main()