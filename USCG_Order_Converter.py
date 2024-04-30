# This order converter app imports CSV files with Kit ID's and converts the order to a SKU based CSV upload.
# Created by Dusty Baker December 2023
#Updated by Dusty Baker March 2024 to add Terminix order converter
import customtkinter
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from PIL import Image



customtkinter.set_appearance_mode("system")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green

# Initialize error_added at the top level of your script
error_added = False




def read_csv_to_dict(input_file):
    data = []
    with open(input_file, 'r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for row in csv_reader:
            data.append(row)
    return data


def manipulate_data(input_data):
    manipulated_data = []
    global error_added

    for row in input_data:
        # Check if the row is the header row and add it to the manipulated data
        if row['12'] == 'Item SKU':
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
            new_row_1['24'] = 'Dusty Baker'
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






        else:

            # If row 12 does not match the specified values, add an "Error" row

            if row['12'] not in ['SO-001', 'SO-002', 'SO-003', 'SO-004',

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

                                 'SVDI-001', 'SVDI-002', 'SVDI-003', 'SVDI-004'

                                 ]:

                error_row = row.copy()

                error_row['19'] = 'Error'

                manipulated_data.append(error_row)
                # Add this line to make error_added a global variable
                error_added = True

            else:

                manipulated_data.append(row)

    return manipulated_data


def write_csv_from_dict(output_file, output_data, fieldnames):
    with open(output_file, 'w', newline='') as csv_file:
        csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        csv_writer.writeheader()
        csv_writer.writerows(output_data)


def convert_csv(input_path, output_path):
    # Read data from the input CSV file into a dictionary
    input_data = read_csv_to_dict(input_path)

    # Manipulate the data
    manipulated_data = manipulate_data(input_data)

    # Get the field names from the input data
    fieldnames = input_data[0].keys() if input_data else []

    # Write the manipulated data to a new CSV file
    write_csv_from_dict(output_path, manipulated_data, fieldnames)

# create a function that creates a tkinter button that calls the convert_csv function
def on_convert_button_click():
    global error_added  # Declare error_added as a global variable
    # Get the input and output file paths
    input_path = filedialog.askopenfilename(title="Select Input File for USCG", filetypes=[("CSV Files", "*.csv")])
    output_path = filedialog.asksaveasfilename(title="Select Output File", defaultextension=".csv",
                                               filetypes=[("CSV Files", "*.csv")])

    # Convert the CSV file
    convert_csv(input_path, output_path)

    # Show a success message
    messagebox.showinfo(title="Semper Paratus", message="Just anther step in the long climb to Liberty!")

    # if error_added is True, show an error message
    if error_added:
        messagebox.showinfo(title="Error", message="Errors have been detected. Please review the CSV file.")

    # Reset error_added to False
    error_added = False

# create a function that creates a tkinter button that calls the convert_Terminix function


root = customtkinter.CTk()
root.title("USCG Order Converter")
root.geometry("500x500")
root.iconbitmap('images/Lambda.ico')

# Create label
Header_Label1 = customtkinter.CTkLabel(root, text="Convert your CSV file for Mantis upload")
Header_Label1.pack(pady=5)

# background image for tab_1
my_image = customtkinter.CTkImage(light_image=Image.open('images/background.png'), dark_image=Image.open('images/background.png'),
                                  size=(300, 300))

my_label = customtkinter.CTkLabel(root, text="", image=my_image)
my_label.pack(side='top', pady=20)

# Create a button that calls the convert_csv function for USCG
convert_button = customtkinter.CTkButton(root,
                                         text="Convert CSV", command=on_convert_button_click,
                                         border_width=2, border_color="gold")
convert_button.pack(side='bottom', pady=20)

def main():
    # Create a custom Tkinter window
    root.mainloop()


if __name__ == "__main__":
    main()