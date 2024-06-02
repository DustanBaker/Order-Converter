# This order converter app imports CSV files with Kit ID's and converts the order to a SKU based CSV upload.
# Created by Dusty Baker December 2023
# Updated by Dusty Baker May 2024 to add UPS shipping options, and EAGLE file manager.
import csv
import importlib
import os
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
import customtkinter
import pandas as pd
from PIL import Image
import pygame
import pyi_splash

pyi_splash.update_text("PyInstaller is a great software!")
pyi_splash.update_text("Second time's a charm!")

    # Close the splash screen. It does not matter when the call
    # to this function is made, the splash screen remains open until
    # this function is called or the Python program is terminated.
pyi_splash.close()

customtkinter.set_appearance_mode("system")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green

# Initialize pygame
pygame.mixer.init()

# Load the sound file
notification = pygame.mixer.Sound('assets/Propaganda.mp3')
intro = pygame.mixer.Sound('assets/LiberTea.mp3')
Error = pygame.mixer.Sound('assets/arm.mp3')

# Initialize error_added variable to keep track of whether an error message has been added
error_added = False

errors = []

# Function to check if files exist
def check_files_exist(file_paths):
    missing_files = [file for file in file_paths if not os.path.isfile(file)]
    if missing_files:
        messagebox.showerror("Missing Files", f"The following files are missing: {', '.join(missing_files)}")
        return False
    return True

# List of required files
required_files = [
    'assets/ups_batch.csv',
    'assets/projects.csv',
    'assets/USCG_data.csv',
    'assets/terminix_data.csv'
]

# Check if required files exist
if check_files_exist(required_files):
    # Try reading the CSV files with 'latin1' encoding first, then try 'ISO-8859-1' if it fails
    try:
        ups_batch = pd.read_csv('assets/ups_batch.csv', encoding='latin1')
        projects = pd.read_csv('assets/projects.csv', encoding='latin1')
        uscg_template = pd.read_csv('assets/USCG_data.csv', encoding='latin1')
        terminix_template = pd.read_csv('assets/terminix_data.csv', encoding='latin1')
    except UnicodeDecodeError:
        messagebox.showerror("Encoding Error", "Error reading the file with 'latin1' encoding. Trying 'ISO-8859-1'...")
        try:
            ups_batch = pd.read_csv('assets/ups_batch.csv', encoding='ISO-8859-1')
            projects = pd.read_csv('assets/projects.csv', encoding='ISO-8859-1')
            uscg_template = pd.read_csv('assets/USCG_data.csv', encoding='ISO-8859-1')
            terminix_template = pd.read_csv('assets/terminix_data.csv', encoding='ISO-8859-1')
        except Exception as e:
            messagebox.showerror("File Error", f"An error occurred while reading the files: {e}")
else:
    # Exit the application if files are missing
    exit()

# Define the maximum number of characters in a column for shipment and WO inputs
allowable_lengths_for_shipments = {
    '4': 70,
    '5': 30,
    '6': 64,
    '7': 64,
    '8': 32,
    '9': 32,
    '10': 32,
    '11': 32,
    '12': 20,
    '14': 64,
    '15': 64,
    '16': 64,
    '17': 64,
    '18': 64,
    '19': 255,
    '21': 64,
    '22': 11,
    '23': 64,
    '24': 64,
    '25': 64,
    '26': 64,
    '27': 11,
    '28': 64,
    '29': 64,
    '30': 64,
    '31': 64
}

# Define the maximum allowable lengths for ASN templates inputs
allowable_lengths_for_ASN = {
    '1': 20,
    '2': 64,
    '3': 20,
    '4': 64,
    '5': 64,
    '6': 20,
    '9': 64,
    '10': 64,
    '11': 64,
    '12': 64,
    '13': 64,
    '14': 64,
    '15': 64,
    '16': 64,
    '17': 64,

}

# Extract the data from column 31
uscg_template['31'] = uscg_template.iloc[:, 7]  # Assuming column index is zero-based


# Function to manipulate the input uscg data using the template
def manipulate_uscg_data_using_template(input_data, template):
    # Copy the first row under the header without manipulation
    manipulated_data = pd.DataFrame([input_data.iloc[0]])

    # Iterate over the rest of the input data starting from the second row
    for _, row in input_data.iloc[1:].iterrows():
        kit_id = row['12']
        matching_template_rows = template[template['Kit ID'] == kit_id]

        if matching_template_rows.empty:
            manipulated_data = pd.concat([manipulated_data, pd.DataFrame([row])])
        else:
            for _, template_row in matching_template_rows.iterrows():
                new_row = row.copy()

                # Update new_row with values from template_row where not NaN
                for column in template_row.index:
                    if column in row and pd.notna(template_row[column]):
                        new_row[column] = template_row[column]

                # Prepend the column 31 value to the existing data in the input row
                if '31' in new_row and '31' in template_row:
                    new_row['31'] = f"{template_row['31']}{row['31']}"

                manipulated_data = pd.concat([manipulated_data, pd.DataFrame([new_row])])

    return manipulated_data


# Function to manipulate the input terminix data using the template
def manipulate_terminix_data_using_template(input_data, template):
    # Copy the first row under the header without manipulation
    manipulated_data = pd.DataFrame([input_data.iloc[0]])

    # Iterate over the rest of the input data starting from the second row
    for _, row in input_data.iloc[1:].iterrows():
        kit_id = row['12']
        matching_template_rows = template[template['Kit ID'] == kit_id]

        if matching_template_rows.empty:
            manipulated_data = pd.concat([manipulated_data, pd.DataFrame([row])])
        else:
            for _, template_row in matching_template_rows.iterrows():
                new_row = row.copy()
                for column in template_row.index:
                    if column in row and pd.notna(template_row[column]):
                        new_row[column] = template_row[column]
                manipulated_data = pd.concat([manipulated_data, pd.DataFrame([new_row])])

    return manipulated_data


# create a function that reads a CSV file into a dictionary and removes additional commas
def check_and_remove_additional_commas(input_file):
    # Read the CSV file into a DataFrame
    df = pd.read_csv(input_file)

    # Remove additional commas from all string fields
    df = df.map(lambda x: x.replace(',', '') if isinstance(x, str) else x)

    return df


# create a function that writes a cleaned CSV file
def write_cleaned_csv(output_file, cleaned_data):
    # Ensure cleaned_data is a DataFrame
    if not isinstance(cleaned_data, pd.DataFrame):
        cleaned_data = pd.DataFrame(cleaned_data)

    # Write the DataFrame to a CSV file
    cleaned_data.to_csv(output_file, index=False)


# Function to check the length of characters in a column
def check_character_length_shipment(df, length_dict, errors):
    for column, max_length in length_dict.items():
        if column in df.columns:
            # Check if any cell in the column exceeds the maximum allowable length
            exceeds_length = df[column].astype(str).apply(len) > max_length
            if exceeds_length.any():
                # Add an error with the index and value of the cell(s) that exceed the length
                error_rows = df[exceeds_length].index.tolist()
                error_values = df.loc[error_rows, column].tolist()
                for row, value in zip(error_rows, error_values):
                    errors.append(
                        f"Column '{column}' row {row + 2} exceeds allowable length of {max_length} characters."
                        f"\nError value: '{value}'")


# Function to check for the length of characters in a column for ASN's
def check_character_length_ASN(df, length_dict, errors):
    for column, max_length in length_dict.items():
        if column in df.columns:
            # Check if any cell in the column exceeds the maximum allowable length
            exceeds_length = df[column].astype(str).apply(len) > max_length
            if exceeds_length.any():
                # Add an error with the index and value of the cell(s) that exceed the length
                error_rows = df[exceeds_length].index.tolist()
                error_values = df.loc[error_rows, column].tolist()
                for row, value in zip(error_rows, error_values):
                    errors.append(
                        f"Column '{column}' row {row + 2} exceeds allowable length of {max_length} characters."
                        f"\nError value: '{value}'")


def USCG_Error_Handling(input_file):
    # Read the CSV file into a DataFrame, skipping the first row after the header
    df = pd.read_csv(input_file, skiprows=[1])

    # Define the valid kit IDs based on the template
    valid_kit_ids = uscg_template['Kit ID'].unique()

    # Initialize an empty list to store errors
    errors = []

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        # Check if the row has more than 11 columns and if the 12th column value is not in valid_kit_ids
        if len(row) > 11 and row.iloc[11] not in valid_kit_ids:
            errors.append(
                f"Error: Incorrect Kit ID found in row {index + 3}")  # Adding 3 to match the original line numbering

    return errors


# Functino to check for errors in the Terminix CSV file
def Terminix_error_handling(input_file):
    # Read the CSV file into a DataFrame, skipping the first row after the header
    df = pd.read_csv(input_file, skiprows=[1])

    # Define the valid kit IDs from the template
    valid_kit_ids = terminix_template['Kit ID'].unique()

    # Initialize an empty list to store errors
    errors = []

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        # Check if the row has more than 11 columns and if the 12th column value is not in valid_kit_ids
        if len(row) > 11 and row.iloc[11] not in valid_kit_ids:
            errors.append(
                f"Error: Incorrect Kit ID found in row {index + 3}")  # Adding 3 to match the original line numbering

    return errors


# Function to convert USCG CSV using pandas and the template
def convert_USCG_csv(input_path, output_path, template):
    # Clean the commas from the input file
    cleaned_data = check_and_remove_additional_commas(input_path)
    write_cleaned_csv(input_path, cleaned_data)

    # Read the CSV file into the pandas DataFrame
    input_data = pd.read_csv(input_path)

    # Initialize errors list
    errors = []

    # Check the character length of the columns in the cleaned data
    check_character_length_shipment(input_data, allowable_lengths_for_shipments, errors)

    # If there are errors, raise a ValueError
    if errors:
        raise ValueError("\n".join(errors))

    # Manipulate the data using the template
    manipulated_data = manipulate_uscg_data_using_template(input_data, template)

    # Write the manipulated data to a new CSV file
    manipulated_data.to_csv(output_path, index=False)


def convert_Terminix_csv(input_path, output_path, template):
    # Clean the commas from the input file
    cleaned_data = check_and_remove_additional_commas(input_path)
    write_cleaned_csv(input_path, cleaned_data)

    # Read the CSV file into the pandas DataFrame
    input_data = pd.read_csv(input_path)

    # Initialize errors list
    errors = []

    # Check the character length of the columns in the cleaned data
    check_character_length_shipment(input_data, allowable_lengths_for_shipments, errors)

    # If there are errors, raise a ValueError
    if errors:
        raise ValueError("\n".join(errors))

    # Manipulate the data using the template
    manipulated_data = manipulate_terminix_data_using_template(input_data, template)

    # Write the manipulated data to a new CSV file
    manipulated_data.to_csv(output_path, index=False)


# Create a function that creates a tkinter button that calls the convert_csv function for USCG
def USCG_convert_button_click():
    global error_added
    input_path = filedialog.askopenfilename(title="Select Input File for USCG", filetypes=[("CSV Files", "*.csv")])

    if not input_path:
        return  # User cancelled the file dialog

    # Check for correct USCG kit ID
    errors = USCG_Error_Handling(input_path)
    if errors:
        Error_window("Error",f"Error: {errors}")
        return

    try:
        # Clean the commas from the input file
        cleaned_data = check_and_remove_additional_commas(input_path)
        write_cleaned_csv(input_path, cleaned_data)

        # Read the cleaned data into a DataFrame
        input_data = pd.read_csv(input_path)

        # Initialize errors list
        errors = []

        # Check the character length of the columns in the cleaned data
        check_character_length_shipment(input_data, allowable_lengths_for_shipments, errors)

        # If there are errors, raise a ValueError
        if errors:
            raise ValueError("\n".join(errors))
    except (OSError, IOError, ValueError) as e:
        Error_window( "Error",f"Error: {e}")
        return

    # Create the default output file path
    current_date = datetime.now().strftime("%m-%d-%Y")
    default_directory = r"T:\3PL Files\Shipment Template"
    default_filename = f"USCG Shipment {current_date}.csv"
    default_output_path = os.path.join(default_directory, default_filename)

    # Check if the default directory exists
    if not os.path.exists(default_directory):
        # Prompt the user to choose a file path if the default path is not available
        output_base_path = filedialog.asksaveasfilename(
            title="Save output file",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=default_filename
        )
        if not output_base_path:
            return  # User cancelled the save dialog
    else:
        output_base_path = default_output_path

    try:
        convert_USCG_csv(input_path, output_base_path, uscg_template)
        Success_window(f"File saved successfully at\n{output_base_path}")
    except (OSError, IOError, ValueError) as e:
        Error_window("Error", f"Error: {e}" )


def Terminix_convert_button_click():
    global error_added
    input_path = filedialog.askopenfilename(title="Select Input File for Terminix", filetypes=[("CSV Files", "*.csv")])

    if not input_path:
        return  # User cancelled the file dialog

    try:
        # Clean the commas from the input file
        cleaned_data = check_and_remove_additional_commas(input_path)
        write_cleaned_csv(input_path, cleaned_data)

        # Read the cleaned data into a DataFrame
        input_data = pd.read_csv(input_path)

        # Initialize errors list
        errors = []

        # Check the character length of the columns in the cleaned data
        check_character_length_shipment(input_data, allowable_lengths_for_shipments, errors)

        # If there are errors, raise a ValueError
        if errors:
            raise ValueError("\n".join(errors))

        # Check for correct Terminix kit ID
        errors.extend(Terminix_error_handling(input_path))
        if errors:
            raise ValueError("\n".join(errors))
    except (OSError, IOError, ValueError) as e:
        Error_window("Error",f"Error: {e}")
        return

    # Create the default output file path
    current_date = datetime.now().strftime("%m-%d-%Y")
    default_directory = r"T:\3PL Files\Shipment Template"
    default_filename = f"Terminix Shipment {current_date}.csv"
    default_output_path = os.path.join(default_directory, default_filename)

    # Check if the default directory exists
    if not os.path.exists(default_directory):
        # Prompt the user to choose a file path if the default path is not available
        output_base_path = filedialog.asksaveasfilename(
            title="Save output file",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=default_filename
        )
        if not output_base_path:
            return  # User cancelled the save dialog
    else:
        output_base_path = default_output_path

    try:
        convert_Terminix_csv(input_path, output_base_path, terminix_template)
        Success_window(f"File saved successfully at\n{output_base_path}")
    except (OSError, IOError, ValueError) as e:
        Error_window("Error", f"Error: {e}")


# Create a functin that converts the input file to UPS format
def UPS_convert_button_click():
    global error_added  # Declare error_added as a global variable
    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File for UPS", filetypes=[("CSV Files", "*.csv")])

    # Clean the input file to remove additional commas
    write_cleaned_csv(input_path, check_and_remove_additional_commas(input_path))

    # Check for correct UPS kit ID


# functinon to filter and save the input file as a shipment
def Eagle_shipment_button_click():
    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File", filetypes=[("CSV Files", "*.csv")])

    if not input_path:
        return  # User cancelled the file dialog

    # Clean the input file to remove additional commas
    cleaned_data = check_and_remove_additional_commas(input_path)
    write_cleaned_csv(input_path, cleaned_data)

    # Initialize errors list
    errors = []

    # Check the character length of the columns in the cleaned data
    check_character_length_shipment(cleaned_data, allowable_lengths_for_shipments, errors)

    # If there are errors, stop execution
    if errors:
        Error_window("Character Length Error", f"Too many Characters, \n".join(errors))
        return

    # Read the third row, first column to get the project number using pandas
    project_number = pd.read_csv(input_path, header=None, nrows=3).iloc[2, 0]

    # Convert project_number to string and strip any whitespace
    project_number = str(project_number).strip()

    # Ensure no extra spaces or unexpected characters in column names
    projects.columns = projects.columns.str.strip()

    # Convert the 'Project Number' column to string and strip any whitespace
    projects['Project Number'] = projects['Project Number'].astype(str).str.strip()

    # Map the project number to the project name using project dataframe from the projects.csv file using pandas
    try:
        project_name_row = projects[projects['Project Number'] == project_number]

        if not project_name_row.empty:
            project_name = project_name_row['Project Name'].values[0]

        else:
            raise IndexError("Project number not found in DataFrame")

    except IndexError:
        Error_window("Project error", f"Project Number '{project_number}' not found.")
        return

    # Create the output file path
    current_date = datetime.now().strftime("%m-%d-%Y")
    default_directory = r"T:\3PL Files\Shipment Template"
    default_filename = f"{project_name} Shipment {current_date}.csv"
    output_path = os.path.join(default_directory, default_filename)

    # Save the cleaned CSV file with the appropriate name
    try:
        os.makedirs(default_directory, exist_ok=True)
        write_cleaned_csv(output_path, cleaned_data)
        Success_window(f"File saved successfully at \n{output_path}")

    except (OSError, IOError):
        # If the default path is unavailable, ask the user for the output directory and base filename
        output_base_path = filedialog.asksaveasfilename(
            title="Save Output File as Shipment",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=default_filename
        )
        if not output_base_path:
            return  # User cancelled the save dialog

        # Save to the selected path
        write_cleaned_csv(output_base_path, cleaned_data)
        Success_window(f"File saved successfully at\n{output_base_path}")


# Function to convert the input file to a Work Order
def Eagle_WO_button_click():
    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File", filetypes=[("CSV Files", "*.csv")])

    if not input_path:
        return  # User cancelled the file dialog

    # Clean the input file to remove additional commas
    cleaned_data = check_and_remove_additional_commas(input_path)
    write_cleaned_csv(input_path, cleaned_data)

    # Initialize errors list
    errors = []

    # Check the character length of the columns in the cleaned data
    check_character_length_shipment(cleaned_data, allowable_lengths_for_shipments, errors)

    # If there are errors, stop execution
    if errors:
        Error_window("Character Length Error", f"Too many Characters \n,".join(errors))
        return

    # Read the third row, first column to get the project number using pandas
    project_number = pd.read_csv(input_path, header=None, nrows=3).iloc[2, 0]

    # Convert project_number to string and strip any whitespace
    project_number = str(project_number).strip()

    # Ensure no extra spaces or unexpected characters in column names
    projects.columns = projects.columns.str.strip()

    # Convert the 'Project Number' column to string and strip any whitespace
    projects['Project Number'] = projects['Project Number'].astype(str).str.strip()

    # Map the project number to the project name using project dataframe from the projects.csv file using pandas
    try:
        project_name_row = projects[projects['Project Number'] == project_number]

        if not project_name_row.empty:
            project_name = project_name_row['Project Name'].values[0]

        else:
            raise IndexError("Project number not found in DataFrame")

    except IndexError:
        Error_window("Project Error",f"Project Number '{project_number}' not found.")
        return

    # Create the output file path
    current_date = datetime.now().strftime("%m-%d-%Y")
    default_directory = r"T:\3PL Files\Shipment Template"
    default_filename = f"{project_name} Work Order {current_date}.csv"
    output_path = os.path.join(default_directory, default_filename)

    # Save the cleaned CSV file with the appropriate name
    try:
        os.makedirs(default_directory, exist_ok=True)
        write_cleaned_csv(output_path, cleaned_data)
        Success_window(f"File saved successfully at\n{output_path}")

    except (OSError, IOError):
        # If the default path is unavailable, ask the user for the output directory and base filename
        output_base_path = filedialog.asksaveasfilename(
            title="Save Output File as Work Order",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=default_filename
        )
        if not output_base_path:
            return  # User cancelled the save dialog

        # Save to the selected path
        write_cleaned_csv(output_base_path, cleaned_data)
        Success_window(f"File saved successfully at\n{output_base_path}")


# Function to filter and save the input file as an ASN
def Eagle_ASN_button_click():
    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File", filetypes=[("CSV Files", "*.csv")])

    if not input_path:
        return  # User cancelled the file dialog

    # Clean the input file to remove additional commas
    cleaned_data = check_and_remove_additional_commas(input_path)
    write_cleaned_csv(input_path, cleaned_data)

    # Initialize errors list
    errors = []

    # Check the character length of the columns in the cleaned data
    check_character_length_ASN(cleaned_data, allowable_lengths_for_shipments, errors)

    # If there are errors, stop execution
    if errors:
        Error_window("Character Length Error", f"Too many Characters, \n".join(errors))
        return

    # Read the fourth row, first column to get the project number using pandas
    project_number = pd.read_csv(input_path, header=None, nrows=4).iloc[3, 0]

    # Convert project_number to string and strip any whitespace
    project_number = str(project_number).strip()

    # Ensure no extra spaces or unexpected characters in column names
    projects.columns = projects.columns.str.strip()

    # Convert the 'Project Number' column to string and strip any whitespace
    projects['Project Number'] = projects['Project Number'].astype(str).str.strip()

    # Map the project number to the project name using project dataframe from the projects.csv file using pandas
    try:
        project_name_row = projects[projects['Project Number'] == project_number]

        if not project_name_row.empty:
            project_name = project_name_row['Project Name'].values[0]

        else:
            raise IndexError("Project number not found in DataFrame")

    except IndexError:
        Error_window("Project error", f"Project Number '{project_number}' not found.")
        return

    # Create the output file path
    current_date = datetime.now().strftime("%m-%d-%Y")
    default_directory = r"T:\3PL Files\ASN Template"
    default_filename = f"{project_name} ASN {current_date}.csv"
    output_path = os.path.join(default_directory, default_filename)

    # Save the cleaned CSV file with the appropriate name
    try:
        os.makedirs(default_directory, exist_ok=True)
        write_cleaned_csv(output_path, cleaned_data)
        Success_window(f"File saved successfully at\n{output_path}")

    except (OSError, IOError):
        # If the default path is unavailable, ask the user for the output directory and base filename
        output_base_path = filedialog.asksaveasfilename(
            title="Save Output File as ASN",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=default_filename
        )
        if not output_base_path:
            return  # User cancelled the save dialog

        # Save to the selected path
        write_cleaned_csv(output_base_path, cleaned_data)
        Success_window(f"File saved successfully at\n{output_base_path}")


# GUI_______________________________________________________________________________________________________

root = customtkinter.CTk()
root.title("Eagle File Manager")
root.geometry("600x650")
root.iconbitmap('assets/Eagle.ico')
#locate the main window in the center of the screen
root.eval('tk::PlaceWindow . center')


# Create label
Header_Label1 = customtkinter.CTkLabel(root, text="Version 1.6\n Now with 50% more Eagle!")
Header_Label1.pack(pady=5)

# create a tab view with custom tkinter
My_tab = customtkinter.CTkTabview(root)
My_tab.pack(expand=1, fill="both")

# create a tab
tab_1 = My_tab.add("The Eagle")
tab_2 = My_tab.add("USCG")
tab_3 = My_tab.add("Terminix")
tab_4 = My_tab.add("UPS")

# TAB 1 / The EAGLE_________________________________________________________________________________________
## Background image for tab_1 / Eagle
Eagle_image = customtkinter.CTkImage(light_image=Image.open('assets/eagle.jpg'),
                                     dark_image=Image.open('assets/eagle.jpg'),
                                     size=(400, 400))

Eagle_label = customtkinter.CTkLabel(tab_1, text="", image=Eagle_image)
Eagle_label.pack(side='top', pady=10)

# Create a button that calls the Eagle_WO_button_click function for Eagle csv filtering and save as a Work Order
Work_order_button_Eagle = customtkinter.CTkButton(tab_1,
                                                  text="Filter and save Work Order", command=Eagle_WO_button_click,
                                                  border_width=2)
Work_order_button_Eagle.pack(side='bottom', pady=10)

# Create a button that calls the Eagle_shipment_button_click function for Eagle csv filtering and save as shipment
Shipment_button_Eagle = customtkinter.CTkButton(tab_1,
                                                text="Filter and save Shipment",
                                                command=Eagle_shipment_button_click,
                                                border_width=2)
Shipment_button_Eagle.pack(side='bottom', pady=10)

# Create a button that called the Eagle_ASN_button_click function for Eagle csv filtering and save as ASN
ASN_button_Eagle = customtkinter.CTkButton(tab_1,
                                           text="Filter and save ASN", command=Eagle_ASN_button_click,
                                           border_width=2)
ASN_button_Eagle.pack(side='bottom', pady=10)

# TAB 2 / USCG_____________________________________________________________________________________________
# background image for tab_2 / USCG
USCG_image = customtkinter.CTkImage(light_image=Image.open('assets/background.png'),
                                    dark_image=Image.open('assets/background.png'),
                                    size=(300, 300))

USCG_label = customtkinter.CTkLabel(tab_2, text="", image=USCG_image)
USCG_label.pack(side='top', pady=20)

# Create a button that calls the convert_csv function for USCG
convert_button = customtkinter.CTkButton(tab_2,
                                         text="Convert CSV", command=USCG_convert_button_click,
                                         border_width=2, border_color="gold")
convert_button.pack(side='bottom', pady=20)

# TAB 3 / Terminix_________________________________________________________________________________________
# Background image for tab_3 / Terminix
terminix_image = customtkinter.CTkImage(light_image=Image.open('assets/terminix.jpg'),
                                        dark_image=Image.open('assets/terminix.jpg'),
                                        size=(450, 200))

terminix_label = customtkinter.CTkLabel(tab_3, text="", image=terminix_image)
terminix_label.pack(side='top', pady=20)

# Legend for tab_2 / Terminix
terminix_legend = customtkinter.CTkLabel(tab_3,
                                         text="standard = ground w/return lbl\n 2day = FedEx 2 Day w/return lbl\n"
                                              "overnight = overnight priority w/return lbl", font=("Helvetica", 16))
terminix_legend.pack(side='top', pady=20)

# Create a button that calls the convert_csv function for Terminix
Convert_button_terminix = customtkinter.CTkButton(tab_3,
                                                  text="Convert CSV", command=Terminix_convert_button_click,
                                                  border_width=2, border_color="Red", fg_color="green")
Convert_button_terminix.pack(side='bottom', pady=20)

# TAB 4 / UPS_____________________________________________________________________________________________
## Background image for tab_3 / UPS
UPS_image = customtkinter.CTkImage(light_image=Image.open('assets/ups.png'),
                                   dark_image=Image.open('assets/ups.png'),
                                   size=(300, 175))

UPS_label = customtkinter.CTkLabel(tab_4, text="", image=UPS_image)
UPS_label.pack(side='top', pady=10)

# create a drop down menu for shipping accounts
Third_party_shipping = customtkinter.CTkComboBox(tab_4,
                                                 values=["Select account", "Premier", "Strivr", "Bank Of America"])
Third_party_shipping.pack(side='top', pady=10)

# create an open text box for the user to enter the weight of the package
Weight = customtkinter.CTkEntry(tab_4, placeholder_text="Weight of packages in lbs", width=180)
Weight.pack(side='top', pady=10)

# create an open text box for the user to enter the length of the package
Length = customtkinter.CTkEntry(tab_4, placeholder_text="Length of packages in inches", width=180)
Length.pack(side='top', pady=10)

# create an open text box for the user to enter the width of the package
Width = customtkinter.CTkEntry(tab_4, placeholder_text="Width of packages in inches", width=180)
Width.pack(side='top', pady=10)

# create an open text box for the user to enter the height of the package
Height = customtkinter.CTkEntry(tab_4, placeholder_text="Height of packages in inches", width=180)
Height.pack(side='top', pady=10)

# Create a button that calls the convert_csv function for UPS batch file conversion________________________
Convert_button_UPS = customtkinter.CTkButton(tab_4,
                                             text="Convert CSV", command=UPS_convert_button_click,
                                             border_width=2, border_color="#FFB500", fg_color="#351C15")
Convert_button_UPS.pack(side='bottom', pady=20)




# custom Tkinter top level window for success message
def Success_window(message):
    new_window = customtkinter.CTkToplevel()
    new_window.title("An Eagle never misses")
    new_window.geometry("500x200")
    new_window.attributes('-topmost', True)
    new_window.iconbitmap('assets/Eagle.ico')

    # Show the eagle image
    Eagle_image = customtkinter.CTkImage(light_image=Image.open('assets/eagle.ico'),
                                         dark_image=Image.open('assets/eagle.ico'),
                                         size=(100, 100))

    # Create a frame to hold the image and text
    content_frame = customtkinter.CTkFrame(new_window)
    content_frame.pack(expand=True, fill='both', padx=10, pady=10)

    # Place the eagle image on the left side
    Eagle_label = customtkinter.CTkLabel(content_frame, text="", image=Eagle_image)
    Eagle_label.pack(side='left', padx=10, pady=10)

    # Create a label for the success message, placed to the right of the image
    success_label = customtkinter.CTkLabel(content_frame, text=message, wraplength=300)
    success_label.pack(side='left', padx=10, pady=10, expand=True, fill='both')

    # Create a button centered at the bottom
    Create_button = customtkinter.CTkButton(new_window, text="Lets GO!", command=new_window.destroy)
    Create_button.pack(side='bottom', pady=20)

    # Play the notification sound
    notification.play()

    # Destroy the window after 10 seconds
    new_window.after(10000, new_window.destroy)


#custom Tkinter top level window for error message
def Error_window(title, message):
    new_window = customtkinter.CTkToplevel(root)
    new_window.title(title)
    new_window.geometry("500x200")
    new_window.iconbitmap('assets/Eagle.ico')  # Set the icon for the error window
    new_window.attributes('-topmost', True)

    # Show the eagle image
    Eagle_image = customtkinter.CTkImage(light_image=Image.open('assets/eagle.ico'),
                                         dark_image=Image.open('assets/eagle.ico'),
                                         size=(100, 100))

    # Create a frame to hold the image and text
    content_frame = customtkinter.CTkFrame(new_window)
    content_frame.pack(expand=True, fill='both', padx=10, pady=10)

    # Place the eagle image on the left side
    Eagle_label = customtkinter.CTkLabel(content_frame, text="", image=Eagle_image)
    Eagle_label.pack(side='left', padx=10, pady=10)

    # Create a label for the error message, placed to the right of the image
    error_label = customtkinter.CTkLabel(content_frame, text=message, wraplength=300)
    error_label.pack(side='left', padx=10, pady=10, expand=True, fill='both')

    # Create a button centered at the bottom
    Create_button = customtkinter.CTkButton(new_window, text="Disappointment", command=new_window.destroy)
    Create_button.pack(side='bottom', pady=20)

    # Play the error sound
    Error.play()

def main():
    root.mainloop()




if __name__ == "__main__":
    main()