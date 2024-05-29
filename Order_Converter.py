# This order converter app imports CSV files with Kit ID's and converts the order to a SKU based CSV upload.
# Created by Dusty Baker December 2023
# Updated by Dusty Baker May 2024 to add UPS shipping options, and EAGLE file manager.
import csv
import os
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
import customtkinter
import pandas as pd
from PIL import Image

customtkinter.set_appearance_mode("system")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green

# Initialize error_added variable to keep track of whether an error message has been added
error_added = False

errors = []



# Try reading the CSV files with 'latin1' encoding first, then try 'ISO-8859-1' if it fails
try:
    ups_batch = pd.read_csv('assets/ups_batch.csv', encoding='latin1')
    projects = pd.read_csv('assets/projects.csv', encoding='latin1')
    uscg_template = pd.read_csv('assets/USCG_data.csv', encoding='latin1')
    terminix_template = pd.read_csv('assets/terminix_data.csv', encoding='latin1')
except UnicodeDecodeError:
    print("Error reading the file with 'latin1' encoding. Trying 'ISO-8859-1'...")
    ups_batch = pd.read_csv('assets/ups_batch.csv', encoding='ISO-8859-1')
    projects = pd.read_csv('assets/projects.csv', encoding='ISO-8859-1')
    uscg_template = pd.read_csv('assets/USCG_data.csv', encoding='ISO-8859-1')
    terminix_template = pd.read_csv('assets/terminix_data.csv', encoding='ISO-8859-1')

# Functinon to define the maximum number of characters in a column
allowable_lengths = {
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


# Function to manipulate the input uscg data using the template
def manipulate_uscg_data_using_template(input_data, template):
    manipulated_data = pd.DataFrame()

    for _, row in input_data.iterrows():
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

# Function to manipulate the input terminix data using the template
def manipulate_terminix_data_using_template(input_data, template):
    manipulated_data = pd.DataFrame()

    for _, row in input_data.iterrows():
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
def check_character_length(df, length_dict, errors):
    for column, max_length in length_dict.items():
        if column in df.columns:
            # Check if any cell in the column exceeds the maximum allowable length
            exceeds_length = df[column].astype(str).apply(len) > max_length
            if exceeds_length.any():
                # Add an error with the index and value of the cell(s) that exceed the length
                error_rows = df[exceeds_length].index.tolist()
                error_values = df.loc[error_rows, column].tolist()
                for row, value in zip(error_rows, error_values):
                    errors.append(f"Column '{column}' row {row + 2} exceeds allowable length of {max_length} characters."
                                  f"\nError value: '{value}'")


def USCG_Error_Handling(input_file):
    # Read the CSV file into a DataFrame
    df = pd.read_csv(input_file)

    # Define the valid kit IDs based on the template
    valid_kit_ids = uscg_template['Kit ID'].unique()

    # Initialize an empty list to store errors
    errors = []

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        # Check if the row has more than 11 columns and if the 12th column value is not in valid_kit_ids
        if len(row) > 11 and row.iloc[11] not in valid_kit_ids:
            errors.append(f"Error: Incorrect Kit ID found in row {index + 2}")  # Adding 2 to match the original line numbering

    return errors


# Functino to check for errors in the Terminix CSV file
def Terminix_error_handling(input_file):
    # Read the CSV file into a DataFrame
    df = pd.read_csv(input_file)

    # Define the valid kit IDs from the template
    valid_kit_ids = terminix_template['Kit ID'].unique()

    # Initialize an empty list to store errors
    errors = []

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        # Check if the row has more than 11 columns and if the 12th column value is not in valid_kit_ids
        if len(row) > 11 and row.iloc[11] not in valid_kit_ids:
            errors.append(f"Error: Incorrect Kit ID found in row {index + 2}")  # Adding 2 to match the original line numbering

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
    check_character_length(input_data, allowable_lengths, errors)

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
    check_character_length(input_data, allowable_lengths, errors)

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
        messagebox.showerror(title="CSV Error", message="\n".join(errors))
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
        check_character_length(input_data, allowable_lengths, errors)

        # If there are errors, raise a ValueError
        if errors:
            raise ValueError("\n".join(errors))
    except (OSError, IOError, ValueError) as e:
        messagebox.showerror(title="Error", message=str(e))
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
            title="Select Output Directory and Base Filename",
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
        messagebox.showinfo(title="Semper Paratus", message=f"File saved successfully at {output_base_path}")
    except (OSError, IOError, ValueError) as e:
        messagebox.showerror(title="Error", message=str(e))




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
        check_character_length(input_data, allowable_lengths, errors)

        # If there are errors, raise a ValueError
        if errors:
            raise ValueError("\n".join(errors))

        # Check for correct Terminix kit ID
        errors.extend(Terminix_error_handling(input_path))
        if errors:
            raise ValueError("\n".join(errors))
    except (OSError, IOError, ValueError) as e:
        messagebox.showerror(title="Error", message=str(e))
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
            title="Select Output Directory and Base Filename",
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
        messagebox.showinfo(title="Sweet Liberty!", message=f"File saved successfully at {output_base_path}")
    except (OSError, IOError, ValueError) as e:
        messagebox.showerror(title="Error", message=str(e))





# Create a functin that converts the input file to UPS format
def UPS_convert_button_click():
    global error_added  # Declare error_added as a global variable
    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File for UPS", filetypes=[("CSV Files", "*.csv")])

    # Clean the input file to remove additional commas
    write_cleaned_csv(input_path, check_and_remove_additional_commas(input_path))

    # Check for correct UPS kit ID






def Eagle_button_click():
    # Get the input file path
    input_path = filedialog.askopenfilename(title="Select Input File", filetypes=[("CSV Files", "*.csv")])

    if not input_path:
        return  # User cancelled the file dialog

    # Clean the input file to remove additional commas
    cleaned_data = check_and_remove_additional_commas(input_path)

    # Initialize errors list
    errors = []

    # Check the character length of the columns in the cleaned data
    check_character_length(cleaned_data, allowable_lengths, errors)

    # If there are errors, stop execution
    if errors:
        messagebox.showerror(title="Character Length Error", message="\n".join(errors))
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
        messagebox.showerror(title="Error", message=f"Project Number '{project_number}' not found.")
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
        messagebox.showinfo(title="An Eagle never misses", message=f"File saved successfully at {output_path}\n"
                                                    "This ultimate dad energy is brought to you by Dusty Baker.")

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
        write_cleaned_csv(output_base_path, cleaned_data)
        messagebox.showinfo(title="An Eagle never misses", message=f"File saved successfully at {output_base_path}\n"
                                                        "This ultimate dad energy is brought to you by Dusty Baker.")





# GUI_______________________________________________________________________________________________________
root = customtkinter.CTk()
root.title("Dusty's Order Converter")
root.geometry("600x600")
root.iconbitmap('assets/Lambda.ico')

# Create label
Header_Label1 = customtkinter.CTkLabel(root, text="Version 1.5,\nNow powered by Pandas!, but like super fast ones that do crossfit.")
Header_Label1.pack(pady=5)

# create a tab view with custom tkinter
My_tab = customtkinter.CTkTabview(root)
My_tab.pack(expand=1, fill="both")

# create a tab
tab_1 = My_tab.add("USCG")
tab_2 = My_tab.add("Terminix")
tab_3 = My_tab.add("UPS")
tab_4 = My_tab.add("Eagle")

# BACKGROUNDS___________________________________________________________________________________________
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
                                                  text="Convert CSV", command=UPS_convert_button_click,
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