# Order-Converter
Tool for converting CSV files

This program was developed by Dusty Baker 2024
The program is written in Python using the Tkinter and CustomTkinter GUI libraries.
The program imports a CSV file that is given as an order request with a "Kit ID" and converts the "Kit ID" to an SKU-based order upload into a WMS.
This version is version 1.1 and has been updated to support the new "Sales order template" needed to upload into the WMS 4/30/2024
This program has error handling by only accepting predefined codes.
The EXE was created using pyinstaller
The install wizard was created using inno setup compiler

Future improvements include more OOP, adding additional clients other than USCG, creating a return ASN for specific Kit IDs, including a splash screen, including a progress bar for longer conversions, and Predefining a location to save the file so that uploads can be executed without saving a file twice. 
