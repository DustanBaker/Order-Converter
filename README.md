# Order-Converter
Tool for converting CSV files

This program was developed by Dusty Baker 2024
The program is written in Python using the Tkinter and CustomTkinter GUI libraries.
The program imports a CSV file that is given as an order request with a "Kit ID" and converts the "Kit ID" to an SKU-based order upload into a WMS.
This version is version 1.2 and has been updated to support the Terminix client with the various shipping levels. 
This program has error handling by only accepting predefined codes.
The EXE was created using pyinstaller
The install wizard was created using inno setup compiler

Future improvements include more OOP, adding additional clients other than USCG and Terminix, creating a return ASN for specific Kit IDs, including a splash screen, including a progress bar for longer conversions

V 1.5
Current features
USCG order converter- Imports order requests from the USCG and converts them to the SKU based entrys that the WMS can handle.
Terminix Order converter- Imports kit ID's for the terminix kits that changes the shipping service level
Eagle file manager - filters out all commas from the csv file and checks the char length
UPS batch file - N/A
