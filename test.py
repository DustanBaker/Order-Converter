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
import pygame

#Custom Tkinter top level window for splash screen

splash_root = customtkinter.CTk()
splash_root.geometry("800x450")
# locate the window in the center of the screen
x_cordinate = int((splash_root.winfo_screenwidth() / 2) - (800 / 2))
y_cordinate = int((splash_root.winfo_screenheight() / 2) - (450 / 2))
splash_root.geometry(f"+{x_cordinate}+{y_cordinate}")
splash_root.overrideredirect(True)
splash_root.attributes('-topmost', True)
#import and play the intro sound while splash screen is displayed



splash_image = customtkinter.CTkImage(light_image=Image.open('assets/splash.png'),
                                      dark_image=Image.open('assets/splash.png'),
                                      size=(800, 600))
splash_label = customtkinter.CTkLabel(splash_root, text="", image=splash_image)
splash_label.pack()


def main_window():
    # Close the splash screen
    splash_root.destroy()

    root = customtkinter.CTk()
    root.title("Eagle File Manager")
    root.geometry("600x650")
    root.iconbitmap('assets/Eagle.ico')

    # Create label
    Header_Label1 = customtkinter.CTkLabel(root, text="Version 1.6")
    Header_Label1.pack(pady=5)

    # create a tab view with custom tkinter
    My_tab = customtkinter.CTkTabview(root)
    My_tab.pack(expand=1, fill="both")

    # create a tab
    tab_1 = My_tab.add("The Eagle")
    tab_2 = My_tab.add("USCG")
    tab_3 = My_tab.add("Terminix")
    tab_4 = My_tab.add("UPS")


splash_root.mainloop()

