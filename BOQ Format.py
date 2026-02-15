import pandas as pd
import re
from tkinter import Tk, filedialog

# Function to select file using GUI
def select_file():
    root = Tk()
    root.withdraw()  # Hide main window
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
    return file_path

# Function to save file using GUI
def save_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    return file_path

# Regular expression to extract element name
regex_pattern = re.compile(r"Name=(.*?),")

# Select input file
file_path = select_file()
if not file_path:
    print("No file selected. Process canceled.")
    exit()

# Load the workbook
xls = pd.ExcelFile(file_path)
output_data = []

# Process each sheet
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Check if the sheet is empty
    if df.empty or df.shape[1] == 0:
        print(f"⚠️ Warning: Sheet '{sheet_name}' is empty or has no columns. Skipping...")
        continue
    
    # Print available columns for debugging
    print(f"📜 Sheet '{sheet_name}' columns: {list(df.columns)}")
    
    # Identify columns dynamically
    col_map = {col.lower(): col for col in df.columns}  # Create a dictionary for columns
    col_element = next((col for col in df.columns if "element" in col.lower()), None)
    col_area = col_map.get("area", None)
    col_volume = col_map.get("volume", None)
    
    # If 'Element Name' column is not found, ask the user for input
    if not col_element:
        print(f"⚠️ Warning: Sheet '{sheet_name}' does not have an 'Element Name' column.")
        print(f"📜 Available columns: {list(df.columns)}")
        col_element = input("👉 Please enter the correct column name for Element Name (or press Enter to skip): ").strip()
        if col_element not in df.columns:
            print(f"🚫 Skipping sheet '{sheet_name}' due to missing column.")
            continue
    
    # Fill missing area/volume with 0 if necessary
    if col_area is None:
        df["area"] = 0
        col_area = "area"
    if col_volume is None:
        df["volume"] = 0
        col_volume = "volume"
    
    # Extract element names
    df["element_name"] = df[col_element].astype(str).str.extract(regex_pattern)
    df.dropna(subset=["element_name"], inplace=True)  # Remove rows with no valid name
    
    # Aggregate data
    aggregated = df.groupby("element_name").agg(
        total_area=(col_area, "sum"),
        total_volume=(col_volume, "sum"),
        count=("element_name", "count")
    ).reset_index()
    
    # Add sheet name column
    aggregated.insert(0, "Sheet Name", sheet_name)
    output_data.append(aggregated)

# Combine all results
if output_data:
    final_df = pd.concat(output_data, ignore_index=True)
    
    # Save output file
    save_path = save_file() or "Aggregated_Data.xlsx"  # Default save file if user cancels
    final_df.to_excel(save_path, index=False)
    print(f"Processing complete. File saved to: {save_path}")
else:
    print("No valid data found to process.")
