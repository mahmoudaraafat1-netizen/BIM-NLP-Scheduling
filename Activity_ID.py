import os
import subprocess
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# ✅ Function to check and install required libraries
def install_if_missing(library):
    try:
        __import__(library)
    except ImportError:
        print(f"⚠️ '{library}' not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", library])
        print(f"✅ '{library}' installed successfully!")

# ✅ Ensure required libraries are installed
install_if_missing("pandas")
install_if_missing("openpyxl")  # Needed for Excel operations

# ✅ Now we can import them safely
import pandas as pd

# ✅ Initialize Tkinter and hide root window
root = tk.Tk()
root.withdraw()  # Hide main window

# ✅ Open file selection dialog
file_path = filedialog.askopenfilename(title="Select Excel File (Activity List & Reference)", 
                                       filetypes=[("Excel Files", "*.xlsx *.xls")])
if not file_path:
    print("❌ No file selected, operation aborted.")
    exit()

# ✅ Load both sheets from Excel
df_activities = pd.read_excel(file_path, sheet_name="Activity List")
df_reference = pd.read_excel(file_path, sheet_name="Reference Dictionary")

# ✅ Ensure there are no NaN values in lookup tables
df_reference = df_reference.fillna("")  # Use empty string instead of GN

# ✅ Extract reference codes from sheet
# Convert dictionary to lookup tables
floor_codes = dict(zip(df_reference["Floor Name"].astype(str).str.lower(), df_reference["Floor Code"]))
phase_codes = dict(zip(df_reference["Phase Name"].astype(str).str.lower(), df_reference["Phase Code"]))
building_codes = dict(zip(df_reference["Building Name"].astype(str).str.lower(), df_reference["Building Code"]))

# ✅ Print available building names for debugging
print("\n📌 Available Buildings in Reference Dictionary:")
print(building_codes.keys())

# ✅ Function to extract floor code from activity name
def get_floor_code(activity_name):
    if pd.isna(activity_name):  
        return "GN"
    
    activity_name = str(activity_name).lower()
    
    for floor, code in floor_codes.items():
        if str(floor).lower() in activity_name:
            return code
    return "GN"  

# ✅ Function to extract phase code from activity name
def get_phase_code(activity_name):
    if pd.isna(activity_name):  
        return "GEN"
    
    activity_name = str(activity_name).lower()
    
    for phase, code in phase_codes.items():
        if str(phase).lower() in activity_name:
            return code
    return "GEN"  

# ✅ Function to extract building code from activity name (remove if not found)
def get_building_code(activity_name):
    if pd.isna(activity_name):  
        return ""  # Return empty string if no match
    
    activity_name = str(activity_name).lower()  

    for building, code in building_codes.items():
        if str(building).lower() in activity_name:
            print(f"✅ Found Building: {building} -> {code}")  # Debugging
            return code
    
    print(f"⚠️ No Building Found for: {activity_name}, skipping building code.")  # Debugging
    return ""  # Return empty if no match found

# ✅ Generate Activity ID for each row (without `Building Code` if not found)
activity_ids = []

for i, row in df_activities.iterrows():
    activity_name = row["Activity Name"]
    floor_code = get_floor_code(activity_name)
    phase_code = get_phase_code(activity_name)
    
    # Get building code (may be empty)
    building_code = get_building_code(activity_name)
    
    task_number = str(10 + (i * 5)).zfill(3)  # 🔥 Task Number starts from 10 and increases by 5
    
    # 🛠 **Create Activity ID without `Building Code` if it's empty**
    if building_code:
        activity_id = f"{building_code}-{floor_code}-{phase_code}-{task_number}"
    else:
        activity_id = f"{floor_code}-{phase_code}-{task_number}"
    
    activity_ids.append(activity_id)

# ✅ Insert the new column **before** "Activity Name"
df_activities.insert(0, "Activity ID", activity_ids)

# ✅ Open save file dialog
output_path = filedialog.asksaveasfilename(title="Select Save Location", defaultextension=".xlsx",
                                           filetypes=[("Excel Files", "*.xlsx *.xls")])

if not output_path:
    print("❌ No save location selected, operation aborted.")
    exit()

# ✅ Save the modified file
df_activities.to_excel(output_path, index=False)

# ✅ Open the file automatically after saving
os.system(f'start EXCEL.EXE \"{output_path}\"')  

print(f"\n✅ Activity IDs generated successfully! File saved at: {output_path}")
