import pandas as pd
import os
from tkinter import Tk, Button, messagebox, filedialog

# Function to generate .resx content from DataFrame
def generate_resx(df):
    resx_template = '''<?xml version="1.0" encoding="utf-8"?>
<root>
  <resheader name="resmimetype">
    <value>text/microsoft-resx</value>
  </resheader>
  <resheader name="version">
    <value>2.0</value>
  </resheader>
  <resheader name="reader">
    <value>System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>
  </resheader>
  <resheader name="writer">
    <value>System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>
  </resheader>
'''

    # Iterate over rows in DataFrame to maintain the order
    for index, row in df.iterrows():
        key = str(row[0]).strip()  # Use the 1st selected column for key
        value = str(row[1]).strip()  # Use the 2nd selected column for value
        resx_template += f'  <data name="{key}" xml:space="preserve">\n'
        resx_template += f'    <value>{value}</value>\n'
        resx_template += '  </data>\n'

    resx_template += '</root>'
    return resx_template

# Function to read Excel and generate .resx files for each sheet
def read_excel_and_generate_resx():
    # Ask the user to select an Excel file
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not excel_file:
        return

    # Ask the user to select an output directory
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    if not output_directory:
        return

    try:
        # Read the Excel file
        xls = pd.ExcelFile(excel_file)

        for sheet_name in xls.sheet_names:
            # Read each sheet into a DataFrame using the correct columns for key and value
            df = pd.read_excel(xls, sheet_name=sheet_name, usecols=[1, 2], engine='openpyxl')

            # Generate .resx content
            resx_content = generate_resx(df)

            # Define the .resx file path
            resx_file = os.path.join(output_directory, f'{sheet_name}.resx')

            # Write .resx content to file
            with open(resx_file, 'w', encoding='utf-8') as file:
                file.write(resx_content)

        messagebox.showinfo("Success", f".resx files generated in: {output_directory}")
    except FileNotFoundError:
        messagebox.showerror("Error", f"Excel file not found: {excel_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main window
root = Tk()
root.title("Excel to .resx Converter")

# Create the button and assign the function
button = Button(root, text="Convert Excel to .resx", command=read_excel_and_generate_resx)
button.pack(pady=20)

# Run the main loop
root.mainloop()
