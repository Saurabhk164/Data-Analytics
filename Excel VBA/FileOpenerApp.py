import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def open_file_from_excel(filename_to_search, excel_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(excel_path)

    # Check if required columns are present
    if 'Filename' not in df.columns or 'Filepath' not in df.columns:
        raise ValueError("Excel file must contain 'Filename' and 'Filepath' columns")

    # Search for the filename in the 'Filename' column
    row = df[df['Filename'] == filename_to_search]

    # Check if the filename was found
    if row.empty:
        messagebox.showerror("Error", f"Filename '{filename_to_search}' not found in the Excel file.")
        return

    # Retrieve the filepath from the corresponding row
    filepath = row.iloc[0]['Filepath']

    # Check if the file exists at the given filepath
    if not os.path.exists(filepath):
        messagebox.showerror("Error", f"Filepath '{filepath}' does not exist.")
        return

    # Open and read the file (for example, if it's a text file)
    with open(filepath, 'r') as file:
        file_content = file.read()

    # Display the file content
    text_box.delete("1.0", tk.END)
    text_box.insert(tk.END, file_content)

def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    excel_path_entry.delete(0, tk.END)
    excel_path_entry.insert(0, file_path)

def search_and_open_file():
    filename_to_search = filename_entry.get()
    excel_path = excel_path_entry.get()
    try:
        open_file_from_excel(filename_to_search, excel_path)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Create the main window
root = tk.Tk()
root.title("File Opener App")

# Create and place the widgets
tk.Label(root, text="Filename:").grid(row=0, column=0, padx=10, pady=10)
filename_entry = tk.Entry(root, width=50)
filename_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Excel File:").grid(row=1, column=0, padx=10, pady=10)
excel_path_entry = tk.Entry(root, width=50)
excel_path_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_excel_file).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Open File", command=search_and_open_file).grid(row=2, column=0, columnspan=3, pady=10)

text_box = tk.Text(root, wrap='word', width=60, height=20)
text_box.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

# Run the application
root.mainloop()

