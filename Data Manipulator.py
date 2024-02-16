import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from fuzzywuzzy import fuzz
import os
import xlsxwriter

def exact_match(file_path, column_name):
    try:
        # Read the file
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xls') or file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file format. Please provide a CSV or Excel file.")

        # Check if the column name exists in the DataFrame
        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in the file.")

        # Find duplicate values
        duplicates = df[df.duplicated(subset=[column_name], keep=False)]

        if duplicates.empty:
            messagebox.showinfo("Duplicate Finder", "No duplicate values found.")
        else:
            # Save the duplicates to a new file
            output_file_path = os.path.splitext(file_path)[0] + "_duplicates.xlsx"
            duplicates.to_excel(output_file_path, index=False)
            messagebox.showinfo("Duplicate Finder", f"Duplicate values saved at {output_file_path}\n\nDuplicate values found in rows: {duplicates.index.tolist()}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def threshold_match(file_path, column_name, threshold=90):
    try:
        # Read the file
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xls') or file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file format. Please provide a CSV or Excel file.")

        # Check if the column name exists in the DataFrame
        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in the file.")

        # Get the values from the specified column
        values = df[column_name].astype(str).str.strip()

        # Add a new column to indicate duplicates
        df['IsDuplicate'] = False

        # Iterate through each value
        for i, value in enumerate(values):
            for j in range(i + 1, len(values)):
                if fuzz.partial_ratio(value, values[j]) >= threshold:
                    # Mark both rows as duplicates
                    df.at[i, 'IsDuplicate'] = True
                    df.at[j, 'IsDuplicate'] = True

        # Save the modified file
        output_file_path = os.path.splitext(file_path)[0] + "_marked.xlsx"
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)

        messagebox.showinfo("Success", f"Duplicate values marked. Output file saved at {output_file_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def increase_percentage(file_path, percentage):
    try:
        # Read the file
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xls') or file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file format. Please provide a CSV or Excel file.")

        # Increase the percentage for all numeric columns
        df = df.apply(lambda x: x * (1 + percentage / 100) if pd.api.types.is_numeric_dtype(x) else x)

        # Save the modified file
        output_file_path = os.path.splitext(file_path)[0] + "_increased.xlsx"
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)

        messagebox.showinfo("Success", f"Percentage increased for all numeric columns. Output file saved at {output_file_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def browse_file(operation):
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls *.xlsx"), ("CSV Files", "*.csv")])
    if file_path:
        if operation == 'duplicate':
            column_name = column_name_entry.get()
            threshold = int(threshold_entry.get())
            if method_var.get() == "Exact Match":
                exact_match(file_path, column_name)
            elif method_var.get() == "Threshold Match":
                threshold_match(file_path, column_name, threshold)
        elif operation == 'increase':
            percentage = float(percentage_entry.get())
            increase_percentage(file_path, percentage)

# Create the main window
root = tk.Tk()
root.title("Data Manipulation")

# Create a custom style for the widgets
style = ttk.Style()
style.configure('Color.TButton', foreground='blue', background='orange', font=('Arial', 10, 'bold'))
style.configure('Icon.TButton', foreground='green', background='yellow', font=('Arial', 10, 'bold italic'))


# Frame for Duplicate Finder
duplicate_frame = ttk.LabelFrame(root, text="Duplicate Finder")
duplicate_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

tk.Label(duplicate_frame, text="Column Name:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
column_name_entry = tk.Entry(duplicate_frame)
column_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

tk.Label(duplicate_frame, text="Threshold:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
threshold_entry = tk.Entry(duplicate_frame)
threshold_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

method_var = tk.StringVar()
method_var.set("Exact Match")
method_label = tk.Label(duplicate_frame, text="Choose Method:")
method_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
method_combobox = ttk.Combobox(duplicate_frame, textvariable=method_var, values=["Exact Match", "Threshold Match"])
method_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="w")

duplicate_button = ttk.Button(duplicate_frame, text="Search Duplicates", command=lambda: browse_file('duplicate'), style='Color.TButton')
duplicate_button.grid(row=3, column=0, columnspan=2, pady=10)

# Frame for Percentage Increase
increase_frame = ttk.LabelFrame(root, text="Percentage Increase")
increase_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

tk.Label(increase_frame, text="Percentage:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
percentage_entry = tk.Entry(increase_frame)
percentage_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

increase_button = ttk.Button(increase_frame, text="Increase Percentage", command=lambda: browse_file('increase'), style='Icon.TButton')
increase_button.grid(row=1, column=0, columnspan=2, pady=10)

design_label = tk.Label(root, text="Design by Aryan Sharma", font=('Arial', 8))
design_label.grid(row=2, column=0, padx=10, pady=5, sticky="e")

version_label = tk.Label(root, text="Version 1.0", font=('Arial', 8))
version_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")

root.mainloop()
