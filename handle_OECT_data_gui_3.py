# -*- coding: utf-8 -*-
"""
Created on Sat Aug 31 17:53:30 2024

@author: 14495
"""
import tkinter as tk
from tkinter import Listbox, Scrollbar, filedialog, messagebox, simpledialog, Entry, IntVar
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import os

def read_from_line(filename, start_line):
    data = {'time': [], 'gate voltage': [], 'drain current': [], 'gate current': []}
    with open(filename, 'r') as file:
        for current_line_number, line in enumerate(file, start=1):
            if current_line_number >= start_line:
                parts = line.strip().split()
                parts[0] = float(parts[0]) / 1000
                parts[2] = float(parts[2]) * 1000000
                data['time'].append(parts[0])
                data['gate voltage'].append(float(parts[1]))
                data['drain current'].append(float(parts[2]))
                data['gate current'].append(float(parts[3]))
    return data

def normalize_dc(data, starting_point, standard):
    normalized_drain_current = []
    distance = data['drain current'][starting_point] - standard
    for current in data['drain current']:
        normalized_current = current - distance
        normalized_drain_current.append(normalized_current)
    return normalized_drain_current

def clean_filename(filename):
    return os.path.basename(filename).replace(':', '_').replace('\\', '_').replace('/', '_').replace('?', '_').replace('*', '_').replace('[', '_').replace(']', '_')

class DataApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Data Processing Application')
        self.geometry('600x600')
        
        # Initialize variables
        self.filepaths = []
        self.starting_point = 449
        self.standard = -450
        self.startline = 15
        self.save_figure = IntVar(value=1)  # Variable to track the checkbox state (1 = checked, 0 = unchecked)
        
        # Create widgets
        self.create_widgets()

    def create_widgets(self):
        # File selection
        tk.Button(self, text='Add Data Files', command=self.select_files).pack(pady=10)
        
        # Frame for Listbox and Scrollbar
        listbox_frame = tk.Frame(self)
        listbox_frame.pack(pady=10)
        
        self.listbox = Listbox(listbox_frame, width=50, height=10)
        self.listbox.pack(side="left", fill="y")
        
        scrollbar = Scrollbar(listbox_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        
        # Configure the scrollbar
        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)

        # Button to clear files from Listbox
        tk.Button(self, text='Clear Files', command=self.clear_files).pack(pady=10)
        
        # Parameter settings
        tk.Button(self, text='Set Parameters', command=self.set_parameters).pack(pady=10)

        # Checkbox to select whether to save the figure
        tk.Checkbutton(self, text="Save figure", variable=self.save_figure).pack(pady=10)
        
        # Outputs
        tk.Label(self, text="Figure output filename:").pack()
        self.fig_output_entry = Entry(self)
        self.fig_output_entry.pack()
        self.fig_output_entry.insert(0, 'output.png')

        tk.Label(self, text="CSV output filename:").pack()
        self.csv_output_entry = Entry(self)
        self.csv_output_entry.pack()
        self.csv_output_entry.insert(0, 'output.csv')
        
        # Process button
        tk.Button(self, text='Process Data', command=self.process_data).pack(pady=10)

    def select_files(self):
        new_filepaths = filedialog.askopenfilenames(title='Select data files')
        if new_filepaths:
            self.filepaths.extend(new_filepaths)  # Append new files to the list
            self.listbox.delete(0, tk.END)  # Clear the listbox
            for filepath in self.filepaths:  # Re-populate the listbox
                self.listbox.insert(tk.END, filepath)
            messagebox.showinfo("Files Selected", f"Added {len(new_filepaths)} files. Total: {len(self.filepaths)} files.")

    def clear_files(self):
        """Clear the listbox and the filepaths list."""
        self.listbox.delete(0, tk.END)  # Clear all items in the listbox
        self.filepaths.clear()  # Clear the filepaths list
        messagebox.showinfo("Cleared", "All files have been cleared from the list.")

    def set_parameters(self):
        # Create a new top-level window for parameter setting
        param_window = tk.Toplevel(self)
        param_window.title("Set Parameters")
        param_window.geometry("300x300")
        param_window.transient(self)
        param_window.grab_set()
        
        # Starting point entry
        tk.Label(param_window, text="Starting Point:").pack(pady=5)
        starting_point_entry = tk.Entry(param_window)
        starting_point_entry.pack(pady=5)
        starting_point_entry.insert(0, str(self.starting_point))

        # Standard value entry
        tk.Label(param_window, text="Standard Value:").pack(pady=5)
        standard_entry = tk.Entry(param_window)
        standard_entry.pack(pady=5)
        standard_entry.insert(0, str(self.standard))

        # Start line entry
        tk.Label(param_window, text="Start Line:").pack(pady=5)
        startline_entry = tk.Entry(param_window)
        startline_entry.pack(pady=5)
        startline_entry.insert(0, str(self.startline))

        # Confirm button
        def on_confirm():
            try:
                self.starting_point = int(starting_point_entry.get())
                self.standard = float(standard_entry.get())
                self.startline = int(startline_entry.get())
                param_window.destroy()
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter valid values for all parameters.")

        tk.Button(param_window, text="Confirm", command=on_confirm).pack(pady=10)

    def process_data(self):
        try:
            if not self.filepaths:
                messagebox.showerror("Error", "No files selected!")
                return
        
            all_data = None  # Use None as the initial state for merging
            colors = plt.cm.rainbow(np.linspace(0, 1, len(self.filepaths)))

            for i, filepath in enumerate(self.filepaths):
                data = read_from_line(filepath, self.startline)
                normalized_dc = normalize_dc(data, self.starting_point, self.standard)
            
                temp_df = pd.DataFrame({
                    'Time': data['time'],
                    clean_filename(filepath): normalized_dc
                    })
            
                if all_data is None:
                    all_data = temp_df
                else:
                    all_data = pd.merge(all_data, temp_df, on='Time', how='outer')

                if self.save_figure.get() == 1:  # Check if the checkbox is selected
                    plt.plot(data['time'], normalized_dc, label=clean_filename(filepath), color=colors[i])
        
            all_data = all_data.sort_values(by='Time').reset_index(drop=True)
        
            if self.save_figure.get() == 1:  # Save the figure only if checkbox is selected
                plt.xlabel('Time (s)')
                plt.ylabel('Normalized drain current (uA)')
                plt.legend()
                plt.savefig(self.fig_output_entry.get())
                plt.show()  # Show the plot after saving to ensure the GUI doesn't freeze
        
            # Save the data to CSV file
            all_data.to_csv(self.csv_output_entry.get(), index=False)
        
            # Show a completion message box
            messagebox.showinfo("Process completed", "Data processing and CSV export completed successfully!")
        
        except Exception as e:
            # show error message
            messagebox.showerror("Error", f"An error occurred during processing: {str(e)}")
    
    
if __name__ == "__main__":
    app = DataApp()
    app.mainloop()
