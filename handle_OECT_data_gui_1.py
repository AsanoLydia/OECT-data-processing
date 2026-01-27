# -*- coding: utf-8 -*-
"""
Created on Thu Aug  8 12:50:28 2024

@author: 14495
"""

import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, Entry, Label, Scrollbar, Listbox
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
        self.geometry('600x400')
        
        # Initialize variables
        self.filepaths = []
        self.starting_point = 449
        self.standard = -450
        self.startline = 15
        
        # Create widgets
        self.create_widgets()

    def create_widgets(self):
        # File selection
        tk.Button(self, text='Add Data Files', command=self.select_files).pack(pady=10)
        self.listbox = Listbox(self, width=50, height=10)
        self.listbox.pack(pady=10)
        Scrollbar(self.listbox, orient="vertical", command=self.listbox.yview).pack(side="right", fill="y")
        
        # Parameter settings
        tk.Button(self, text='Set Parameters', command=self.set_parameters).pack(pady=10)
        
        # Outputs
        tk.Label(self, text="Figure output filename:").pack()
        self.fig_output_entry = Entry(self)
        self.fig_output_entry.pack()
        self.fig_output_entry.insert(0, 'output.png')

        tk.Label(self, text="Excel output filename:").pack()
        self.excel_output_entry = Entry(self)
        self.excel_output_entry.pack()
        self.excel_output_entry.insert(0, 'output.xlsx')
        
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

    def set_parameters(self):
        self.starting_point = simpledialog.askinteger("Parameter", "Enter the starting point for normalization:", initialvalue=self.starting_point)
        self.standard = simpledialog.askfloat("Parameter", "Enter the standard value for normalization:", initialvalue=self.standard)
        self.startline = simpledialog.askinteger("Parameter", "Enter the line number to start reading from:", initialvalue=self.startline)

    def process_data(self):
        if not self.filepaths:
            messagebox.showerror("Error", "No files selected!")
            return
        
        all_data = pd.DataFrame()
        colors = plt.cm.rainbow(np.linspace(0, 1, len(self.filepaths)))

        for i, filepath in enumerate(self.filepaths):
            data = read_from_line(filepath, self.startline)
            normalized_dc = normalize_dc(data, self.starting_point, self.standard)
            
            plt.plot(data['time'], normalized_dc, label=clean_filename(filepath), color=colors[i])
            
            file_label = clean_filename(filepath)
            all_data[file_label] = normalized_dc
        
        all_data.index = data['time']
        plt.xlabel('Time (s)')
        plt.ylabel('Normalized drain current (uA)')
        plt.legend()
        plt.savefig(self.fig_output_entry.get())
        
        with pd.ExcelWriter(self.excel_output_entry.get(), engine='openpyxl') as writer:
            all_data.to_excel(writer, sheet_name='All Data', index_label='Time (s)')
        
        plt.show()  # Show the plot after saving to ensure the GUI doesn't freeze

if __name__ == "__main__":
    app = DataApp()
    app.mainloop()
