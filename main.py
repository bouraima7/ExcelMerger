# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


# define a Gui class
class ExcelMergerGUI:
    def __init__(self, root):
        self.root = root  # store the root window onject as an attribute of the class object
        self.root.title("Excel Merger")

        # create buttons for welcome message
        welcome_label = tk.Label(root, text="Welcome, this is a cross-referencing app!")
        welcome_label.pack(pady=10)

        # create buttons for file selection
        self.file1_button = tk.Button(root, text="Select Excel File 1", command=self.get_file1)
        self.file1_button.pack(pady=10)

        self.file2_button = tk.Button(root, text="Select Excel File 2", command=self.get_file2)
        self.file2_button.pack(pady=10)

        # create labels and labels and enrty for column name
        self.column_label = tk.Label(root, text="Enter Column Name to Sort By:")
        self.column_label.pack()
        self.column_entry = tk.Entry(root)
        self.column_entry.pack(pady=10)

        # create a ,erge and save button
        self.merge_button = tk.Button(root, text="Merge and Save as CSV", command=self.merge_and_save_csv)
        self.merge_button.pack(pady=20)

    # get the path of the first excel file
    def get_file1(self):
        self.file1_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    # get the path of the second secel file
    def get_file2(self):
        self.file2_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    def merge_and_save_csv(self):
        if not hasattr(self, 'file1_path') or not hasattr(self, 'file2_path'):  # determine whetehr the file paths exist within the selc object
            return

        column_name = self.column_entry.get().strip()

        if not column_name:
            messagebox.showwarning("Input Error", "Please enter a column name.")
            return

        try:
            # Read Excel files into Pandas DataFrames
            df1 = pd.read_excel(self.file1_path)
            df2 = pd.read_excel(self.file2_path)

            # Merge tables and order by specified column and Severity
            merged_df = pd.concat([df1, df2], ignore_index=True)
            merged_df = merged_df.sort_values(by=[column_name, 'Severity'], ascending=[True, True])

            # Save merged result as CSV
            self.save_to_csv(merged_df)

        # carch any exceptions and assign it to e
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")  # display the rror to the user on screen

    def save_to_csv(self, dataframe):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])

        if file_path:
            dataframe.to_csv(file_path, index=False)
            messagebox.showinfo("Save Successful", f"The merged result has been saved as CSV to {file_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerGUI(root)
    root.mainloop()