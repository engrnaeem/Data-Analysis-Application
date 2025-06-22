import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class DataAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer and Analyzer")
        
        self.file_path = None  # Path to the loaded Excel file
        self.data = None  # DataFrame to store the loaded data
        self.column_vars = {}  # Dictionary to store column names and their associated BooleanVar for checkboxes

        self.create_widgets()  # Initialize the GUI components

    def create_widgets(self):
        # Application heading
        self.titl = tk.Label(self.root, text="Data Analysis Application", font=("Arial", 16))
        self.titl.pack(pady=10)

        # Load file button
        self.loading_btn = tk.Button(self.root, text="Choose Excel File", command=self.load_file)
        self.loading_btn.pack(pady=10)

        # Frame for horizontal buttons
        self.btn_frm = tk.Frame(self.root)
        self.btn_frm.pack(pady=10)
        
        # View Head & Tail button
        self.dsply_btn = tk.Button(self.btn_frm, text="View Head & Tail", command=self.view_head_tail)
        self.dsply_btn.grid(row=0, column=0, padx=5)
        
        # Analyze Data button
        self.analyze_btn = tk.Button(self.btn_frm, text="Analyze Data", command=self.analyze_data)
        self.analyze_btn.grid(row=0, column=1, padx=5)
        
        # Clean Data button
        self.clear_btn = tk.Button(self.btn_frm, text="Clean Data", command=self.clean_data)
        self.clear_btn.grid(row=0, column=2, padx=5)
        
        # Plot Data button
        self.plt_btn = tk.Button(self.btn_frm, text="Plot Data", command=self.plot_data)
        self.plt_btn.grid(row=0, column=3, padx=5)
        
        # Frame for checkboxes to select columns for plotting
        self.chkbx_frm = tk.Frame(self.root)
        self.chkbx_frm.pack(pady=10)
        
        # Text box for displaying data and analysis results
        self.txt_bx = tk.Text(self.root, wrap='none', height=20, width=80)
        self.txt_bx.pack(pady=10)

    def load_file(self):
        # Open a file dialog to select an Excel file
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file_path:
            try:
                # Read the selected Excel file into a DataFrame
                self.data = pd.read_excel(self.file_path)
                # Display checkboxes for each column
                self.display_column_checkboxes()
                messagebox.showinfo("Success", "File loaded successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")
    
    def display_column_checkboxes(self):
        # Clear any existing checkboxes
        for widget in self.chkbx_frm.winfo_children():
            widget.destroy()
        
        self.column_vars = {}
        # Create a checkbox for each column in the DataFrame
        for column in self.data.columns:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(self.chkbx_frm, text=column, variable=var, command=self.update_selected_columns)
            chk.pack(side=tk.LEFT)
            self.column_vars[column] = var

    def update_selected_columns(self):
        # Update the list of selected columns based on checkbox states
        self.selected_columns = [col for col, var in self.column_vars.items() if var.get()]

    def view_head_tail(self):
        # Display the first and last five rows of the DataFrame
        if self.data is not None:
            self.txt_bx.delete(1.0, tk.END)
            head_tail = self.data.head().to_string() + "\n\n" + self.data.tail().to_string()
            self.txt_bx.insert(tk.END, head_tail)
        else:
            messagebox.showwarning("Warning", "No data loaded")
    
    def analyze_data(self):
        # Display descriptive statistics of the DataFrame
        if self.data is not None:
            analysis = self.data.describe().to_string()
            self.txt_bx.delete(1.0, tk.END)
            self.txt_bx.insert(tk.END, analysis)
        else:
            messagebox.showwarning("Warning", "No data loaded")

    def clean_data(self):
        # Remove rows with missing values (NaNs) from the DataFrame
        if self.data is not None:
            self.data.dropna(inplace=True)
            messagebox.showinfo("Success", "Data cleaned (NaN values dropped)")
        else:
            messagebox.showwarning("Warning", "No data loaded")
    
    def plot_data(self):
        # Plot the selected columns of the DataFrame
        if self.data is not None and self.selected_columns:
            selected_columns = self.selected_columns
            if selected_columns:
                self.plot_options(selected_columns)
            else:
                messagebox.showwarning("Warning", "No columns selected")
        else:
            messagebox.showwarning("Warning", "No data loaded")

    def plot_options(self, columns):
        # Create a new window for selecting plot type
        plot_window = tk.Toplevel(self.root)
        plot_window.title("Select Plot Type")

        # Define a function to handle plot type selection
        def plot(plot_type):
            plot_window.destroy()
            self.create_plot(columns, plot_type)
        
        # Create buttons for each plot type
        tk.Button(plot_window, text="Line Plot", command=lambda: plot('line')).pack(pady=5)
        tk.Button(plot_window, text="Scatter Plot", command=lambda: plot('scatter')).pack(pady=5)
        tk.Button(plot_window, text="Bar Chart", command=lambda: plot('bar')).pack(pady=5)
        tk.Button(plot_window, text="Histogram", command=lambda: plot('hist')).pack(pady=5)

    def create_plot(self, columns, plot_type):
        # Create the selected type of plot for the chosen columns
        plt.figure(figsize=(10, 6))
        if plot_type == 'line':
            for col in columns:
                plt.plot(self.data[col], label=col)
            plt.legend()
        elif plot_type == 'scatter':
            if len(columns) >= 2:
                plt.scatter(self.data[columns[0]], self.data[columns[1]])
                plt.xlabel(columns[0])
                plt.ylabel(columns[1])
            else:
                messagebox.showwarning("Warning", "Scatter plot requires at least two columns")
                return
        elif plot_type == 'bar':
            self.data[columns].plot(kind='bar', ax=plt.gca())
        elif plot_type == 'hist':
            self.data[columns].plot(kind='hist', ax=plt.gca(), alpha=0.5)

        plt.title(f'{plot_type.capitalize()} Plot')
        plt.grid(True)

        # Embed the plot in the Tkinter window
        canvas = FigureCanvasTkAgg(plt.gcf(), master=self.root)
        canvas.get_tk_widget().pack()
        canvas.draw()

if __name__ == "__main__":
    # Create the main Tkinter window and run the application
    root = tk.Tk()
    app = DataAnalysisApp(root)
    root.mainloop()
