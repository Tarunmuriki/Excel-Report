import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from report_generator import ReportGenerator

class ExcelReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Report Generator")
        self.root.geometry("600x400")
        
        self.input_file = ""
        self.output_file = ""
        self.setup_ui()
        
    def setup_ui(self):
        # Input file section
        input_frame = ttk.LabelFrame(self.root, text="Input", padding="10")
        input_frame.pack(fill="x", padx=10, pady=5)
        
        self.input_label = ttk.Label(input_frame, text="No file selected")
        self.input_label.pack(side="left", padx=5)
        
        ttk.Button(input_frame, text="Browse", command=self.select_input).pack(side="right")
        
        # Report options section
        options_frame = ttk.LabelFrame(self.root, text="Report Options", padding="10")
        options_frame.pack(fill="x", padx=10, pady=5)
        
        # Add checkboxes for report options
        self.include_pivot = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Include Pivot Tables", 
                       variable=self.include_pivot).pack(anchor="w")
        
        self.include_charts = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Include Charts", 
                       variable=self.include_charts).pack(anchor="w")
        
        self.include_stats = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Include Summary Statistics", 
                       variable=self.include_stats).pack(anchor="w")
        
        # Generate button
        ttk.Button(self.root, text="Generate Report", 
                  command=self.generate_report).pack(pady=20)
        
    def select_input(self):
        filename = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv")]
        )
        if filename:
            self.input_file = filename
            self.input_label.config(text=filename.split("/")[-1])
            
    def generate_report(self):
        if not self.input_file:
            messagebox.showerror("Error", "Please select an input file first!")
            return
        
        # Verify input file is CSV
        if not self.input_file.lower().endswith('.csv'):
            messagebox.showerror("Error", "Input file must be a CSV file!")
            return
            
        output_file = filedialog.asksaveasfilename(
            title="Save Excel Report",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if output_file:
            try:
                # Ensure output file has .xlsx extension
                if not output_file.lower().endswith('.xlsx'):
                    output_file += '.xlsx'
                
                generator = ReportGenerator(self.input_file)
                generator.generate_report(
                    output_file,
                    include_pivot=self.include_pivot.get(),
                    include_charts=self.include_charts.get(),
                    include_stats=self.include_stats.get()
                )
                messagebox.showinfo("Success", "Report generated successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate report: {str(e)}")
                print(f"Error details: {str(e)}")

def main():
    root = tk.Tk()
    app = ExcelReportApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()