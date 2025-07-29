import tkinter as tk
from tkinter import filedialog, messagebox
from merge_subnets import main  # Use backend logic

class SubnetMergerTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Subnet Merger Tool")

        # Labels
        tk.Label(root, text="Input File:").grid(row=0, column=0, padx=10, pady=10)
        tk.Label(root, text="Output File:").grid(row=1, column=0, padx=10, pady=10)

        # Entry fields
        self.input_file = tk.Entry(root, width=50)
        self.input_file.grid(row=0, column=1, padx=10, pady=10)

        self.output_file = tk.Entry(root, width=50)
        self.output_file.grid(row=1, column=1, padx=10, pady=10)

        # Browse buttons
        tk.Button(root, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=10, pady=10)
        tk.Button(root, text="Save As", command=self.browse_output).grid(row=1, column=2, padx=10, pady=10)

        # Merge Button
        tk.Button(root, text="Merge Subnets", command=self.merge_subnets,
                  bg="grey", fg="white", width=20).grid(row=2, column=1, pady=20)

    def browse_input(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel or PDF Files", "*.xlsx *.xls *.pdf")])
        if file_path:
            self.input_file.delete(0, tk.END)
            self.input_file.insert(0, file_path)

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.output_file.delete(0, tk.END)
            self.output_file.insert(0, file_path)

    def merge_subnets(self):
        input_path = self.input_file.get()
        output_path = self.output_file.get()

        if input_path and output_path:
            try:
                main(input_path, output_file=output_path)
                messagebox.showinfo("Success", f"Subnets merged successfully!\nSaved to:\n{output_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred during merging:\n{e}")
        else:
            messagebox.showerror("Missing Fields", "Please provide both an input file and an output file path.")

if __name__ == "__main__":
    root = tk.Tk()
    app = SubnetMergerTool(root)
    root.mainloop()
