import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


class TxtToExcelGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("TXT to Excel Converter")
        self.root.geometry("500x200")

        self.txt_path = ""
        self.save_path = ""

        # Select TXT file
        tk.Button(root, text="Select TXT File", command=self.select_txt).pack(pady=10)
        self.txt_label = tk.Label(root, text="No file selected")
        self.txt_label.pack()

        # Select Save Location
        tk.Button(root, text="Choose Save Location", command=self.select_save).pack(pady=10)
        self.save_label = tk.Label(root, text="No save location selected")
        self.save_label.pack()

        # Convert Button
        tk.Button(root, text="Convert to Excel", command=self.convert, bg="green", fg="white").pack(pady=20)

    def select_txt(self):
        path = filedialog.askopenfilename(
            title="Select text file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if path:
            self.txt_path = path
            self.txt_label.config(text=os.path.basename(path))

    def select_save(self):
        path = filedialog.asksaveasfilename(
            title="Save Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.save_path = path
            self.save_label.config(text=os.path.basename(path))

    def convert(self):
        if not self.txt_path:
            messagebox.showerror("Error", "Select a TXT file first.")
            return

        if not self.save_path:
            messagebox.showerror("Error", "Choose where to save the Excel file.")
            return

        try:
            with open(self.txt_path, "r", encoding="utf-8") as f:
                lines = [line.strip() for line in f if line.strip()]

            headers = re.split(r"\s+", lines[0])
            rows = []

            for line in lines[1:]:
                parts = re.split(r"\s+", line)

                if len(parts) < len(headers):
                    parts += [""] * (len(headers) - len(parts))
                elif len(parts) > len(headers):
                    parts = parts[:len(headers)]

                rows.append(parts)

            df = pd.DataFrame(rows, columns=headers)
            df.to_excel(self.save_path, index=False)

            messagebox.showinfo("Success", "File converted successfully.")

        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = TxtToExcelGUI(root)
    root.mainloop()