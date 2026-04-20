import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def txt_to_xlsx():
    root = tk.Tk()
    root.withdraw()

    txt_path = filedialog.askopenfilename(
        title="Select text file",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )

    if not txt_path:
        return

    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip()]

        if not lines:
            messagebox.showerror("Error", "The selected file is empty.")
            return

        # First non-empty line is treated as header
        headers = re.split(r"\s+", lines[0])

        # Remaining lines are data rows
        rows = []
        for line in lines[1:]:
            parts = re.split(r"\s+", line)

            # If row is shorter than header, pad it
            if len(parts) < len(headers):
                parts += [""] * (len(headers) - len(parts))

            # If row is longer than header, trim it
            elif len(parts) > len(headers):
                parts = parts[:len(headers)]

            rows.append(parts)

        df = pd.DataFrame(rows, columns=headers)

        output_path = os.path.splitext(txt_path)[0] + ".xlsx"
        df.to_excel(output_path, index=False)

        messagebox.showinfo("Success", f"Excel file created:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert file.\n\n{e}")


if __name__ == "__main__":
    txt_to_xlsx()