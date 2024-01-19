import tkinter as tk
from tkinter import filedialog
import json
import ezodf

def odf_to_json(odf_path, json_path):
    doc = ezodf.opendoc(odf_path)
    sheet = doc.sheets[0]  # Assuming the first sheet is the targeted one

    data = {}

    for row_idx, row in enumerate(sheet.rows()):
        if row_idx == 0:
            headers = [cell.value for cell in row]
        else:
            # Data which is stored in a dictionary with the first column acting as the key
            item_id = row[0].value
            item_data = dict(zip(headers[1:], [cell.value for cell in row[1:]]))
            data[item_id] = item_data

    with open(json_path, 'w') as json_file:
        json.dump(data, json_file, indent=4)

def browse_file(entry_widget):
    file_path = filedialog.askopenfilename(title="Select File", filetypes=[("ODS Files", "*.ods")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def browse_json():
    file_path = filedialog.asksaveasfilename(title="Save JSON File", defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    entry_json.delete(0, tk.END)
    entry_json.insert(0, file_path)

def convert_to_json():
    file_path = entry_odf.get()
    if file_path.endswith(".ods"):
        json_path = entry_json.get()
        odf_to_json(file_path, json_path)
        status_label.config(text="Conversion complete!")
    else:
        status_label.config(text="Unsupported file format!")

# GUI
root = tk.Tk()
root.title("ODS to JSON - WaffleAWT")

# ODS file
label_odf = tk.Label(root, text="ODS File:")
label_odf.grid(row=0, column=0, padx=10, pady=5, sticky=tk.E)
entry_odf = tk.Entry(root, width=40)
entry_odf.grid(row=0, column=1, padx=5, pady=5)
btn_browse_odf = tk.Button(root, text="Browse", command=lambda: browse_file(entry_odf))
btn_browse_odf.grid(row=0, column=2, padx=5, pady=5)  # Updated command here

# JSON file
label_json = tk.Label(root, text="JSON File:")
label_json.grid(row=1, column=0, padx=10, pady=5, sticky=tk.E)
entry_json = tk.Entry(root, width=40)
entry_json.grid(row=1, column=1, padx=5, pady=5)
btn_browse_json = tk.Button(root, text="Browse", command=browse_json)
btn_browse_json.grid(row=1, column=2, padx=5, pady=5)

# Button
btn_convert = tk.Button(root, text="Convert to JSON", command=convert_to_json)
btn_convert.grid(row=2, column=1, pady=10)

# Status label
status_label = tk.Label(root, text="")
status_label.grid(row=3, column=1)

root.mainloop()

# Follow me on @github: github.com/waffleawt :)