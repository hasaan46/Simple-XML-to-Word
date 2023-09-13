import json
from docx import Document
from tkinter import Tk, Label, Button, filedialog, messagebox
from tkinter import CENTER


# XML Project by Hasan

def select_json_file():
    root = Tk()
    root.withdraw()
    json_file = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if json_file:
        generate_word(json_file)
        messagebox.showinfo("Information", "Word-Datei wurde erstellt.")

def generate_word(json_file, indent=0):
    with open(json_file) as file:
        json_data = json.load(file)

        doc = Document()
        for key, value in json_data.items():
            doc.add_paragraph(f"{get_indent(indent)}{key}: {format_value(value, indent)}")

        word_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if word_file:
            doc.save(word_file)
            messagebox.showinfo("Information", f"Word-Datei wurde als {word_file} gespeichert.")

def format_value(value, indent):
    if isinstance(value, dict):
        result = ""
        for k, v in value.items():
            result += f"\n{get_indent(indent+1)}{k}: {format_value(v, indent+1)}"
        return result
    elif isinstance(value, list):
        result = ""
        for item in value:
            result += f"\n{get_indent(indent+1)}- {format_value(item, indent+1)}"
        return result
    else:
        return str(value)

def get_indent(indent):
    return "\t" * indent

# GUI erstellen
root = Tk()
root.title("JSON Converter by Hasan")

# Größe der GUI anpassen
root.geometry("400x200")

# Text und Button in der Mitte zentrieren
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)

label = Label(root, text="Wähle eine JSON-Datei aus:")
label.pack(pady=20)

button = Button(root, text="Datei auswählen", command=select_json_file)
button.pack(pady=10)

close_button = Button(root, text="Schließen", command=root.destroy)
close_button.pack(pady=10)

root.mainloop()