import xml.etree.ElementTree as ET
from docx import Document
from tkinter import Tk, Label, Button, filedialog, messagebox

# XML to Word Project by Hasan

def select_xml_file():
    root = Tk()
    root.withdraw()
    xml_file = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if xml_file:
        generate_word(xml_file)
        messagebox.showinfo("Information", "Word-Datei wurde erstellt.")

def generate_word(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    doc = Document()
    process_element(root, doc)

    word_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    if word_file:
        doc.save(word_file)
        messagebox.showinfo("Information", f"Word-Datei wurde als {word_file} gespeichert.")

def process_element(element, doc, indent=0):
    doc.add_paragraph(f"{get_indent(indent)}{element.tag}: {element.text.strip()}")
    for child in element:
        process_element(child, doc, indent + 1)

def get_indent(indent):
    return "\t" * indent

# GUI erstellen
root = Tk()
root.title("XML to Word Converter by Hasan")

# Größe der GUI anpassen
root.geometry("400x200")

# Text und Button in der Mitte zentrieren
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)

label = Label(root, text="Wähle eine XML-Datei aus:")
label.pack(pady=20)

button = Button(root, text="Datei auswählen", command=select_xml_file)
button.pack(pady=10)

close_button = Button(root, text="Schließen", command=root.destroy)
close_button.pack(pady=10)

root.mainloop()
