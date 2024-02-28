import pandas as pd
from format_data import FormatData
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog

class FileGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("File Generator")
        self.file_name = tk.StringVar()
        self.xml_path = tk.StringVar()
        self.create_widgets()

    def select_excel_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        self.file_name.set(filename)
        self.f = str(filename)
        return self.f

    def select_xml_file(self):
        filename = filedialog.asksaveasfilename(filetypes=[("XML files", "*.xml")])
        self.xml_path.set(filename)
        self.xml_filename = str(filename)
        return self.xml_filename

    def generate_file(self):
        format_data = FormatData(self.f)
        data = format_data.make_full_data()
        if '.xml' in self.xml_filename:
            filepath = self.xml_filename
        else:
            filepath = f'{self.xml_filename}.xml'    
        with open(filepath, 'w', encoding='utf-8') as text:
            text.write(data)
        messagebox.showinfo("Успех", "Конфигурация успешно сформирована")    

    def create_widgets(self):
        excel_label = tk.Label(self.root, text="Выбрать Excel файл:")
        excel_label.pack()

        excel_entry = tk.Entry(self.root, textvariable=self.file_name)
        excel_entry.pack()

        excel_button = tk.Button(self.root, text="Выбрать Excel файл", command=self.select_excel_file)
        excel_button.pack()

        xml_label = tk.Label(self.root, text="Путь сохранения XML:")
        xml_label.pack()

        xml_entry = tk.Entry(self.root, textvariable=self.xml_path)
        xml_entry.pack()

        xml_button = tk.Button(self.root, text="Путь сохранения XML", command=self.select_xml_file)
        xml_button.pack()

        generate_button = tk.Button(self.root, text="Сгенерировать файл", command=self.generate_file)
        generate_button.pack()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    file_generator = FileGenerator()
    file_generator.run()
