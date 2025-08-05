import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from comtypes import client
import pdfkit
import pypandoc

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер документов")
        
        # Переменные для хранения путей
        self.source_dir = tk.StringVar()
        self.target_dir = tk.StringVar()
        self.conversion_mode = tk.StringVar(value="rtf -> docx")
        
        # Создаем интерфейс
        self.create_widgets()
        
    def create_widgets(self):
        # Поля для выбора директорий
        tk.Label(self.root, text="Исходная директория:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(self.root, textvariable=self.source_dir, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.root, text="Обзор...", command=self.browse_source).grid(row=0, column=2, padx=5, pady=5)
        
        tk.Label(self.root, text="Целевая директория:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(self.root, textvariable=self.target_dir, width=50).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self.root, text="Обзор...", command=self.browse_target).grid(row=1, column=2, padx=5, pady=5)
        
        # Выбор режима конвертации
        tk.Label(self.root, text="Режим конвертации:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        modes = [
            "rtf -> docx",
            "docx -> pdf",
            "rtf -> pdf",
            "html -> docx",
            "html -> rtf"
        ]
        tk.OptionMenu(self.root, self.conversion_mode, *modes).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        # Кнопка конвертации
        tk.Button(self.root, text="Конвертировать", command=self.convert_files, bg="green", fg="white").grid(row=3, column=1, pady=10)
        
    def browse_source(self):
        directory = filedialog.askdirectory()
        if directory:
            self.source_dir.set(directory)
    
    def browse_target(self):
        directory = filedialog.askdirectory()
        if directory:
            self.target_dir.set(directory)
    
    def convert_files(self):
        source = self.source_dir.get()
        target = self.target_dir.get()
        mode = self.conversion_mode.get()
        
        if not source or not target:
            messagebox.showerror("Ошибка", "Выберите исходную и целевую директории")
            return
        
        try:
            # Создаем целевую директорию, если ее нет
            os.makedirs(target, exist_ok=True)
            
            # Обрабатываем файлы в зависимости от режима
            if mode == "rtf -> docx":
                self.convert_rtf_to_docx(source, target)
            elif mode == "docx -> pdf":
                self.convert_docx_to_pdf(source, target)
            elif mode == "rtf -> pdf":
                self.convert_rtf_to_pdf(source, target)
            elif mode == "html -> docx":
                self.convert_html_to_docx(source, target)
            elif mode == "html -> rtf":
                self.convert_html_to_rtf(source, target)
                
            messagebox.showinfo("Успех", "Конвертация завершена успешно!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
    
    def convert_rtf_to_docx(self, source, target):
        """Конвертация RTF в DOCX"""
        word = client.CreateObject("Word.Application")
        word.Visible = False
        
        for filename in os.listdir(source):
            if filename.lower().endswith('.rtf'):
                input_file = os.path.join(source, filename)
                output_file = os.path.join(target, os.path.splitext(filename)[0] + '.docx')
                
                doc = word.Documents.Open(input_file)
                doc.SaveAs(output_file, FileFormat=16)  # 16 = wdFormatDocumentDefault
                doc.Close()
        
        word.Quit()
    
    def convert_docx_to_pdf(self, source, target):
        """Конвертация DOCX в PDF"""
        word = client.CreateObject("Word.Application")
        word.Visible = False
        
        for filename in os.listdir(source):
            if filename.lower().endswith('.docx'):
                input_file = os.path.join(source, filename)
                output_file = os.path.join(target, os.path.splitext(filename)[0] + '.pdf')
                
                doc = word.Documents.Open(input_file)
                doc.SaveAs(output_file, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
        
        word.Quit()
    
    def convert_rtf_to_pdf(self, source, target):
        """Конвертация RTF в PDF"""
        word = client.CreateObject("Word.Application")
        word.Visible = False
        
        for filename in os.listdir(source):
            if filename.lower().endswith('.rtf'):
                input_file = os.path.join(source, filename)
                output_file = os.path.join(target, os.path.splitext(filename)[0] + '.pdf')
                
                doc = word.Documents.Open(input_file)
                doc.SaveAs(output_file, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
        
        word.Quit()
    
    def convert_html_to_docx(self, source, target):
        """Конвертация HTML в DOCX"""
        for filename in os.listdir(source):
            if filename.lower().endswith('.html'):
                input_file = os.path.join(source, filename)
                output_file = os.path.join(target, os.path.splitext(filename)[0] + '.docx')
                
                pypandoc.convert_file(input_file, 'docx', outputfile=output_file)
    
    def convert_html_to_rtf(self, source, target):
        """Конвертация HTML в RTF"""
        for filename in os.listdir(source):
            if filename.lower().endswith('.html'):
                input_file = os.path.join(source, filename)
                output_file = os.path.join(target, os.path.splitext(filename)[0] + '.rtf')
                
                pypandoc.convert_file(input_file, 'rtf', outputfile=output_file)

if __name__ == "__main__":
    # Проверяем зависимости
    try:
        import comtypes
        import pdfkit
        import pypandoc
    except ImportError:
        print("Установите необходимые зависимости:")
        print("pip install comtypes pdfkit pypandoc")
        exit(1)
    
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()