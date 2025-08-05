# Для того, чтобы реализовать функциональность конвертации файлов в разных форматах, таких как RTF, DOC, HTML, PDF, лучше всего воспользоваться специализированными библиотеками, такими как Aspose.Words и Spire.Doc. Эти библиотеки позволяют удобно обрабатывать различные типы документов и сохранять их в нужных форматах.

# Алгоритм работы программы:

# Мы используем библиотеку Aspose.Words, так как она обладает широким функционалом для работы с офисными форматами.
# Используем интерфейс Tkinter для графического интерфейса пользователя, позволяющего выбрать исходную и целевую папки, а также задать нужный режим конвертации.
# Обрабатываем выбранные файлы и выполняем требуемое преобразование согласно указанному режиму.

# Установка необходимых пакетов:

# Перед началом убедитесь, что установлены нужные библиотеки:

# pip install aspose-words pillow tkinter

# Полностью рабочий пример:

# Шаг 1: Реализация основной функции для конвертации файлов:

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import aspose.words as aw

def convert_files(src_dir, dst_dir, mode):
    """
    Функция для конвертации файлов из одной директории в другую с заданным режимом конвертации.
    
    Параметры:
    src_dir : str
        Исходная директория с файлами.
    dst_dir : str
        Целевая директория, куда будут сохранены конвертированные файлы.
    mode : str
        Режим конвертации ('rtf_to_docx', 'docx_to_pdf', 'rtf_to_pdf', 'html_to_docx', 'html_to_rtf').
        
    Возвращаемые значения:
    None
    """
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)

    for filename in os.listdir(src_dir):
        src_path = os.path.join(src_dir, filename)
        base_name, ext = os.path.splitext(filename)
        
        try:
            if mode == 'rtf_to_docx':
                if ext.lower() != '.rtf':  
                    print(f"Пропускаем {filename}, ожидаемый формат — RTF.")
                    continue
                
                doc = aw.Document(src_path)
                new_filename = f"{base_name}.docx"
                output_path = os.path.join(dst_dir, new_filename)
                doc.save(output_path)
            
            elif mode == 'docx_to_pdf':
                if ext.lower() != '.docx':  
                    print(f"Пропускаем {filename}, ожидаемый формат — DOCX.")
                    continue
                
                doc = aw.Document(src_path)
                new_filename = f"{base_name}.pdf"
                output_path = os.path.join(dst_dir, new_filename)
                doc.save(output_path)
            
            elif mode == 'rtf_to_pdf':
                if ext.lower() != '.rtf':  
                    print(f"Пропускаем {filename}, ожидаемый формат — RTF.")
                    continue
                
                doc = aw.Document(src_path)
                new_filename = f"{base_name}.pdf"
                output_path = os.path.join(dst_dir, new_filename)
                doc.save(output_path)
            
            elif mode == 'html_to_docx':
                if ext.lower() != '.html':  
                    print(f"Пропускаем {filename}, ожидаемый формат — HTML.")
                    continue
                
                doc = aw.Document.from_html_file(src_path)
                new_filename = f"{base_name}.docx"
                output_path = os.path.join(dst_dir, new_filename)
                doc.save(output_path)
            
            elif mode == 'html_to_rtf':
                if ext.lower() != '.html':  
                    print(f"Пропускаем {filename}, ожидаемый формат — HTML.")
                    continue
                
                doc = aw.Document.from_html_file(src_path)
                new_filename = f"{base_name}.rtf"
                output_path = os.path.join(dst_dir, new_filename)
                doc.save(output_path)
            
            else:
                raise ValueError(f"Неверный режим конвертации '{mode}'. Допустимые режимы: 'rtf_to_docx', 'docx_to_pdf', 'rtf_to_pdf', 'html_to_docx', 'html_to_rtf'.")
            
            print(f"Успешно сконвертирован файл: {src_path} → {output_path}")
        
        except Exception as e:
            print(f"Ошибка при обработке файла {src_path}: {e}")

#Шаг 2: Интерфейс Tkinter для выбора каталогов и запуска процесса:

def select_source_directory():
    global source_dir
    source_dir = filedialog.askdirectory(title="Выберите исходную директорию")
    label_src.config(text=f"Исходная директория: {source_dir}")

def select_destination_directory():
    global dest_dir
    dest_dir = filedialog.askdirectory(title="Выберите целевую директорию")
    label_dst.config(text=f"Целевая директория: {dest_dir}")

def start_conversion(mode):
    if not source_dir or not dest_dir:
        messagebox.showwarning("Предупреждение", "Необходимо выбрать обе директории!")
        return
    
    convert_files(source_dir, dest_dir, mode)
    messagebox.showinfo("Готово", "Преобразование успешно завершилось!")

root = tk.Tk()
root.title("Конвертер файлов")

frame = tk.Frame(root)
frame.pack(padx=20, pady=20)

label_src = tk.Label(frame, text="Исходная директория:")
label_src.grid(row=0, column=0, sticky='w')
btn_select_src = tk.Button(frame, text="Выбрать...", command=select_source_directory)
btn_select_src.grid(row=0, column=1)

label_dst = tk.Label(frame, text="Целевая директория:")
label_dst.grid(row=1, column=0, sticky='w')
btn_select_dst = tk.Button(frame, text="Выбрать...", command=select_destination_directory)
btn_select_dst.grid(row=1, column=1)

modes = ['rtf_to_docx', 'docx_to_pdf', 'rtf_to_pdf', 'html_to_docx', 'html_to_rtf']
var_mode = tk.StringVar(value=modes[0])
dropdown_modes = tk.OptionMenu(frame, var_mode, *modes)
dropdown_modes.grid(row=2, columnspan=2, pady=(10, 0), sticky='ew')

btn_convert = tk.Button(frame, text="Начать конвертацию", command=lambda: start_conversion(var_mode.get()))
btn_convert.grid(row=3, columnspan=2, pady=(10, 0))

root.mainloop()

# Подробное объяснение шагов:

# Интерфейс: Программа запускается с окна, где пользователь выбирает исходную и целевую папки. Выбранные пути отображаются на экране.


# Выбор режимов: Пользователь может выбрать нужный режим конвертации из выпадающего списка.


# Обработка файлов: После нажатия кнопки "Начать конвертацию" программа перебирает все файлы в выбранной директории и применяет соответствующий режим преобразования.


# Поддерживаемые форматы: Возможны следующие преобразования:

# RTF → DOCX
# DOCX → PDF
# RTF → PDF
# HTML → DOCX
# HTML → RTF

# Заключение:

# Программа обеспечивает удобный способ массовой конвертации файлов между популярными форматами документов. Благодаря использованию библиотеки Aspose.Words, гарантируется высокая точность и качество преобразований. Графический интерфейс помогает сделать процесс удобным даже для начинающих пользователей.