import os
import sorter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def select_source_dir():
    """ Открывает диалог выбора исходного каталога """
    dir_name = filedialog.askdirectory()
    if dir_name:
        source_entry.delete(0, tk.END)
        source_entry.insert(0, dir_name)


def select_target_dir():
    """ Открывает диалог выбора целевого каталога """
    dir_name = filedialog.askdirectory()
    if dir_name:
        target_entry.delete(0, tk.END)
        target_entry.insert(0, dir_name)


def execute_process():
    """ Выполняет основную операцию (копирование, перенос и т.п.) """
    source_dir = source_entry.get()
    target_dir = target_entry.get()
    mode = modes_dict.get(mode_combobox.get())
    
    if not source_dir or not target_dir:
        messagebox.showwarning("Ошибка", "Укажите оба каталога!")
        return
    
    # Добавьте сюда вашу логику для выбранной операции
    print(f"Источник: {source_dir}, Цель: {target_dir}, Режим: {mode}")
    result_message = sorter.photosorter(source_dir,target_dir,mode)
    messagebox.showinfo("Информация", f"Операция '{mode}' выполнена.\n{result_message}")


# Настройка основного окна
root = tk.Tk()
root.title("Сортировка файлов. созданием поддиректорий с датой создания, переворотом и обратным геокодированием на основе EXIF данных")
root.geometry("800x300")

# Исходный каталог
initial_directory = os.getcwd()
source_label = tk.Label(root, text="Выберите исходную директорию с фотографиями:")
source_label.pack()
source_entry = tk.Entry(root, width=100)
source_entry.insert(0, initial_directory) # Устанавливаем начальное значение
source_entry.pack()
source_browse_button = tk.Button(root, text="Обзор...", command=select_source_dir)
source_browse_button.pack()

# Целевой каталог
target_label = tk.Label(root, text="Целевой каталог:")
target_label.pack()
target_entry = tk.Entry(root, width=100)
target_entry.pack()
target_browse_button = tk.Button(root, text="Обзор...", command=select_target_dir)
target_browse_button.pack()

# Выбор режима
mode_label = tk.Label(root, text="Режим:")
mode_label.pack()

#modes = ["Только посчитать файлы", "Создавать подкаталоги с датой создания и переворачивать", "И прикладывать к дате и адрес на основе геолокации"]
# Готовим словарь режимов
modes_dict = {
    "Только посчитать": "count",
    "Создавать подкаталоги с датой создания и переворачивать": "create",
    "И прикладывать к дате и адрес на основе геолокации": "geotag"
}

# Список отображаемых значений
modes_list = list(modes_dict.keys())

mode_combobox = ttk.Combobox(root, values=modes_list,width=100)
mode_combobox.current(0)  # Устанавливаем первый пункт по умолчанию
mode_combobox.pack()

# Кнопка запуска
execute_button = tk.Button(root, text="Запустить", command=execute_process)
execute_button.place(x=350, y=230)
#execute_button.pack()

# Кнопка закрытия
exit_button = tk.Button(root, text="Выход", command=root.destroy)
exit_button.place(x=420, y=230)
#exit_button.pack()

def on_escape(event):
    root.destroy()  # Закрываем главное окно приложения

root.bind('<Escape>', on_escape)

# Запуск основного цикла
root.mainloop()
