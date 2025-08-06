# Вот модифицированная версия проекта с вынесенными функциями конвертации для использования из других программ и сервисов:

# ```python
import os
import json
import sys
import subprocess
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import Optional, Tuple, List

class DocumentConverter:
    """Класс для конвертации документов, который можно использовать в других программах"""
    
    @staticmethod
    def find_libreoffice() -> Optional[str]:
        """Поиск пути к LibreOffice в системе"""
        paths = []
        
        if sys.platform == 'win32':
            paths = [
                'soffice',
                'libreoffice',
                r'C:\Program Files\LibreOffice\program\soffice.exe',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
            ]
        else:
            paths = [
                'libreoffice',
                'soffice',
                '/usr/bin/libreoffice',
                '/usr/bin/soffice',
                '/Applications/LibreOffice.app/Contents/MacOS/soffice'
            ]
        
        for path in paths:
            try:
                if sys.platform == 'win32' and ('\\' in path or '/' in path):
                    if os.path.exists(path):
                        return path
                else:
                    process = subprocess.run(['which', path], 
                                          stdout=subprocess.PIPE, 
                                          stderr=subprocess.PIPE)
                    if process.returncode == 0:
                        return path.strip()
            except:
                continue
        return None

    @staticmethod
    def convert_file(input_path: str, 
                   output_path: str, 
                   output_format: str = 'docx',
                   progress_callback: Optional[callable] = None) -> Tuple[bool, str]:
        """
        Конвертирует один файл
        
        Args:
            input_path: Путь к исходному файлу
            output_path: Путь для сохранения результата
            output_format: Формат для конвертации (docx, pdf, odt)
            progress_callback: Функция для отслеживания прогресса
            
        Returns:
            Tuple[bool, str]: (Успех/Неудача, Сообщение об ошибке)
        """
        try:
            libreoffice_path = DocumentConverter.find_libreoffice()
            if not libreoffice_path:
                return False, "LibreOffice не найден. Убедитесь, что он установлен."

            if progress_callback:
                progress_callback(10, f"Начало конвертации {os.path.basename(input_path)}")

            # Специальная обработка для HTML -> DOCX
            if input_path.lower().endswith('.html') and output_path.lower().endswith('.docx'):
                return DocumentConverter.convert_html_to_docx(input_path, output_path, progress_callback)

            # Подготовка команды
            if sys.platform == 'win32':
                command = f'"{libreoffice_path}" --headless --convert-to {output_format} --outdir "{os.path.dirname(output_path)}" "{input_path}"'
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                
                process = subprocess.run(
                    command,
                    startupinfo=startupinfo,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    shell=True,
                    encoding='utf-8',
                    errors='ignore'
                )
            else:
                command = [
                    libreoffice_path,
                    '--headless',
                    '--convert-to',
                    output_format,
                    '--outdir',
                    os.path.dirname(output_path),
                    input_path
                ]
                process = subprocess.run(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    encoding='utf-8',
                    errors='ignore'
                )

            if progress_callback:
                progress_callback(50, f"Обработка {os.path.basename(input_path)}")

            if process.returncode != 0:
                error_msg = process.stderr if process.stderr else "Неизвестная ошибка"
                return False, f"Ошибка конвертации: {error_msg}"

            # Проверяем результат
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            possible_output = os.path.join(
                os.path.dirname(output_path),
                f"{base_name}.{output_format.split(':')[0]}"
            )
            
            if os.path.exists(possible_output):
                if possible_output != output_path:
                    os.rename(possible_output, output_path)
                if progress_callback:
                    progress_callback(100, f"Успешно: {os.path.basename(input_path)}")
                return True, "Конвертация успешно завершена"
            else:
                return False, "Файл не был создан после конвертации"

        except Exception as e:
            return False, f"Ошибка при конвертации: {str(e)}"

    @staticmethod
    def convert_html_to_docx(input_path: str, 
                           output_path: str,
                           progress_callback: Optional[callable] = None) -> Tuple[bool, str]:
        """Специальный метод для конвертации HTML в DOCX"""
        try:
            if progress_callback:
                progress_callback(20, "Начало конвертации HTML в DOCX")

            # Создаем временный ODT файл
            temp_odt = os.path.join(tempfile.gettempdir(), f"temp_{os.path.basename(input_path)}.odt")
            
            # Конвертируем HTML -> ODT
            success, msg = DocumentConverter.convert_file(input_path, temp_odt, 'odt', progress_callback)
            if not success:
                return False, msg

            if progress_callback:
                progress_callback(60, "Конвертация ODT в DOCX")

            # Конвертируем ODT -> DOCX
            result = DocumentConverter.convert_file(temp_odt, output_path, 'docx', progress_callback)
            
            # Удаляем временный файл
            try:
                os.remove(temp_odt)
            except:
                pass
            
            return result
        
        except Exception as e:
            return False, f"Ошибка при конвертации HTML в DOCX: {str(e)}"

    @staticmethod
    def convert_directory(source_dir: str, 
                        target_dir: str, 
                        conversion_mode: str = "docx -> pdf",
                        progress_callback: Optional[callable] = None) -> Tuple[int, int, List[Tuple[str, str]]]:
        """
        Конвертирует все файлы в директории
        
        Args:
            source_dir: Исходная директория
            target_dir: Целевая директория
            conversion_mode: Режим конвертации
            progress_callback: Функция для отслеживания прогресса
            
        Returns:
            Tuple[int, int, List[Tuple[str, str]]]: 
                (Успешно, Всего, Список ошибок в формате (имя файла, ошибка))
        """
        try:
            os.makedirs(target_dir, exist_ok=True)
            
            input_ext = conversion_mode.split()[0]
            files = [f for f in os.listdir(source_dir) if f.lower().endswith(f'.{input_ext}')]
            
            if not files:
                return 0, 0, [("", "Нет файлов для конвертации")]
            
            success_count = 0
            errors = []
            
            for i, filename in enumerate(files, 1):
                if progress_callback:
                    progress_callback(int(i/len(files)*100), f"Обработка {filename}")
                
                input_file = os.path.join(source_dir, filename)
                output_ext = conversion_mode.split('->')[1].strip()
                output_file = os.path.join(target_dir, f"{os.path.splitext(filename)[0]}.{output_ext}")
                
                success, msg = DocumentConverter.convert_file(
                    input_file, 
                    output_file,
                    output_ext,
                    lambda p, m: progress_callback(p, m) if progress_callback else None
                )
                
                if success:
                    success_count += 1
                else:
                    errors.append((filename, msg))
            
            return success_count, len(files), errors
            
        except Exception as e:
            return 0, 0, [("", f"Ошибка при обработке директории: {str(e)}")]

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер документов (LibreOffice)")
        
        # Конфигурационный файл для сохранения настроек
        self.config_file = Path.home() / ".doc_converter_settings.json"
        print(self.config_file)
        # Загружаем сохраненные настройки
        self.settings = self.load_settings()
        
        # Переменные для хранения путей (инициализируем из сохраненных настроек)
        self.source_dir = tk.StringVar(value=self.settings.get("source_dir", ""))
        self.target_dir = tk.StringVar(value=self.settings.get("target_dir", ""))
        self.conversion_mode = tk.StringVar(value=self.settings.get("conversion_mode", "docx -> pdf"))
        
        # Привязываем событие закрытия по клавише Esc
        self.root.bind('<Escape>', lambda event: self.save_and_exit())
        
        # Привязываем сохранение настроек при закрытии окна
        self.root.protocol("WM_DELETE_WINDOW", self.save_and_exit)
        
        # Создаем интерфейс
        self.create_widgets()
    
    def load_settings(self):
        """Загружает настройки из файла"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Ошибка загрузки настроек: {e}")
        return {}
    
    def save_settings(self):
        """Сохраняет текущие настройки в файл"""
        try:
            settings = {
                "source_dir": self.source_dir.get(),
                "target_dir": self.target_dir.get(),
                "conversion_mode": self.conversion_mode.get()
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения настроек: {e}")
    
    def save_and_exit(self, event=None):
        """Сохраняет настройки и закрывает программу"""
        self.save_settings()
        self.root.destroy()
        
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
            "docx -> pdf",
            "docx -> odt",
            "odt -> docx",
            "odt -> pdf",
            "rtf -> docx",
            "rtf -> odt",
            "rtf -> pdf",
            "html -> docx",
            "html -> odt",
            "html -> pdf"
        ]
        tk.OptionMenu(self.root, self.conversion_mode, *modes).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        # Прогрессбар
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=3, padx=5, pady=5)
        
        # Метка для отображения текущего файла
        self.current_file_label = tk.Label(self.root, text="", fg="blue")
        self.current_file_label.grid(row=4, column=0, columnspan=3, padx=5, pady=5)
        
        # Кнопки
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        # tk.Button(button_frame, text="Конвертировать", command=self.convert_files, bg="green", fg="white").pack(side="left", padx=10)
        # tk.Button(button_frame, text="Закрыть", command=self.root.destroy, bg="red", fg="white").pack(side="right", padx=10)
        tk.Button(button_frame, text="Конвертировать", command=self.convert_files, fg="green").pack(side="left", padx=10)
        tk.Button(button_frame, text="Закрыть", command=self.root.destroy, fg="red").pack(side="right", padx=10)

    def browse_source(self):
        directory = filedialog.askdirectory()
        if directory:
            self.source_dir.set(directory)
            self.save_settings()  # Сохраняем сразу после выбора
    
    def browse_target(self):
        directory = filedialog.askdirectory()
        if directory:
            self.target_dir.set(directory)
            self.save_settings()  # Сохраняем сразу после выбора

    def update_progress(self, value, filename=""):
        self.progress['value'] = value
        self.current_file_label.config(text=f"Обработка: {filename}" if filename else "")
        self.root.update_idletasks()

    def convert_with_libreoffice(self, input_file, output_file, output_format):
        """Конвертация файлов с помощью LibreOffice в headless режиме"""
        try:
            libreoffice_path = self.find_libreoffice()
            if not libreoffice_path:
                raise Exception("LibreOffice не найден. Убедитесь, что он установлен.")

            # Для HTML файлов используем специальный подход
            if input_file.lower().endswith('.html') and output_file.lower().endswith('.docx'):
                return self.convert_html_to_docx(input_file, output_file)

            # Остальная логика конвертации для других форматов
            if sys.platform == 'win32':
                # Для Windows используем специальные параметры
                command = f'"{libreoffice_path}" --headless --convert-to {output_format} --outdir "{os.path.dirname(output_file)}" "{input_file}"'
                
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                
                process = subprocess.run(
                    command,
                    startupinfo=startupinfo,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    shell=True,
                    encoding='utf-8',
                    errors='ignore'
                )
            else:
                # Для других ОС
                command = [
                    libreoffice_path,
                    '--headless',
                    '--convert-to',
                    output_format,
                    '--outdir',
                    os.path.dirname(output_file),
                    input_file
                ]
                process = subprocess.run(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    encoding='utf-8',
                    errors='ignore'
                )

            if process.returncode != 0:
                error_msg = process.stderr if process.stderr else "Неизвестная ошибка"
                raise Exception(f"Ошибка конвертации: {error_msg}")

            # Проверяем результат
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            possible_output = os.path.join(
                os.path.dirname(output_file),
                f"{base_name}.{output_format.split(':')[0]}"
            )
            
            if os.path.exists(possible_output):
                if possible_output != output_file:
                    os.rename(possible_output, output_file)
                return True
            else:
                raise Exception("Файл не был создан после конвертации")

        except Exception as e:
            raise Exception(f"Ошибка при конвертации {input_file}: {str(e)}")

    def convert_html_to_docx(self, input_file, output_file):
        """Специальный метод для конвертации HTML в DOCX"""
        try:
            # Создаем временный ODT файл как промежуточный формат
            temp_odt = os.path.join(tempfile.gettempdir(), f"temp_{os.path.basename(input_file)}.odt")
            
            # Конвертируем HTML -> ODT
            self.convert_with_libreoffice(input_file, temp_odt, 'odt')
            
            # Конвертируем ODT -> DOCX
            self.convert_with_libreoffice(temp_odt, output_file, 'docx')
            
            # Удаляем временный файл
            try:
                os.remove(temp_odt)
            except:
                pass
            
            return True
        
        except Exception as e:
            # Если не получилось через промежуточный ODT, пробуем альтернативный метод
            try:
                return self.convert_html_to_docx_fallback(input_file, output_file)
            except Exception as fallback_e:
                raise Exception(f"Основной и альтернативный методы не сработали: {str(e)} | {str(fallback_e)}")

    def convert_html_to_docx_fallback(self, input_file, output_file):
        """Альтернативный метод конвертации HTML в DOCX"""
        try:
            from htmldocx import HtmlToDocx
            from docx import Document
            
            with open(input_file, 'r', encoding='utf-8') as f:
                html_content = f.read()
            
            # Создаем новый документ
            doc = Document()
            
            # Инициализируем парсер
            parser = HtmlToDocx()
            
            # Добавляем HTML-контент в документ
            parser.add_html_to_document(html_content, doc)
            
            # Сохраняем результат
            doc.save(output_file)
            return True
            
        except ImportError:
            raise Exception("Для альтернативного метода требуется установка python-docx и htmldocx")
        except Exception as e:
            raise Exception(f"Ошибка в альтернативном методе: {str(e)}")

    def get_libreoffice_format(self, mode):
        """Получаем формат для LibreOffice на основе режима конвертации"""
        format_map = {
            'pdf': 'pdf',
            'docx': 'docx',
            'odt': 'odt',
            'rtf': 'docx',  # Изменено для RTF -> DOCX
            'html': 'docx'
        }
        return format_map.get(mode.split('->')[1].strip(), "")

    def find_libreoffice(self):
        """Поиск пути к LibreOffice в системе"""
        # Проверяем стандартные пути
        paths = []
        
        if sys.platform == 'win32':
            # Пути для Windows
            paths = [
                'soffice',
                'libreoffice',
                r'C:\Program Files\LibreOffice\program\soffice.exe',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
            ]
        else:
            # Пути для Linux/Mac
            paths = [
                'libreoffice',
                'soffice',
                '/usr/bin/libreoffice',
                '/usr/bin/soffice',
                '/Applications/LibreOffice.app/Contents/MacOS/soffice'
            ]
        
        for path in paths:
            try:
                # Для Windows проверяем существование файла
                if sys.platform == 'win32' and ('\\' in path or '/' in path):
                    if os.path.exists(path):
                        return path
                else:
                    # Для других ОС проверяем через which
                    process = subprocess.run(['which', path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    if process.returncode == 0:
                        return path.strip()
            except:
                continue
        return None

    def get_output_extension(self, mode):
        """Получаем расширение файла на основе режима конвертации"""
        format_map = {
            'pdf': 'pdf',
            'docx': 'docx',
            'odt': 'odt'
        }
        return format_map.get(mode.split('->')[1].strip(), "")

    def convert_files(self):
        """Обработка файлов с сохранением последнего режима конвертации"""
        try:
            # Сохраняем текущий режим перед конвертацией
            self.save_settings()
            
            # Остальная логика конвертации...
            source = self.source_dir.get()
            target = self.target_dir.get()
            mode = self.conversion_mode.get()
            
            if not source or not target:
                messagebox.showerror("Ошибка", "Выберите исходную и целевую директории")
                return
        
            # Создаем целевую директорию, если ее нет
            os.makedirs(target, exist_ok=True)
            
            # Получаем список файлов для обработки
            input_ext = mode.split()[0]
            files = [f for f in os.listdir(source) if f.lower().endswith(f'.{input_ext}')]
            
            if not files:
                messagebox.showwarning("Предупреждение", f"Нет файлов .{input_ext} для конвертации в выбранной директории")
                return
            
            # Настраиваем прогрессбар
            self.progress['maximum'] = len(files)
            self.progress['value'] = 0
            self.current_file_label.config(text="")
            
            # Обрабатываем файлы
            success_count = 0
            for i, filename in enumerate(files, 1):
                self.update_progress(i, filename)
                
                input_file = os.path.join(source, filename)
                output_ext = self.get_output_extension(mode)
                output_file = os.path.join(target, f"{os.path.splitext(filename)[0]}.{output_ext}")
                
                try:
                    # Определяем формат для LibreOffice
                    output_format = self.get_libreoffice_format(mode)
                    
                    # Выполняем конвертацию
                    if self.convert_with_libreoffice(input_file, output_file, output_format):
                        success_count += 1
                except Exception as e:
                    print(f"Ошибка при обработке {filename}: {str(e)}")
            
            messagebox.showinfo("Успех", f"Конвертация завершена!\nУспешно: {success_count} из {len(files)}")
            self.update_progress(0)  # Сбрасываем прогрессбар
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
            self.update_progress(0)  # Сбрасываем прогрессбар


def convert_file_cli(input_path: str, output_path: str, output_format: str = 'docx'):
    """Функция для вызова из командной строки"""
    success, msg = DocumentConverter.convert_file(input_path, output_path, output_format)
    print(f"Результат: {'Успешно' if success else 'Ошибка'}")
    print(msg)
    return 0 if success else 1


if __name__ == "__main__":
    # Если вызвано из командной строки с аргументами
    if len(sys.argv) > 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        format = sys.argv[3] if len(sys.argv) > 3 else 'docx'
        sys.exit(convert_file_cli(input_file, output_file, format))
    else:
        # Запуск GUI
        try:
            root = tk.Tk()
            app = FileConverterApp(root)
            root.mainloop()
        except Exception as e:
            print(f"Ошибка: {str(e)}")
            input("Нажмите Enter для выхода...")

# Из другого Python-скрипта:

# python
# from converter import DocumentConverter

# # Конвертация одного файла
# success, message = DocumentConverter.convert_file("input.doc", "output.pdf", "pdf")

# # Конвертация всей директории
# success_count, total_files, errors = DocumentConverter.convert_directory(
#     "source_folder",
#     "target_folder",
#     "docx -> pdf",
#     lambda progress, msg: print(f"{progress}%: {msg}")
# )
# Из командной строки:

# bash
# python converter.py input.html output.docx
# python converter.py input.odt output.pdf pdf
# Как модуль в сервисе:

# python
# from fastapi import FastAPI
# from converter import DocumentConverter

# app = FastAPI()

# @app.post("/convert")
# async def convert_file(input_path: str, output_path: str):
#     success, message = DocumentConverter.convert_file(input_path, output_path)
#     return {"success": success, "message": message}
