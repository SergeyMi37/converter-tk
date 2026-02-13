import os, json, sys, subprocess, tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

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
        
        # Выбор режима конвертации:
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
            "html -> pdf",
            "pdf -> txt",
            "doc -> txt",
            "docx -> txt"
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

    def convert_pdf_to_txt(self, input_file, output_file):
        """Конвертация PDF в текстовый файл с помощью PyPDF2"""
        try:
            from PyPDF2 import PdfReader
            
            with open(input_file, 'rb') as f:
                reader = PdfReader(f)
                
                text = ""
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    text += page.extract_text()
                
                with open(output_file, 'w', encoding='utf-8') as txt_file:
                    txt_file.write(text)
                
                return True
                
        except ImportError:
            raise Exception("Для конвертации PDF в TXT требуется PyPDF2. Установите: pip install PyPDF2")
        except Exception as e:
            raise Exception(f"Ошибка при конвертации PDF в TXT: {str(e)}")

    def convert_doc_to_txt(self, input_file, output_file):
        """Конвертация DOC/DOCX в текстовый файл с помощью LibreOffice"""
        try:
            # Используем LibreOffice для конвертации в текстовый формат
            output_format = "txt"
            return self.convert_with_libreoffice(input_file, output_file, output_format)
        except Exception as e:
            raise Exception(f"Ошибка при конвертации DOC/DOCX в TXT: {str(e)}")

    def get_libreoffice_format(self, mode):
        """Получаем формат для LibreOffice на основе режима конвертации"""
        format_map = {
            'pdf': 'pdf',
            'docx': 'docx',
            'odt': 'odt',
            'rtf': 'docx',  # Изменено для RTF -> DOCX
            'html': 'docx',
            'txt': 'txt',
            'doc': 'txt',
            'docx': 'txt'
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
            'odt': 'odt',
            'txt': 'txt',
            'doc': 'txt',
            'docx': 'txt'
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
            input_ext = mode.split(' -> ')[0].strip()
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
                    if mode == "pdf -> txt":
                        if self.convert_pdf_to_txt(input_file, output_file):
                            success_count += 1
                    elif mode in ("doc -> txt", "docx -> txt"):
                        if self.convert_doc_to_txt(input_file, output_file):
                            success_count += 1
                    else:
                        if self.convert_with_libreoffice(input_file, output_file, output_format):
                            success_count += 1
                except Exception as e:
                    print(f"Ошибка при обработке {filename}: {str(e)}")
            
            messagebox.showinfo("Успех", f"Конвертация завершена!\nУспешно: {success_count} из {len(files)}")
            self.update_progress(0)  # Сбрасываем прогрессбар
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
            self.update_progress(0)  # Сбрасываем прогрессбар

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = FileConverterApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Ошибка: {str(e)}")
        print("Убедитесь, что LibreOffice установлен и доступен в PATH")
        input("Нажмите Enter для выхода...")
