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
        print(f"Файл настроек: {self.config_file}")
        
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
        
        # Проверяем доступность OCR инструментов
        self.check_ocr_availability()
        
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
    
    def check_ocr_availability(self):
        """Проверяет доступность OCR инструментов"""
        print("\n" + "="*50)
        print("ПРОВЕРКА OCR ИНСТРУМЕНТОВ")
        print("="*50)
        
        # Проверяем Tesseract
        try:
            import pytesseract
            tesseract_version = pytesseract.get_tesseract_version()
            print(f"✓ Tesseract найден, версия: {tesseract_version}")
            
            # Проверяем языки
            languages = self.check_tesseract_languages()
            if 'rus' in languages:
                print("✓ Русский язык ДОСТУПЕН для OCR")
            else:
                print("✗ Русский язык НЕ ДОСТУПЕН для OCR")
                print("  Для распознавания русского текста установите русский язык в Tesseract")
                print("  https://github.com/UB-Mannheim/tesseract/wiki")
            
            if 'eng' in languages:
                print("✓ Английский язык доступен для OCR")
            else:
                print("✗ Английский язык НЕ ДОСТУПЕН для OCR")
                
        except ImportError:
            print("✗ pytesseract не установлен")
            print("  Установите: pip install pytesseract")
        except Exception as e:
            print(f"✗ Ошибка при проверке Tesseract: {e}")
            print("  Tesseract OCR не найден в системе")
            print("  Скачайте и установите: https://github.com/UB-Mannheim/tesseract/wiki")
        
        # Проверяем ocrmypdf
        try:
            import ocrmypdf
            print(f"✓ ocrmypdf установлен")
        except ImportError:
            print("✗ ocrmypdf не установлен (опционально)")
            print("  Установите: pip install ocrmypdf[tesseract]")
        
        # Проверяем PyMuPDF
        try:
            import fitz
            print(f"✓ PyMuPDF установлен")
        except ImportError:
            print("✗ PyMuPDF не установлен")
            print("  Установите: pip install PyMuPDF")
        
        print("="*50 + "\n")
    
    def check_tesseract_languages(self):
        """Проверяет доступные языки в Tesseract"""
        try:
            import pytesseract
            languages = pytesseract.get_languages()
            print(f"Доступные языки Tesseract: {languages}")
            return languages
        except Exception as e:
            print(f"Не удалось проверить языки Tesseract: {e}")
            return []
    
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
            "pdf -> txt (OCR)",
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
            return self.extract_text_from_pdf(input_file, output_file)
                
        except ImportError:
            raise Exception("Для конвертации PDF в TXT требуется PyPDF2. Установите: pip install PyPDF2")
        except Exception as e:
            # Если PyPDF2 не смог извлечь текст, пробуем использовать OCR
            print(f"PyPDF2 не смог извлечь текст, пробуем OCR: {str(e)}")
            return self.convert_pdf_to_txt_ocr(input_file, output_file)

    def convert_pdf_to_txt_ocr(self, input_file, output_file):
        """Конвертация PDF в текстовый файл с использованием OCR для сканированных PDF"""
        print(f"\nНачинаем OCR-конвертацию PDF: {input_file}")
        temp_pdf_path = None
        
        # Настройка языков для OCR - русский и английский
        ocr_langs = "rus+eng"
        print(f"Языки OCR: {ocr_langs}")
        
        try:
            # Пробуем использовать OCRmyPDF (основной метод)
            import ocrmypdf
            
            print(f"Используем OCRmyPDF для обработки PDF с языком: {ocr_langs}")
            # Создаем временный файл для PDF с распознанным текстом
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                temp_pdf_path = temp_pdf.name
            
            # Используем OCRmyPDF с указанием языков
            ocrmypdf.ocr(
                input_file, 
                temp_pdf_path, 
                deskew=True,
                language=ocr_langs,  # Передаем языки для OCR
                force_ocr=True,  # Принудительно применяем OCR даже если есть текст
                output_type='pdf'
            )
            
            # Извлекаем текст из обработанного PDF
            success = self.extract_text_from_pdf(temp_pdf_path, output_file)
            
            # Удаляем временный файл
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                os.unlink(temp_pdf_path)
            
            print("OCR-конвертация успешно завершена")
            return success
            
        except ImportError:
            print("OCRmyPDF не установлен, пробуем альтернативный метод")
            # Если нет ocrmypdf, пробуем прямой метод с pytesseract
            return self.convert_pdf_to_txt_direct_ocr(input_file, output_file, ocr_langs)
            
        except Exception as e:
            error_msg = str(e)
            print(f"Ошибка при использовании OCRmyPDF: {error_msg}")
            
            # Если не хватает внешних программ
            if "Could not find program" in error_msg:
                if "tesseract" in error_msg:
                    print("Tesseract не найден в системе")
                    message = "Для OCR требуется Tesseract. Скачайте с https://github.com/UB-Mannheim/tesseract/wiki\n"
                    message += "При установке выберите языки: English, Russian"
                    raise Exception(message)
                elif "ghostscript" in error_msg.lower():
                    print("Ghostscript не найден, пробуем прямой метод")
                    return self.convert_pdf_to_txt_direct_ocr(input_file, output_file, ocr_langs)
            
            # Очищаем временные файлы
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.unlink(temp_pdf_path)
                except:
                    pass
            
            # Пробуем прямой метод OCR как запасной вариант
            print("Пробуем прямой OCR метод")
            return self.convert_pdf_to_txt_direct_ocr(input_file, output_file, ocr_langs)

    def convert_pdf_to_txt_direct_ocr(self, input_file, output_file, languages="rus+eng"):
        """
        Прямая OCR конвертация PDF в текст без использования ocrmypdf
        Использует PyMuPDF и pytesseract с поддержкой русского языка
        """
        print(f"\nИспользуем прямой OCR метод с языками: {languages}")
        
        try:
            # Проверяем наличие необходимых библиотек
            import fitz  # PyMuPDF
            import pytesseract
            from PIL import Image, ImageEnhance
            import io
            
            # Проверяем доступность указанных языков в Tesseract
            try:
                available_langs = pytesseract.get_languages()
                print(f"Доступные языки Tesseract: {available_langs}")
                
                # Проверяем наличие русского языка
                lang_parts = languages.split('+')
                available_parts = []
                for lang in lang_parts:
                    if lang in available_langs:
                        available_parts.append(lang)
                    else:
                        print(f"Язык '{lang}' не установлен в Tesseract")
                
                if available_parts:
                    languages = '+'.join(available_parts)
                    print(f"Используем доступные языки: {languages}")
                else:
                    languages = 'eng'
                    print(f"Нет запрошенных языков, используем: {languages}")
                    
            except Exception as e:
                print(f"Не удалось проверить языки Tesseract: {e}")
            
            # Открываем PDF
            doc = fitz.open(input_file)
            total_pages = len(doc)
            full_text = []
            
            # Настройки OCR с поддержкой русского языка
            custom_config = f'--oem 3 --psm 6 -l {languages}'
            alt_config = '--oem 1 --psm 3'  # Альтернативные настройки
            
            print(f"Обрабатываем {total_pages} страниц с OCR")
            
            for page_num in range(total_pages):
                print(f"OCR обработка страницы {page_num + 1}/{total_pages}")
                page = doc[page_num]
                
                # Пытаемся извлечь встроенный текст
                page_text = page.get_text()
                
                # Если текста мало (менее 100 символов) или это сканированная страница
                if len(page_text.strip()) < 100:
                    print(f"Страница {page_num + 1}: выполняем OCR (текста мало: {len(page_text.strip())} символов)")
                    
                    # Получаем изображение страницы с высоким разрешением
                    zoom_matrix = fitz.Matrix(2.0, 2.0)  # Увеличиваем разрешение в 2 раза
                    pix = page.get_pixmap(matrix=zoom_matrix, alpha=False)
                    
                    # Конвертируем в PIL Image
                    img_data = pix.tobytes("png")
                    image = Image.open(io.BytesIO(img_data))
                    
                    # Улучшаем изображение для OCR
                    # Конвертируем в оттенки серого
                    if image.mode != 'L':
                        image = image.convert('L')
                    
                    # Увеличиваем контраст
                    enhancer = ImageEnhance.Contrast(image)
                    image = enhancer.enhance(2.0)
                    
                    # Увеличиваем резкость
                    enhancer = ImageEnhance.Sharpness(image)
                    image = enhancer.enhance(1.5)
                    
                    # Выполняем OCR с указанными языками
                    try:
                        ocr_text = pytesseract.image_to_string(
                            image, 
                            config=custom_config
                        )
                        if ocr_text.strip():
                            full_text.append(ocr_text)
                            print(f"Страница {page_num + 1}: OCR успешен, получено {len(ocr_text)} символов")
                        else:
                            print(f"Страница {page_num + 1}: OCR не дал результатов")
                            # Пробуем с другими настройками
                            ocr_text = pytesseract.image_to_string(
                                image, 
                                config=alt_config,
                                lang=languages.split('+')[0] if '+' in languages else languages
                            )
                            if ocr_text.strip():
                                full_text.append(ocr_text)
                                print(f"Страница {page_num + 1}: альтернативный OCR успешен")
                    except Exception as ocr_error:
                        print(f"Ошибка OCR на странице {page_num + 1}: {ocr_error}")
                        # Если русский не работает, пробуем только английский
                        if 'rus' in languages:
                            try:
                                print(f"Страница {page_num + 1}: пробуем только английский")
                                ocr_text = pytesseract.image_to_string(
                                    image, 
                                    config='--oem 3 --psm 6',
                                    lang='eng'
                                )
                                if ocr_text.strip():
                                    full_text.append(ocr_text)
                            except:
                                pass
                else:
                    # Используем встроенный текст
                    full_text.append(page_text)
                    print(f"Страница {page_num + 1}: использован встроенный текст ({len(page_text.strip())} символов)")
            
            # Сохраняем результат с BOM для лучшей совместимости с Windows
            with open(output_file, 'w', encoding='utf-8-sig') as txt_file:
                txt_file.write('\n\n'.join(full_text))
            
            doc.close()
            print(f"Прямая OCR-конвертация успешно завершена. Сохранено в: {output_file}")
            return True
            
        except ImportError as e:
            missing_packages = []
            try:
                import fitz
            except ImportError:
                missing_packages.append("PyMuPDF (pip install PyMuPDF)")
            
            try:
                import pytesseract
            except ImportError:
                missing_packages.append("pytesseract (pip install pytesseract)")
            
            try:
                from PIL import Image
            except ImportError:
                missing_packages.append("Pillow (pip install Pillow)")
            
            if missing_packages:
                error_msg = f"Для OCR требуются пакеты:\n" + "\n".join(missing_packages)
                error_msg += "\n\nТакже требуется Tesseract OCR (https://github.com/UB-Mannheim/tesseract/wiki)"
                error_msg += "\nПри установке Tesseract выберите языки: English, Russian"
                raise Exception(error_msg)
            else:
                raise Exception(f"Ошибка при импорте: {str(e)}")
        
        except Exception as e:
            raise Exception(f"Ошибка при прямой OCR-конвертации: {str(e)}")

    def extract_text_from_pdf(self, pdf_path, output_file):
        """Вспомогательная функция для извлечения текста из PDF файла"""
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            raise Exception("PyPDF2 не установлен. Установите: pip install PyPDF2")
        
        with open(pdf_path, 'rb') as f:
            reader = PdfReader(f)
            text = ""
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n\n"
            
            with open(output_file, 'w', encoding='utf-8-sig') as txt_file:
                txt_file.write(text)
        
        return True

    def convert_doc_to_txt(self, input_file, output_file):
        """Конвертация DOC/DOCX в текстовый файл с помощью LibreOffice"""
        try:
            # Используем LibreOffice для конвертации в текстовый формат
            output_format = "txt:Text"
            return self.convert_with_libreoffice(input_file, output_file, output_format)
        except Exception as e:
            raise Exception(f"Ошибка при конвертации DOC/DOCX в TXT: {str(e)}")

    def get_libreoffice_format(self, mode):
        """Получаем формат для LibreOffice на основе режима конвертации"""
        target_format = mode.split(' -> ')[1].strip()
        
        format_map = {
            'pdf': 'pdf',
            'docx': 'docx',
            'odt': 'odt',
            'txt': 'txt:Text',
            'doc': 'txt:Text'
        }
        
        # Специальные случаи для конвертации в txt
        if target_format == 'txt' and mode.startswith(('doc', 'docx')):
            return 'txt:Text'
        
        return format_map.get(target_format, target_format)

    def find_libreoffice(self):
        """Поиск пути к LibreOffice в системе"""
        # Проверяем стандартные пути
        paths = []
        
        if sys.platform == 'win32':
            # Пути для Windows
            program_files = os.environ.get('ProgramFiles', 'C:\\Program Files')
            program_files_x86 = os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)')
            
            paths = [
                'soffice',
                'libreoffice',
                os.path.join(program_files, 'LibreOffice', 'program', 'soffice.exe'),
                os.path.join(program_files_x86, 'LibreOffice', 'program', 'soffice.exe'),
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
                '/usr/local/bin/libreoffice',
                '/usr/local/bin/soffice',
                '/Applications/LibreOffice.app/Contents/MacOS/soffice'
            ]
        
        for path in paths:
            try:
                if sys.platform == 'win32':
                    # Для Windows проверяем существование файла
                    if os.path.exists(path):
                        return path
                    # Также пробуем найти через where
                    try:
                        result = subprocess.run(['where', 'soffice'], 
                                              capture_output=True, text=True, shell=True)
                        if result.returncode == 0:
                            return result.stdout.strip().split('\n')[0]
                    except:
                        pass
                else:
                    # Для Linux/Mac проверяем через which
                    result = subprocess.run(['which', path], 
                                          capture_output=True, text=True)
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
            except:
                continue
        
        return None

    def get_output_extension(self, mode):
        """Получаем расширение файла на основе режима конвертации"""
        target_format = mode.split(' -> ')[1].strip()
        
        format_map = {
            'pdf': 'pdf',
            'docx': 'docx',
            'odt': 'odt',
            'txt': 'txt',
            'doc': 'txt'
        }
        return format_map.get(target_format, target_format)

    def convert_files(self):
        """Обработка файлов с сохранением последнего режима конвертации"""
        try:
            # Сохраняем текущий режим перед конвертацией
            self.save_settings()
            
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
            
            # Нормализуем расширение
            ext_map = {
                'docx': '.docx',
                'doc': '.doc',
                'pdf': '.pdf',
                'odt': '.odt',
                'rtf': '.rtf',
                'html': '.html'
            }
            
            search_ext = ext_map.get(input_ext, f'.{input_ext}')
            
            files = [f for f in os.listdir(source) 
                    if f.lower().endswith(search_ext) or 
                    (input_ext == 'pdf' and f.lower().endswith('.pdf'))]
            
            if not files:
                messagebox.showwarning("Предупреждение", 
                                     f"Нет файлов с расширением {search_ext} для конвертации в выбранной директории")
                return
            
            # Настраиваем прогрессбар
            self.progress['maximum'] = len(files)
            self.progress['value'] = 0
            self.current_file_label.config(text="")
            
            # Обрабатываем файлы
            success_count = 0
            error_files = []
            
            for i, filename in enumerate(files, 1):
                self.update_progress(i, filename)
                
                input_file = os.path.join(source, filename)
                output_ext = self.get_output_extension(mode)
                output_file = os.path.join(target, f"{os.path.splitext(filename)[0]}.{output_ext}")
                
                try:
                    print(f"\nОбработка {i}/{len(files)}: {filename}")
                    
                    # Выполняем конвертацию
                    if mode == "pdf -> txt":
                        if self.convert_pdf_to_txt(input_file, output_file):
                            success_count += 1
                            print(f"✓ Успешно: {filename} -> {os.path.basename(output_file)}")
                    elif mode == "pdf -> txt (OCR)":
                        if self.convert_pdf_to_txt_ocr(input_file, output_file):
                            success_count += 1
                            print(f"✓ Успешно (OCR): {filename} -> {os.path.basename(output_file)}")
                    elif mode in ("doc -> txt", "docx -> txt"):
                        if self.convert_doc_to_txt(input_file, output_file):
                            success_count += 1
                            print(f"✓ Успешно: {filename} -> {os.path.basename(output_file)}")
                    else:
                        output_format = self.get_libreoffice_format(mode)
                        if self.convert_with_libreoffice(input_file, output_file, output_format):
                            success_count += 1
                            print(f"✓ Успешно: {filename} -> {os.path.basename(output_file)}")
                            
                except Exception as e:
                    error_msg = str(e)
                    print(f"✗ Ошибка при обработке {filename}: {error_msg}")
                    error_files.append(f"{filename}: {error_msg[:100]}...")
            
            # Показываем результат
            result_msg = f"Конвертация завершена!\nУспешно: {success_count} из {len(files)}"
            if error_files:
                result_msg += f"\n\nОшибки ({len(error_files)}):\n" + "\n".join(error_files[:5])
                if len(error_files) > 5:
                    result_msg += f"\n... и еще {len(error_files) - 5} ошибок"
            
            if error_files:
                messagebox.showwarning("Результат конвертации", result_msg)
            else:
                messagebox.showinfo("Успех", result_msg)
            
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
        print(f"Критическая ошибка: {str(e)}")
        print("\nУбедитесь, что:")
        print("1. LibreOffice установлен и доступен в PATH")
        print("2. Для OCR функций установлены необходимые пакеты:")
        print("   pip install PyPDF2 pytesseract Pillow PyMuPDF")
        print("3. Tesseract OCR установлен (https://github.com/UB-Mannheim/tesseract/wiki)")
        print("   При установке Tesseract выберите языки: English, Russian")
        input("\nНажмите Enter для выхода...")