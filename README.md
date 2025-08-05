
## Фото-сортировцик с созданием поддиректорий с датой создания, переворотом и обратным геокодированием на основе EXIF данных

``` bash
git clone https://github.com/SergeyMi37/photo-sorter
cd photo-sorter
```

Create virtual environment (optional)
``` bash
python3 -m venv env
source env/bin/activate
```

Create virtual environment for Windows
``` bash
python -m venv env
source env/Scripts/activate
```

Install all requirements:
``` bash
pip install -r requirements.txt
```

Укажите исходную и целевую директорию
``` bash
python selector-dirs-tk.py 
```

### Во время написания программы не один ИИ не пострадал, но потрепан был изрядно GigaChat (@gigachat_bot)
Его собрат GigaCode смотрелся не очень, но старался помогать в VsCode)


 Для сборки исполняемого файла (*.exe) вашей программы на Tkinter для Windows используется инструмент **PyInstaller**. Вот пошаговая инструкция по сборке релиза приложения:

### Шаги по созданию релиза

#### 1. Установка PyInstaller
Сначала убедитесь, что у вас установлен PyInstaller. Если нет, установите его командой:
```bash
pip install pyinstaller
```
[[4](https://habr.com/ru/sandbox/182102/)]

#### 3. Сборка с помощью PyInstaller
Откройте терминал и перейдите в папку вашего проекта. Затем выполните команду:
```bash
pyinstaller --onefile --windowed selector-dirs-tk.py
```
Параметры команды:
- `--onefile`: Создает один файл `.exe`, содержащий всю необходимую библиотеку.
- `--windowed`: Запускает ваше приложение без консольного окна.

Это создаст два каталога: `build` и `dist`. В каталоге `dist` появится ваш исполняемый файл.

### Дополнительная настройка сборок

#### Настройка иконки приложения
Вы можете добавить свою иконку приложению, используя следующий формат:
```bash
pyinstaller --onefile --windowed --icon путь_до_иконки.ico имя_скрипта.py
```
Иконка должна иметь расширение `.ico`.

#### Добавление зависимых файлов
Если ваша программа зависит от внешних ресурсов (например, изображений или конфигурационных файлов), добавьте их следующим образом:
```bash
pyinstaller --onefile --windowed --add-data "путь_до_файла;название_каталога" имя_скрипта.py
```
Например, если у вас есть картинка `"logo.png"` в корне проекта, команда будет выглядеть так:
```bash
pyinstaller --onefile --windowed --add-data "logo.png;." имя_скрипта.py
```
Файл скопируется в тот же каталог, где находится исполняемый файл.

### Создание инсталлятора
Чтобы упростить установку пользователям, можно создать полноценный инсталлятор. Для этого удобно воспользоваться программой **InstallForge**:

1. Скачайте и установите InstallForge.
2. Откройте проект и настройте общую информацию о приложении (название, версия, компания и т.п.).
3. Перейдите в раздел "Files" и добавьте файлы из каталога `dist`, включая сам исполняемый файл и дополнительные ресурсы.
4. Установите необходимые настройки для установки (например, создание ярлыков на рабочем столе и в меню "Пуск").
5. Соберите готовый инсталляционный пакет, нажав соответствующую кнопку.

Таким образом, у вас получится полноценное приложение с удобным процессом установки.

Следуя этим шагам, вы сможете легко собирать релизы ваших приложений на Tkinter для Windows.\[[1](https://www.pythonguis.com/tutorials/packaging-tkinter-applications-windows-pyinstaller/)]\[[2](https://www.pythonguis.com/tutorials/packaging-tkinter-applications-windows-pyinstaller/)]

*Для ответа использовал актуальные интернет-источники:*

 1. [www.pythonguis.com: Packaging Tkinter applications for Windows, with PyInstaller...](https://www.pythonguis.com/tutorials/packaging-tkinter-applications-windows-pyinstaller/)
 2. [www.pythonguis.com: Packaging Tkinter applications for Windows, with PyInstaller...](https://www.pythonguis.com/tutorials/packaging-tkinter-applications-windows-pyinstaller/)
 3. [coderslegacy.com: Using Tkinter with Python Pyinstaller to create Exe - CodersLegacy](https://coderslegacy.com/python/tkinter-pyinstaller/)
 4. [habr.com: Tkinter: кратко для начинающих / Песочница / Хабр](https://habr.com/ru/sandbox/182102/)
 5. [medium.com: Building Cross-Platform Applications with Tkinter and... | Medium](https://medium.com/tomtalkspython/building-cross-platform-applications-with-tkinter-and-pyinstaller-d7a10163c550)