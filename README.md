# Photos → Word / Фото → Word

[English](#english) · [Русский](#русский)

## English

A small utility for anyone who sends screenshots to an AI: it collects images into a single Word document and numbers them in order.

### How to use

1. Take a screenshot with `PrtScn`, or press `Win+Shift+S` and copy the selected area.
2. Open **`Запустить.bat`**. On first launch, choose the interface language: **English (USA)**, **Russian** or **Chinese (中文)**.
3. Press **Ctrl+V** — the screenshot appears in the window. Repeat for every image you need.
4. If necessary, change the order with **Earlier / Later**, or delete unwanted images.
5. Click **Save to Word** and choose where to save the `.docx` file.

The document will look like this:

```
1.
[photo 1]
2.
[photo 2]
...
```

Images are scaled to the A4 page width and are never enlarged beyond their original size.

### Hotkeys

| Keys | Action |
|---|---|
| `Ctrl+V` | paste a screenshot from the clipboard |
| `Delete` | delete the selected photo |
| `Esc` | clear the selection |

### More

- **Add files** lets you add existing images from disk.
- **Clear** removes all photos at once.
- The **Dark / Light** button in the top-right corner switches the appearance. The selected theme is remembered.
- The globe button with the language code (**EN / RU / 中文**) changes the interface language at any time. The choice is saved together with the theme.
- If `PrtScn` does not put an image on the clipboard, check **Windows Settings → Accessibility → Keyboard → “Use the PrtScn button to open screen snipping”**. You can use the snipping overlay, or press `Win+Shift+S` and copy the selected area.
- Settings are stored separately for each Windows user in `%APPDATA%\PhotoToWord\settings.json`. They are not embedded in the EXE. Delete that folder to reset the language and theme.
- `PhotoToWord.exe` is standalone: Python is not required, and the translations are already embedded. No `locales` folder is created or needed next to the EXE. The first launch may take a few seconds. The EXE is unsigned, so Windows SmartScreen may show a warning — click **More info → Run anyway**.
- The `Pillow`, `python-docx` and `customtkinter` dependencies are installed automatically by `Запустить.bat`. To run from source manually:

```
python -m pip install -r requirements.txt
python photo_to_word.py
```

### Files

- `PhotoToWord.exe` — ready-to-run Windows app (Python not required, translations embedded)
- `photo_to_word.py` — source code
- `locales/` — translation files for running from source (`en.json`, `ru.json`, `zh.json`)
- `Запустить.bat` — launcher: starts the `.exe` first, otherwise runs the Python version
- `requirements.txt` — libraries needed to run from source

---

## Русский

Мини-программа для тех, кто отправляет скриншоты в ИИ: собирает изображения в один документ Word и нумерует их по порядку.

### Как пользоваться

1. Сделайте скриншот клавишей `PrtScn` или нажмите `Win+Shift+S` и скопируйте выбранную область.
2. Откройте **`Запустить.bat`**. При первом запуске выберите язык интерфейса: **English (USA)**, **Russian** или **Chinese (中文)**.
3. Нажмите **Ctrl+V** — скриншот появится в окне. Повторите для всех нужных изображений.
4. При необходимости поменяйте порядок кнопками **«Раньше» / «Позже»** или удалите лишнее.
5. Нажмите **«Сохранить в Word»** и выберите, куда сохранить файл `.docx`.

В документе получится:

```
1.
[фото 1]
2.
[фото 2]
...
```

Фото автоматически вписываются в ширину страницы A4 и не растягиваются больше исходного размера.

### Горячие клавиши

| Клавиши | Действие |
|---|---|
| `Ctrl+V` | вставить скриншот из буфера |
| `Delete` | удалить выбранное фото |
| `Esc` | снять выделение |

### Дополнительно

- **«Добавить файлы»** — можно добавить готовые изображения с диска.
- **«Очистить»** — убрать все фото сразу.
- Кнопка **«Тёмная» / «Светлая»** в правом верхнем углу переключает оформление. Выбранная тема запоминается.
- Кнопка с глобусом и кодом языка (**EN / RU / 中文**) меняет язык интерфейса в любой момент. Выбор сохраняется вместе с темой.
- Если `PrtScn` не копирует изображение в буфер, проверьте **Параметры Windows → Специальные возможности → Клавиатура → «Использовать клавишу PrtScn для открытия фрагмента и эскиза»**. Можно выбрать область в открывшемся оверлее или нажать `Win+Shift+S` и скопировать выделенную область.
- Настройки хранятся отдельно для каждого пользователя Windows в `%APPDATA%\PhotoToWord\settings.json`. Они не встроены в EXE. Чтобы сбросить язык и тему, удалите эту папку.
- `PhotoToWord.exe` — самостоятельная программа: Python не требуется, переводы уже встроены. Папка `locales` рядом с EXE не создаётся и не нужна. Первый запуск может занять несколько секунд. EXE не подписан, поэтому Windows SmartScreen может показать предупреждение — нажмите **«Подробнее» → «Выполнить в любом случае»**.
- Зависимости `Pillow`, `python-docx` и `customtkinter` устанавливаются автоматически через `Запустить.bat`. Чтобы запустить исходник вручную:

```
python -m pip install -r requirements.txt
python photo_to_word.py
```

### Файлы

- `PhotoToWord.exe` — готовая программа для Windows (Python не требуется, переводы встроены)
- `photo_to_word.py` — исходный код
- `locales/` — файлы переводов для запуска из исходников (`en.json`, `ru.json`, `zh.json`)
- `Запустить.bat` — запуск: сначала пробует `.exe`, иначе запускает Python-версию
- `requirements.txt` — список библиотек для запуска из исходников
