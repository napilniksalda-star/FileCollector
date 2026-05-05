# FileCollector V3.0 «Phoenix»

WinForms-приложение для Windows: читает список имён файлов из Excel, рекурсивно ищет их в указанных папках и копирует в папку сбора с отчётом о результатах.

## Возможности

- Чтение списка имён из столбца B первого листа Excel (.xlsx / .xls)
- Автоматическое определение строк-заголовков и служебных значений
- Очистка спецсимволов из ячеек (неразрывные пробелы, smart-кавычки, тире и т. п.)
- Поиск через построенный за один проход индекс файлов (O(1) на запрос)
- Преднабор расширений (CAD/SolidWorks/Inventor/CATIA/Creo) + ручной ввод
- Предпросмотр найденных файлов с возможностью снять/выбрать вручную
- Асинхронное копирование с retry для заблокированных файлов и защитой от path-traversal
- Пакетный логгер (UI не блокируется при тысячах сообщений)
- Отчёт о ненайденных файлах в папке сбора
- Экспорт результатов в HTML (через built-in plugin)
- Drag-and-drop, клавиатурные сокращения (Ctrl+O / Ctrl+P / Ctrl+S / Esc)

## Требования

- Windows 10/11 x64
- .NET SDK 10.0 — только для сборки из исходников

## Сборка

```powershell
dotnet restore
dotnet build -c Release
```

Исполняемый файл: `bin/Release/net10.0-windows/win-x64/FileCollector.exe`

## Single-file публикация

```powershell
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

Результат: `bin/Release/net10.0-windows/win-x64/publish/FileCollector.exe`

## Быстрый старт

1. Подготовьте Excel с именами файлов в столбце **B** первого листа.
2. В приложении выберите Excel.
3. Добавьте одну или несколько папок для поиска (можно перетащить).
4. Укажите папку для сбора результатов.
5. Отметьте нужные расширения (по умолчанию `.pdf` и `.dxf`).
6. Нажмите «Предпросмотр» (Ctrl+P), снимите галочки с лишних строк, затем «Запустить копирование» (Ctrl+S).

## Структура проекта

```
FileCollector/
├── Program.cs                  # точка входа, EPPlus license + Application.Run
├── MainForm.cs                 # контроллер: состояние, обработчики, оркестрация
├── MainForm.Layout.cs          # построение UI (partial-class)
├── Core/                       # бизнес-логика, без зависимостей от WinForms
│   ├── Models.cs               # PreviewItem, CopyOperation, CopyStatus, RunSummary
│   ├── VersionInfo.cs          # централизованные константы версии
│   ├── Excel/                  # чтение Excel, очистка ячеек, детект заголовков
│   ├── Search/                 # FileIndex - однопроходный индекс файлов
│   ├── Copy/                   # FileCopyService + PathSafety
│   ├── Logging/                # IAppLogger, BatchUiLogger
│   └── Reporting/              # NotFoundReportWriter
├── Plugins/                    # подключаемые модули
│   ├── IPlugin.cs              # контракты IPlugin / IFileProcessorPlugin / IExportPlugin
│   ├── PluginManager.cs        # загрузка из DLL + Register для built-in
│   └── BuiltIn/                # HtmlExportPlugin, DuplicateDetectorPlugin, PdfAnalyzerPlugin
├── Resources/                  # иконка приложения, лого
├── tools/                      # вспомогательные скрипты (generate-icon.ps1)
├── app.manifest
├── FileCollector.csproj
└── FileCollector.sln
```

## Архитектура V3.0 (что изменилось vs V2.1)

- Монолит `MainForm.cs` (2145 строк) разделён на тонкий контроллер + UI-конструктор + слой `Core/`.
- `FileIndex` строится один раз и отвечает на запросы за O(1) — было `Parallel.ForEach` рекурсивно по дереву на каждую пару (имя, расширение).
- `FileCopyService` копирует через async-stream вместо `File.Copy` + `Thread.Sleep`; учитывает `CancellationToken`.
- `PathSafety.SafeCombine` валидирует имя назначения и блокирует path-traversal.
- `BatchUiLogger` группирует сообщения и обновляет UI таймером — больше нет `Control.Invoke` на каждую строку лога.
- `CopyOperation.Status` стал enum'ом `CopyStatus` (раньше — свободная строка).
- Система плагинов подключена: HTML-экспорт работает через `HtmlExportPlugin`.

## Лицензирование

EPPlus используется в режиме `NonCommercial`. При коммерческом использовании оформите соответствующую лицензию EPPlus.
