# Test.ExportToExcel (Revit 2022)

<<<<<<< codex/create-revit-add-in-for-excel-export-krhb6z
Плагин для Autodesk Revit 2022, который добавляет кнопки **«Экспорт в CSV»** и **«Экспорт в XLSX»** в Ribbon и выгружает экземпляры элементов модели в `.csv` и `.xlsx`.
=======
Плагин для Autodesk Revit 2022, который добавляет кнопку **«Экспорт в CSV»** в Ribbon и выгружает экземпляры элементов модели в `.csv`.
>>>>>>> main

## 1. Что делает плагин

- Собирает элементы через `FilteredElementCollector(...).WhereElementIsNotElementType()`.
- Формирует таблицу с колонками:
  - `id_имя семейства_имя типа`
  - `ElementId`
  - `Category`
  - `Family`
  - `Type`
  - далее все параметры (Built-in → Project → Shared, сортировка по имени).
- Значения параметров:
  - есть и заполнен → значение,
  - есть, но пустой → `""`,
  - отсутствует у элемента → `"no"`.

## 2. Формат выгрузки

<<<<<<< codex/create-revit-add-in-for-excel-export-krhb6z
- Формат 1: **CSV** (UTF-8 BOM, разделитель `;`)
- Формат 2: **XLSX**
=======
- Формат: **CSV**
- Кодировка: **UTF-8 BOM**
- Разделитель: `;` (удобно для русской локали Excel)
>>>>>>> main

## 3. Структура проекта

```text
/plugin/Test.ExportToExcel
  App.cs
  Test.ExportToExcel.csproj
  Test.ExportToExcel.addin
  README.md
  /Commands
    ExportToExcelCommand.cs
<<<<<<< codex/create-revit-add-in-for-excel-export-krhb6z
    ExportToXlsxCommand.cs
=======
>>>>>>> main
  /Infrastructure
    FileLogger.cs
  /Models
    ElementExportRow.cs
    ExportData.cs
    ParameterColumn.cs
  /Services
    ElementExportDataService.cs
    CsvExportService.cs
<<<<<<< codex/create-revit-add-in-for-excel-export-krhb6z
    XlsxExportService.cs
=======
>>>>>>> main
  /UI
    ExportProgressWindow.cs
  /Properties
    AssemblyInfo.cs
```

## 4. UX в Revit

<<<<<<< codex/create-revit-add-in-for-excel-export-krhb6z
1. Кнопки: вкладка `BOQ Tools` → панель `Export` → **«Экспорт в CSV»** и **«Экспорт в XLSX»**.
2. Открывается `SaveFileDialog` с нужным фильтром (`*.csv` или `*.xlsx`).
=======
1. Кнопка: вкладка `BOQ Tools` → панель `Export` → **«Экспорт в CSV»**.
2. Открывается `SaveFileDialog` с фильтром `*.csv`.
>>>>>>> main
3. Показывается прогресс-окно с кнопкой `Cancel`.

## 5. Логирование

Лог пишется в:

```text
%AppData%\Test.ExportToExcel\export.log
```

## 6. Сборка

1. Откройте `plugin/Test.ExportToExcel/Test.ExportToExcel.csproj` в Visual Studio 2022.
2. Соберите `Release`.

CLI:

```powershell
msbuild .\plugin\Test.ExportToExcel\Test.ExportToExcel.csproj /t:Restore
msbuild .\plugin\Test.ExportToExcel\Test.ExportToExcel.csproj /p:Configuration=Release
```

## 7. Подключение в Revit

1. Скопируйте `Test.ExportToExcel.dll` в папку плагина.
2. В `Test.ExportToExcel.addin` укажите корректный `<Assembly>` путь.
3. Положите `.addin` в:

```text
%ProgramData%\Autodesk\Revit\Addins\2022\
```

или:

```text
%AppData%\Autodesk\Revit\Addins\2022\
```
