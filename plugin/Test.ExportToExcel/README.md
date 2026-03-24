# Test.ExportToExcel (Revit 2022)

Плагин для Autodesk Revit 2022, который добавляет кнопку **«Экспорт в Excel»** в Ribbon и выгружает **максимально все экземпляры элементов модели** в `.xlsx`-файл.

## 1. Назначение

Плагин собирает данные по экземплярам элементов (`WhereElementIsNotElementType`) и экспортирует:

- служебные поля:
  - `id_имя семейства_имя типа`
  - `ElementId`
  - `Category`
  - `Family`
  - `Type`
- параметры в порядке:
  1. Built-in
  2. Project
  3. Shared
- сортировка параметров внутри каждой группы: по имени, по алфавиту.

## 2. Структура проекта

```text
/plugin/Test.ExportToExcel
  App.cs
  Test.ExportToExcel.csproj
  Test.ExportToExcel.addin
  packages.config
  README.md
  /Commands
    ExportToExcelCommand.cs
  /Infrastructure
    FileLogger.cs
  /Models
    ElementExportRow.cs
    ExportData.cs
    ParameterColumn.cs
  /Services
    ElementExportDataService.cs
    ExcelExportService.cs
  /UI
    ExportProgressWindow.cs
  /Properties
    AssemblyInfo.cs
```

## 3. Технологии

- .NET Framework 4.8
- C# 7.3
- Revit API 2022 (`RevitAPI.dll`, `RevitAPIUI.dll`)
- ClosedXML (экспорт XLSX)
- WPF (окно прогресса с Cancel)

## 4. Как работает кнопка

1. Пользователь нажимает кнопку **«Экспорт в Excel»** (вкладка `BOQ Tools`, панель `Export`).
2. Открывается `SaveFileDialog` с фильтром `*.xlsx`.
3. Запускается окно прогресса (WPF):
   - этап 1: сбор данных
   - этап 2: запись Excel
   - кнопка `Cancel` прерывает процесс.
4. Формируется Excel с единым набором колонок параметров (объединение параметров по всем элементам).
5. Значения:
   - параметр есть и заполнен → значение
   - параметр есть, но пустой → `""`
   - параметра нет у конкретного элемента → `"no"`
   - для `ElementId`-параметров выгружается числовой ID
   - для числовых/длиновых/площадных значений приоритет `AsValueString()`.

## 5. Логирование

Лог пишется в файл:

```text
%AppData%\Test.ExportToExcel\export.log
```

В лог попадают:

- старт/завершение экспорта
- отмена пользователем
- ошибки и stack trace

## 6. Сборка проекта

### Вариант A: Visual Studio 2022

1. Откройте `Test.ExportToExcel.csproj`.
2. Выполните restore NuGet пакетов (`ClosedXML` и зависимости).
3. Убедитесь, что доступны ссылки на:
   - `C:\Program Files\Autodesk\Revit 2022\RevitAPI.dll`
   - `C:\Program Files\Autodesk\Revit 2022\RevitAPIUI.dll`
4. Соберите проект в `Release`.

### Вариант B: MSBuild

```powershell
nuget restore .\Test.ExportToExcel.csproj
msbuild .\Test.ExportToExcel.csproj /p:Configuration=Release
```

## 7. Подключение в Revit

1. Скопируйте `Test.ExportToExcel.dll` в папку, где будет храниться плагин (например `C:\RevitPlugins\Test.ExportToExcel\`).
2. Откройте файл `Test.ExportToExcel.addin`.
3. Замените значения `<Assembly>` на полный путь к DLL.
4. Скопируйте `.addin` в одну из папок Revit 2022:

```text
%ProgramData%\Autodesk\Revit\Addins\2022\
```

или для текущего пользователя:

```text
%AppData%\Autodesk\Revit\Addins\2022\
```

5. Запустите Revit 2022.
6. В Ribbon появится вкладка `BOQ Tools`, панель `Export`, кнопка **«Экспорт в Excel»**.

## 8. Важные замечания

- В `.addin` содержатся 2 записи в одном файле:
  - `Type="Application"` (регистрация `IExternalApplication`)
  - `Type="Command"` (регистрация `IExternalCommand`)
- Проект рассчитан на Revit 2022 и .NET Framework 4.8.
- Если в модели очень много элементов, экспорт может занять заметное время.
