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
- ClosedXML (основной экспорт XLSX)
- NPOI (fallback при ошибке IsolatedStorage в ClosedXML)
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

- При превышении лимитов Excel данные автоматически разбиваются на несколько листов:
  - по строкам (лимит 1 048 576),
  - по столбцам параметров (лимит 16 384).
- Значения перед записью санитизируются (удаляются недопустимые XML-символы),
  а слишком длинные значения обрезаются до 32 767 символов (лимит ячейки Excel).

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
2. Выполните restore NuGet пакетов (`ClosedXML` и зависимости) через PackageReference.
3. Убедитесь, что доступны ссылки на:
   - `C:\Program Files\Autodesk\Revit 2022\RevitAPI.dll`
   - `C:\Program Files\Autodesk\Revit 2022\RevitAPIUI.dll`
4. Соберите проект в `Release`.

### Вариант B: MSBuild

```powershell
cd <корень_репозитория>
msbuild .\plugin\Test.ExportToExcel\Test.ExportToExcel.csproj /t:Restore
msbuild .\plugin\Test.ExportToExcel\Test.ExportToExcel.csproj /p:Configuration=Release
```

> Проект использует `PackageReference`, поэтому отдельный `packages.config` не требуется.


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


## 9. Диагностика ошибки `IsolatedStorageException`

Если в логе есть ошибка вида:

- `Unable to determine the identity of domain`
- stack trace внутри `MS.Internal.IO.Packaging` / `ClosedXML`

это обычно проявляется на очень больших выгрузках в окружении Revit/.NET Framework.

Что уже сделано в коде:

- основной путь: экспорт через ClosedXML;
- fallback: при `IsolatedStorageException` автоматическое сохранение через NPOI в тот же `.xlsx`.

Практические рекомендации:

1. Выгружайте сначала на локальный диск (например `C:\Temp\...xlsx`), а потом переносите на сетевой диск.
2. Убедитесь, что в пути нет блокировок по правам и антивирусом.
3. Если экспорт очень большой, попробуйте временно закрыть лишние приложения для увеличения доступной памяти процесса Revit.

Также если Excel открывает файл с recoveryLog и пишет про повреждение `sheet1.xml`,
обычная причина — недопустимые XML-символы в текстовых параметрах или превышение лимитов Excel.
Текущая версия экспорта это учитывает: выполняет санитизацию значений и разбивку на листы.
