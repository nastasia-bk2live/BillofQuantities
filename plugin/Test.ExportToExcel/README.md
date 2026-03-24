# Revit plugin: Test.ExportToExcel

Отдельная сборка и add-in для Revit 2022.

## Что делает кнопка
Кнопка **«Экспорт в Excel»**:
- Открывает `SaveFileDialog` для выбора пути `.xlsx`.
- Выгружает **все экземпляры** (`WhereElementIsNotElementType`) в Excel.
- Формирует первый столбец как `id_имя семейства_имя типа`.
- Добавляет служебные столбцы: `ElementId`, `Category`, `Family`, `Type`.
- Далее добавляет параметры:
  1. сначала встроенные (`BuiltInParameter`),
  2. затем пользовательские параметры проекта/общие параметры.
- Внутри каждой группы параметры сортируются по имени (A-Z).
- Если параметра у элемента нет — пишется `no`.
- Если параметр есть, но пустой — пишется пустая строка.
- Для `ElementId` параметров выводится числовой `Id`.

## Файлы
- `App.cs` — регистрация кнопки в Ribbon.
- `ExportToExcelCommand.cs` — логика экспорта.
- `Test.ExportToExcel.csproj` — проект net48.
- `Test.ExportToExcel.addin` — манифест для Revit.

## Сборка
Откройте `Test.ExportToExcel.csproj` в Visual Studio 2022 и соберите в `Release`.
