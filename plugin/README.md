# Revit plugins

В каталоге `plugin` находятся два проекта под Revit 2022:

1. `Test.Hello` — тестовая кнопка "Привет".
2. `Test.ExportToExcel` — кнопка "Экспорт в Excel" для выгрузки экземпляров и параметров в `.xlsx`.

## Общие требования
Оба проекта на `net48` и ссылаются на:
- `C:\Program Files\Autodesk\Revit 2022\RevitAPI.dll`
- `C:\Program Files\Autodesk\Revit 2022\RevitAPIUI.dll`

## Быстрая сборка (Visual Studio)
1. Откройте нужный `.csproj` в Visual Studio 2022.
2. Выберите конфигурацию `Release`.
3. Выполните **Build > Build Solution**.
4. DLL будет в `bin\Release` проекта.

## Установка addin
1. Скопируйте `.addin` в `%ProgramData%\Autodesk\Revit\Addins\2022\`.
2. Внутри `.addin` укажите фактический путь к собранной DLL в теге `<Assembly>`.
