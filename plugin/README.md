# Revit plugin: Test.Hello

Проект `Test.Hello` расположен в `plugin/Test.Hello` и предназначен для Revit 2022.

## Что делает
- При запуске Revit регистрируется вкладка **Привет**.
- Создаётся панель **Привет**.
- Добавляется кнопка **Привет**.
- По нажатию кнопки выполняется `TaskDialog.Show("Плагин", "Привет!")`.

## Сборка
Проект на `net48` с ссылками на:
- `C:\Program Files\Autodesk\Revit 2022\RevitAPI.dll`
- `C:\Program Files\Autodesk\Revit 2022\RevitAPIUI.dll`

## Установка
1. Соберите `Test.Hello.dll`.
2. Скопируйте `Test.Hello.addin` в:
   - `%ProgramData%\Autodesk\Revit\Addins\2022\`
3. В `Test.Hello.addin` замените путь в `<Assembly>` на фактический путь к `Test.Hello.dll`.
