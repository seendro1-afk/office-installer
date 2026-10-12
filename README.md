# Office Installer (MVP)

Русскоязычный сайт и исходный код Windows-приложения для установки **Office LTSC 2024** с выбором Word, Excel, PowerPoint, Outlook, OneNote и Access.

> Важно: Office LTSC 2024 — корпоративная редакция, требующая соответствующих прав и действующей лицензии. Программа не выполняет активацию и не включает обход лицензирования.

## Запуск Windows-приложения

1. Установите .NET 8 SDK на **Windows 10/11**.
2. Скачайте официальный **Office Deployment Tool** от Microsoft и распакуйте его, чтобы получить `setup.exe`.
3. В каталоге `windows` выполните `dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o publish`.
4. Запустите `publish/OfficeInstaller.exe` **от имени администратора**.
5. Отметьте приложения, укажите `setup.exe` из ODT и нажмите «Установить выбранное».

Приложение формирует XML с `ExcludeApp`, после чего вызывает официальный ODT. Оно не удаляет старые версии Office автоматически, не собирает ключи и не обещает совместимость с любыми уже установленными редакциями. Проверяйте установку на тестовой машине.

## Сайт

Файл `index.html` можно разместить через GitHub Pages: Settings → Pages → Deploy from a branch → main / root. До публикации собранного EXE в GitHub Releases кнопка Windows ведёт на пустую страницу релизов.

## Публикация EXE

Соберите проект на Windows, проверьте антивирусом, по возможности подпишите сертификатом издателя и загрузите EXE в Releases. Не публикуйте ключи активации в репозитории.

## Источники

- https://learn.microsoft.com/office/ltsc/2024/deploy
- https://learn.microsoft.com/microsoft-365-apps/deploy/office-deployment-tool-configuration-options
