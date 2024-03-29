Эта папка содержит библиотеки в виде JAR-файлов. Модуль постобработки требует
для своей работы LibreOffice UNO Runtime Environment, который поставляется с
каждым выпуском LibreOffice.

Путь к классам LibreOffice задаётся переменной `LIBREOFFICE_CLASSES`. Путь к
основной директории LibreOffice - переменной `LIBREOFFICE_HOME`. Если эти
переменные не указаны, программа предпримет попытку найти классы LibreOffice в
следующих папках:

- На Windows: 
  - `%WINDIR%\..\Program Files\LibreOffice\program\classes`;
  - `%WINDIR%\..\Program Files (x86)\LibreOffice\program\classes`;
- На Linux:
  - `/usr/lib/libreoffice/program/classes`;
  - `/opt/libreoffice*/program/classes`;
- На macOS:
  - `/Applications/LibreOffice.app/Contents/Resources/java`.

Для других платформ автоматический поиск на данный момент не реализован.