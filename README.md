# arch_rename_export
Проект переименовки выгрузки PDF и DWG файлов из Revit

### Установка и использование
#### Клонируем репозиторий

#### Устанавливаем виртуальное окружение 
```
python -m venv venv
```
#### Запускаем виртуальное окружение
```
venv\Scripts\activate.bat
```
#### Устанавливаем библиотеки
```
pip install -r requirements.txt
```
#### Выгрузка в исполняемый файл
```
pyinstaller --onefile '.\Запустить при выгрузке с АР dwg--pdf.py'
```

### Работа проекта
1. Необходимо сохранить текущуей папке директории скопировать dist/Запустить при выгрузке с АР dwg--pdf.exe и template.xlsx
2. Создать 2 папки PDF и DWG
3. Произвести вырузку из реавита в папку DWG
4. Файл заполняет в файле template.xlsx на листе "files" список всех файлов из директории DWG
5. После чего пользователь закрывает приложение и открывает template.xlsx на листе "filenames" формирует в столбце "A" правильное имя файлов а в B, C и остальные варианты названий выгруженных файлов, которые надо переимновать.
6. После чего снова запустить исполняемый файл Запустить при выгрузке с АР dwg--pdf.exe
