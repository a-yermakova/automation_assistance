## Automation Assistance 
Проект, выполненный во время стажировки в УК "Арсагера".
Приложение помогает в автоматизации процесса заполнения аналитических моделей информацией из отчетов эмитентов.

### Структура проекта:
- [automation_assistance.py](https://github.com/a-yermakova/automation_assistance/blob/main/automation_assistance.py "automation_assistance.py") - функционал приложения. Обработка XLSX файлов происходит с помощью библиотеки [openpyxl](https://openpyxl.readthedocs.io/en/stable/), работать с файлами формата INI помогает библиотека [configparser](https://docs.python.org/3/library/configparser.html/). 

- [automation_web_gui.py](https://github.com/a-yermakova/automation_assistance/blob/main/automation_web_gui.py "automation_web_gui.py") - веб интерфейс приложения, реализованный с помощью библиотеки [streamlit](https://docs.streamlit.io/).

- [automation_assistance_exceptions.py](https://github.com/a-yermakova/automation_assistance/blob/main/automation_assistance_exceptions.py "automation_assistance_exceptions.py") - модуль пользовательских исключений.

- [balance_config.ini](https://github.com/a-yermakova/automation_assistance/blob/main/balance_config.ini "balance_config.ini"), [cashflow_config.ini](https://github.com/a-yermakova/automation_assistance/blob/main/cashflow_config.ini "cashflow_config.ini"), [income_config.ini](https://github.com/a-yermakova/automation_assistance/blob/main/income_config.ini "income_config.ini"), 
[segments_config.ini](https://github.com/a-yermakova/automation_assistance/blob/main/segments_config.ini "segments_config.ini") - сокращенная версия конфигурационных файлов, приравнивает одинаковые по смыслу названия статей.

#### Просмотр приложения доступен по ссылке:

https://a-yermakova-automation-assistance-automation-web-gui-m86l5j.streamlitapp.com/