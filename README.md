# gen_contracts
Generate docx/pdf files from docx template and json/csv datas by python. 

1)	Creation of templates for the formation of documents.

Templates are created in docx format in MS-Word editor. The data is uploaded in json format to files.
To display data fields, their names are framed with two curly braces: {{DocN}}
You can also use control structures:

1)	Conditions
{% if isTel==’X’ %}
     <text1>
{% else %}
     <text2>
{% endif %}
2)	Cycles
{%tr for i in inet_svcs %}
  < output of arbitrary text and fields of the inet_svcs array. For example, to display the tp field from the inet_svcs array, you need to specify {{i.tp}} >
{%tr endfor %}

You can also use csv format, but it will not be possible to use tables with a set of rows inside the document (one-to-many relationship)

2)	Launching the program for generating contracts in docx and pdf.
To generate documents based on the created templates and uploaded data, you need to run the gen_contracts.exe program with a set of parameters (if you run it without parameters, you will be prompted in a dialog to confirm the default value of each parameter or replace it with a new one):
parameter	Default value 	description
--data_path	data	Path to data files
--data_mask 	*.json	Data file mask (*.json or *.csv)
--templ_path 	templates	Path to docx templates
--output_path 	output	Folder for output docx and pdf files and execution log file
--no-pdf 		--no-pdf - don't create pdf (docx only) --pdf - create pdf
--no-replace 		--no-replace - do not recreate previously created files
--log_file 	output.log	--replace - recreate files
--encode	ansi	Log file name
--csv_delimiter	;	Delimiter in csv file
		
 
The program works according to the following algorithm
1) Open json/csv files with data one by one
2) Parses json/csv fields,
3) for each line creates a docx file with the specified name in the _filename field according to the template from the _docx_template field
4) creates a pdf from the created docx (if the --pdf parameter is specified)

Run example:
>gen_contracts.exe --data_path data --data_mask *.json --templ_path templates --output_path output --no-pdf --no-replace --log_file output.log --encode ansi --csv_delimiter ;

Log example:
2022-04-02 22:18:02,755 - INFO - Open data file: data\test1.csv
2022-04-02 22:18:02,833 - INFO - File generated: "output/100005".docx
2022-04-02 22:18:02,912 - INFO - File generated: "output/100006".docx
2022-04-02 22:19:03,112 - INFO - Open data file: data\test.json
2022-04-02 22:19:03,206 - INFO - File generated: "output/100001".docx
2022-04-02 22:19:03,284 - INFO - File generated: "output/100002".docx
2022-04-02 22:19:

---------------------------------------------------------------------------------------------------------------
--RUSSIAN
-------------------------------------------------------------------------------------------------------------
Генератор docx/pdf файлов из шаблонов docx и данных в json или csv форматах на питоне.
1)	Создание шаблонов для формирования документов.

Шаблоны создаются в формате docx в редакторе MS-Word. Данные выгружаются в формате json в файлы. 
Для вывода полей данных их наименования обрамляются двумя фигурными скобками: {{DocN}}
Также,  можно использовать управляющие конструкции:
1)	Условия
{% if isTel==’X’ %}
     <текст1>
{% else %}
     <текст2>
{% endif %}
2)	Циклы
{%tr for i in inet_svcs %}
  <вывод произвольного текста и полей массива inet_svcs. Например, для вывода поля tp из массива inet_svcs нужно указать {{i.tp}}>
{%tr endfor %}

Можно использовать и csv формат, но при этом не будет возможности использовать таблицы с набором строк внутри документа (связь один ко многим)


2)	Запуск программы формирования договоров в docx и pdf.
Для генерации документов по созданным шаблонам и выгруженным данным нужно запустить программу gen_contracts.exe с набором параметров (если запускать без параметров, то в диалоге будет предложено подтвердить значение по умолчанию каждого параметра или заменить новым):
параметр	Значение по умолчанию	описание
--data_path	data	Путь к файлам с данными
--data_mask 	*.json	Маска файлов с данными (*.json или *.csv)
--templ_path 	templates	Путь к шаблонам docx
--output_path 	output	Папка для выходных файлов docx и pdf и файла журнала выполнения
--no-pdf 		--no-pdf – не создавать pdf (только docx)
--pdf – создавать pdf
--no-replace 		--no-replace – не пересоздавать уже созданные ранее файлы 
--replace – пересоздавать файлы
--log_file 	output.log	Имя файла журнала
--encode	ansi	Кодировка в файле с данными (ansi,utf-8,..)
--csv_delimiter	;	Разделитель полей в csv-файле
		
 
Программа работает по следующему алгоритму
1)	Открывает файлы json/csv с данным по очереди
2)	Разбирает поля json/csv, 
3)	для каждой строки создает файл docx с указанным именем в поле _filename по шаблону из поля _docx_template
4)	из созданного docx создает pdf (если задан параметр --pdf)

Пример запуска:
>gen_contracts.exe --data_path data --data_mask *.json --templ_path templates --output_path output --no-pdf --no-replace --log_file output.log --encode ansi --csv_delimiter ;

Пример лога:
2022-04-02 22:18:02,755 - INFO - Open data file: data\test1.csv
2022-04-02 22:18:02,833 - INFO - File generated: "output/100005".docx
2022-04-02 22:18:02,912 - INFO - File generated: "output/100006".docx
2022-04-02 22:19:03,112 - INFO - Open data file: data\test.json
2022-04-02 22:19:03,206 - INFO - File generated: "output/100001".docx
2022-04-02 22:19:03,284 - INFO - File generated: "output/100002".docx
2022-04-02 22:19:03,362 - INFO - File generated: "output/100003".docx

