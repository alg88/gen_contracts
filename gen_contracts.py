# -*- coding: utf-8 -*-
'''
Created : 2022-02-17
@author: AlexG

pre-install:
1) https://github.com/elapouya/python-docx-template
pip install docxtpl
Примеры смотреть в \contracts\templates\tests\

2)docx2pdf
pip install docx2pdf

3)pyinstaller - для создания exe-файла
pip install pyinstaller
# Затем перейти в папку с Вашим файлом .py в командной строке (при помощи команды cd)
# Запустить команду pyinstaller не забудьте указать имя вашего скрипта
pyinstaller --onefile <your_script_name>.py
# Всё - у вас в папке появится папка src и там будет .exe файл.

4)pip install click
'''

from docxtpl import DocxTemplate
from docx2pdf import convert
import json  # Подключили библиотеку
import csv
import glob
import click
import os.path
import logging
import traceback

@click.option('--data_path', '-d', default='data', prompt='Folder for data files: ')
@click.option('--data_mask', '-m', default='*.json', prompt='Data file mask: ')
@click.option('--templ_path', '-t', default='templates', prompt='Folder for docs templates: ')
@click.option('--output_path', '-o', default='output', prompt='Folder for output files: ')
@click.option('--pdf/--no-pdf', default=False, prompt='Whether to generate pdf: ')
@click.option('--replace/--no-replace', default=False, prompt='Replace existing files: ')
@click.option('--log_file', '-o', default='output.log', prompt='Log file: ')
@click.option('--encode', '-e', default='ansi', prompt='encode(ansi,utf-8,..): ')
@click.option('--csv_delimiter', '-d', default=';', prompt='Delimiter in csv file: ')
@click.command()
def gen_contract(data_path, data_mask, templ_path, output_path, pdf, replace, log_file, encode,csv_delimiter):
    try:
#    logging.basicConfig(filename=output_path+'/'+log_file, level=logging.INFO)
        logger = logging.getLogger('gen_contracts')
        logger.setLevel(logging.INFO)
        fh = logging.FileHandler(filename=output_path+'/'+log_file)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        logger.addHandler(fh)

        data_files = glob.glob(data_path + '/' + data_mask)

        for df in data_files:
            log_txt = 'Open data file: '+df
            print(log_txt)
            logger.info(log_txt)
            if data_mask[-4:]=='json':
                with open(df, 'r', encoding=encode) as f:  # открыли файл
                    data = json.load(f)  # загнали все из файла в переменную
            elif data_mask[-3:]=='csv':
                with open(df, 'r', encoding=encode) as csvf:
                    csvReader = csv.DictReader(csvf,delimiter = csv_delimiter)
#                    dm=[]
#                    dm['contracts']= [ row for row in csvReader ]
                    dm = [ row for row in csvReader ]
                    data_str = json.dumps(dm,ensure_ascii=False, indent = 4)
                    data = json.loads(data_str)
#                    print(data)

#        print(data)  # вывели результат на экран

            l_templ_file=''
            tpl=''
#            for d in data['contracts']:  # создали цикл, который будет работать построчно
            for d in data:  # создали цикл, который будет работать построчно
                l_file_name = output_path+'/'+d['_filename']
                is_file_created = False
                if os.path.exists(l_file_name+ '.docx')==False or replace==True:
                    l_templ_file_test = templ_path+ '/' + d['_docx_template']
                    if l_templ_file != l_templ_file_test:
                        if os.path.exists(l_templ_file_test)==True:
                            l_templ_file = l_templ_file_test
                            tpl = DocxTemplate(l_templ_file)
                        else:
                            l_templ_file = ''
                            log_txt = 'Template not found: "'+l_templ_file_test+'" for file "'+l_file_name + '.docx"'
                            print(log_txt)
                            logger.error(log_txt)
                    if l_templ_file != '':
                        tpl.render(d)
                        tpl.save(l_file_name + '.docx')
                        is_file_created = True
                        log_txt = 'File generated: "'+ l_file_name + '".docx'
                        print(log_txt)
                        logger.info(log_txt)
                else:
                    log_txt = 'File: '+ l_file_name + '.docx already exists'
                    print(log_txt)
                    logger.info(log_txt)
                if pdf==True and (is_file_created == True or os.path.exists(l_file_name+'.pdf')==False and os.path.exists(l_file_name+'.docx')==True):
                    log_txt = 'Generate PDF:'+l_file_name+'.pdf'
                    print(log_txt)
                    logger.info(log_txt)
                    convert(l_file_name+'.docx', l_file_name+'.pdf', True)
    except Exception as e:
        logger.exception(e)
        print(e)
        traceback.print_exc()

if __name__ == "__main__":
    gen_contract()
