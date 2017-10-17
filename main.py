#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2016 - 2017, asterodeia <d_wang_890227@outlook.com>


from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import ProcessPoolExecutor
from os import getcwd
from os import path
import time

from Config import load_configs_from_file
from QuestionRecompositor import *

logging.basicConfig(level=logging.WARNING,
                    format='%(asctime)s %(levelname)s: %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')


def worker(source, sheet, question_type, out_dir, configs):
    document_template_path = r'.\data\template.docx'

    if question_type in ('SingleChoice', 'MultiChoice'):
        builder = ObjectiveQuestionWriter(document_template_path)
    elif question_type in ('CaseAnalyse', 'BrieflyAnswer'):
        builder = SubjectiveQuestionWriter(document_template_path)
    else:
        builder = TrueFalseWriter(document_template_path)
    config = configs[question_type]

    qr = QuestionRecompositor(config, None)
    qr.recompose(source, sheet, builder)
    doc = builder.document
    doc.save(path.join(out_dir, config.write['title']) + '.docx')


def main():
    cwd = getcwd()
    base_dir = path.join(cwd, 'data')
    source = {r'线路架空线专业单选.xls': 'SingleChoice',
              r'线路架空线专业多选.xls': 'MultiChoice',
              r'线路架空线专业判断.xls': 'TrueFalse',
              r'线路架空线专业案例.xls': 'CaseAnalyse',
              r'线路架空线专业问答.xls': 'BrieflyAnswer'}
    config_path = path.join(cwd, 'config.json')

    configs = load_configs_from_file(config_path)

    sheet_name = 'Sheet1'
    output_directory = 'E:\output'

    # with ThreadPoolExecutor() as executor:
    with ProcessPoolExecutor() as executor:
        for k, v in source.items():
            executor.submit(worker, path.join(base_dir, k), sheet_name, v, output_directory, configs)

if __name__ == '__main__':
    start_time = time.time()
    main()
    end_time = time.time()
    print('used time = {}'.format(end_time - start_time))
