#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2016 - 2017, asterodeia <d_wang_890227@outlook.com>


from os import getcwd
from os import path

from Config import load_configs_from_file
from QuestionRecompositor import *

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')


def main():
    cwd = getcwd()
    base_dir = path.join(cwd, 'data')
    source_path = [r'线路架空线专业单选.xls', r'线路架空线专业多选.xls', r'线路架空线专业判断.xls',
                   r'线路架空线专业案例.xls', r'线路架空线专业问答.xls']
    config_path = path.join(cwd, 'config.json')
    document_template_path = path.join(cwd, 'data', 'template.docx')

    configs = load_configs_from_file(config_path)

    sheet_name = 'Sheet1'

    index = 0
    for question_type, c in configs.items():
        if question_type in ('SingleChoice', 'MultiChoice'):
            builder = ObjectiveQuestionWriter(document_template_path)
        elif question_type in ('CaseAnalyse', 'BrieflyAnswer'):
            builder = SubjectiveQuestionWriter(document_template_path)
        else:
            builder = TrueFalseWriter(document_template_path)

        qr = QuestionRecompositor(c, None)
        qr.recompose(os.path.join(base_dir, source_path[index]), sheet_name, builder)
        doc = builder.document
        doc.save(path.join(cwd, c.write['title']) + '.docx')

        index += 1


if __name__ == '__main__':
    main()
