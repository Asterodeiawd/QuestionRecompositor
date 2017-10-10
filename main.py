#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2016 - 2017, asterodeia <d_wang_890227@outlook.com>

"""
usage:
./QuestionRecompositor.py -i in.xls(x) -o output(.docx) --no_title --conventional_letter
--change_answer_letter --separate_choices --separator '\s'
"""

from argparse import ArgumentParser
from os import getcwd
from os import path

from Config import load_configs_from_file
# from QuestionRecompositor import (QuestionRecompositor,
# write_doc_sc)
from QuestionRecompositor import *

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')


def main():
    cwd = getcwd()
    source_path = path.join(cwd, 'data', r'线路架空线专业问答.xls')
    # source_path = r"C:\Users\asterodeia\Desktop\变压器新增题库汇总.xlsx"
    config_path = path.join(cwd, 'config.json')
    document_template_path = path.join(cwd, 'data', 'template.docx')

    configs = load_configs_from_file(config_path)
    c = configs[3]
    print(c.name)

    name = 'Sheet1'

    # sc_builder = SingleChoiceWriter(document_template_path)
    # builder = TrueFalseWriter(document_template_path)
    builder = SubjectiveQuestionWriter(document_template_path)
    qr = QuestionRecompositor(c, None)
    qr.recompose(source_path, name, builder)
    doc = builder._document
    doc.save(path.join(cwd, name) + '5.docx')


def main2():
    parser = ArgumentParser(description='a python script to convert structured excel to word')
    parser.add_argument('-s', '--single-choice', metavar='filename', action='store',
                        help='source structured excel file')
    parser.add_argument('-m', '--multi-choice', metavar='filename', action='store',
                        help='target word file name (default: output.docx)')
    parser.add_argument('-t', '--true-false', action='store_true',
                        help="to specify the excel file don't have a title row")
    parser.add_argument('-o', '--output-dir', metavar='dirname', action='store',
                        help='target directory to store the generated word files')
    args = parser.parse_args()
    pass


def temp_test():
    shift_character('1', 1)


if __name__ == '__main__':
    main()
    # temp_test()
