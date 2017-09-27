#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2016 - 2017, asterodeia <noachic.wang@gmail.com>

"""
usage:
./QuestionRecompositor.py -i in.xls(x) -o output(.docx) --no_title --conventional_letter
--change_answer_letter --separate_choices --separator '\s'
"""

import abc
# from docx import Document
from argparse import ArgumentParser
from os import getcwd
from os import path

import pandas as pd
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.shared import Cm
import re
import logging


class QuestionRecompositor(object):
    def __init__(self, config):
        self.config = config
        pass

    @staticmethod
    def column_name_2_index(name):
        return ord(name) - ord('A') + 1

    # TODO: CHANGE HERE TO SUPPORT REAL COLUMN NAME, NOT ONLY A SINGLE LETTER
    @staticmethod
    def column_index_2_name(index):
        return chr(int(index) + ord('A') - 1)

    def format_question(self, data):
        if self.config.conventional_letter:
            for entry in data:
                for index in range(len(entry['choices'])):
                    entry['choices'][index] = \
                        chr(ord('A') + index) + '. ' + entry['choices'][index]

        if self.config.change_answer_letter:
            for entry in data:
                original_answer = entry['answers'][0]
                entry['answers'] = []
                for c in original_answer:
                    entry['answers'].append(
                        self.column_index_2_name(c)
                    )
                    # for index in range(len(entry['answers'])):
                    # entry['answers'][index] = \
                    # self.column_index_2_name(entry['answers'][index])

        return data

    def recompose2(self):
        try:
            doc = Document(self.config.word_template)
            doc.save('test.docx')
        except Exception as e:
            print('create word application failed, please try again')
            print(e)

        else:
            try:
                data = self.read()
                data = self.format_question(data)

                for index, entry in enumerate(data):
                    print('processing question No. {}'.format(index))
                    # question
                    for item in entry['question']:
                        doc.add_paragraph(item, style='Heading 2')

                    # choices:
                    if 'choices' in entry:
                        max_ans_len = max(map(len, entry['choices']))
                        p = doc.add_paragraph('')
                        pf = p.paragraph_format
                        if max_ans_len <= 10:
                            separater = ['\t', '\t']

                            for pos in [1 + x * 4.5 for x in range(4)]:
                                pf.tab_stops.add_tab_stop(Cm(pos))

                        elif max_ans_len <= 22:
                            separater = ['\t', '\n\t']
                            pf.tab_stops.add_tab_stop(Cm(1))
                            pf.tab_stops.add_tab_stop(Cm(10))

                        else:
                            separater = ['\n\t', '\n\t']
                            pf.tab_stops.add_tab_stop(Cm(1))

                        choice_count = len(entry['choices'])
                        choice_string = ['\t']
                        for cnt, item in enumerate(entry['choices']):
                            if cnt != choice_count - 1:
                                choice_string += item + separater[cnt % 2]
                            else:
                                choice_string += item
                        p.add_run(''.join(choice_string))

                    # answer
                    answer_string = ''.join(entry['answers'])
                    p = doc.add_paragraph("答案：" + answer_string, style='答案')
                    pf = p.paragraph_format
                    pf.tab_stops.add_tab_stop(Cm(18.5), alignment=WD_TAB_ALIGNMENT.RIGHT)

                # TODO: filename
                doc.save(path.join(getcwd(), 'test.docx'))
            except Exception as e:
                print(e)
            finally:
                pass


class Config(object):
    def __init__(self):
        pass

    pass


def get_data(book, sheet, no_header) -> pd.DataFrame:
    """
    :param header_row_count: None for no header, int for a list like header
    :return: DataFrame, if there's no header, a zero based header will be added
    """

    header_row_count = None if no_header else 0
    dataframe = pd.read_excel(book, sheet, header=header_row_count, dtype=str)

    if no_header:
        column_count = dataframe.shape[1]
        column_name = [number_2_char(i) for i in range(column_count)]
        dataframe.columns = column_name

    return dataframe


def shift_character(ch, i):
    return chr(ord(ch) + i)


def number_2_char(n, base=0):
    assert n >= base
    return chr(ord('A') + n - base)


class ObjectiveQuestions(object):
    __metaclass__ = abc.ABCMeta

    def __init__(self, config):
        self.config = config
        self.data = None

    @abc.abstractmethod
    def write_doc(self, doc, data):
        return

    @abc.abstractmethod
    def process_data(self, data):
        return


class SingleChoice(ObjectiveQuestions):
    def __init__(self, config):
        temp_config = config.copy()
        if not config['use_column_name']:
            for col in ['question_col', 'answer_col', 'choice_col']:
                if isinstance(config[col], str):
                    temp_config[col] = list(config[col])

        super().__init__(temp_config)

    def process_data(self, data):

        # this function only use the process block in config
        config = self.config['process']

        needed_data_column = config['question_col'] + \
                             config['answer_col'] + \
                             config['choice_col']
        needed_data = data[needed_data_column]

        if config['trim']:
            pattern = re.compile(r'[\r\n]')
            # def trim(s):
            # return re.sub(pattern, '', s)

            needed_data = needed_data.applymap(lambda s: re.sub(pattern, '', s))

        if config['add_conventional_letter']:

            for i in range(len(config['choice_col'])):
                conventional_letter = number_2_char(i)

                def add_conventional_letter(s):
                    return s if s == 'nan' else '{}. {}'.format(conventional_letter, s)

                # don't know if there is a better way to update just one column of dataframe
                ser = needed_data[config['choice_col'][i]]
                ser = ser.apply(add_conventional_letter)

                needed_data[config['choice_col'][i]] = ser

        if config['change_answer_letter']:
            # TODO: first column of answer, change this?
            answer_col = needed_data[config['answer_col'][0]]

            # def change_answer_letter(s):
                # answer = [number_2_char(int(ch), 1) for ch in s]
                # return ''.join(answer)

            ser = answer_col.apply(lambda s: [number_2_char(int(ch), 1) for ch in s])
            needed_data[config['answer_col'][0]] = ser

        ret = []
        for entry in needed_data:
            pass

        return needed_data

    def write_doc(self, doc, data):
        # add title
        doc.add_paragraph(self.config['title'], style='Heading 1')


class MultiChoice(ObjectiveQuestions):
    pass


class TrueFalseChoice(ObjectiveQuestions):
    pass


class BriefAnswer(ObjectiveQuestions):
    pass


def main():
    parser = ArgumentParser(description='a python script to convert structured excel file to word')
    parser.add_argument('-i', '--input', metavar='filename', action='store',
                        help='source structured excel file')
    parser.add_argument('-o', '--output', metavar='filename', action='store',
                        help='target word file name (default: output.docx)')
    parser.add_argument('-T', '--no-title', action='store_true',
                        help="to specify the excel file don't have a title row")
    args = parser.parse_args()

    pass


def dummy_main():
    c = Config
    c.workbook = r"E:\PythonProjects\QuestionRecompositor\data\线路架空线专业判断.xls"
    c.sheetname = 'Sheet1'
    c.no_title = False
    # c.needed_columns = {'question':['E'], 'choices':['F', 'G', 'H', 'I'],
    # 'answers':['K']}
    c.needed_columns = {'question': ['D'], 'answers': ['E']}
    # c.word_template = r"E:\PythonProjects\QuestionRecompositor\data\template.docx"
    c.word_template = r"E:\PythonProjects\QuestionRecompositor\data\template.docx"
    c.conventional_letter = False
    c.change_answer_letter = False
    qr = QuestionRecompositor(c)
    # qr.read()
    qr.recompose2()


def dummy_main2():
    cwd = getcwd()
    print(cwd)
    rwp = ObjectiveQuestions('a')
    workbook = r"线路架空线专业判断.xls"
    sheetname = 'Sheet1'
    rwp.read_data(path.join(cwd, 'data', workbook), sheetname)


if __name__ == '__main__':
    dummy_main2()
