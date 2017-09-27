#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2016 - 2017, asterodeia <noachic.wang@gmail.com>

"""
usage:
./QuestionRecompositor.py -i in.xls(x) -o output(.docx) --no_title --conventional_letter
--change_answer_letter --separate_choices --separator '\s'
"""

import json
import logging
import re
from os import getcwd
from os import path

import pandas as pd
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.shared import Cm

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')


def get_raw_data(book, sheet, config) -> pd.DataFrame:
    logging.info('reading data from excel file')

    no_header = config['no_header']
    header_row_count = None if no_header else 0
    try:
        dataframe = pd.read_excel(book, sheet, header=header_row_count, dtype=str)
    except FileNotFoundError as e:
        logging.error('data source file: {} not found, please check the file name'.format(book))
        logging.info('exit, please check and rerun this script')
        exit(-1)

    except Exception as e:
        logging.error('unknown error')
        exit(-1)

    else:

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


def unify_process(config):
    logging.info('unifying process part in config')
    if not config['use_column_name']:
        for col in ['question_col', 'answer_col', 'choice_col']:
            if isinstance(config[col], str):
                config[col] = list(config[col])


def process_data(data, config):
    logging.info('starting to process data')

    needed_data_column = config['question_col'] + config['answer_col'] + config['choice_col']
    needed_data = data[needed_data_column]

    if config['trim']:
        logging.info('\ttrimming data')
        pattern = re.compile(r'[\r\n]')
        needed_data = needed_data.applymap(lambda s: re.sub(pattern, '', s))

    if config['add_conventional_letter']:

        logging.info('\tadding conventional letter to choices')
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
        logging.info('\tchanging answer number to letters')
        answer_col = needed_data[config['answer_col'][0]]

        # def change_answer_letter(s):
        # answer = [number_2_char(int(ch), 1) for ch in s]
        # return ''.join(answer)

        ser = answer_col.apply(lambda s: ''.join([number_2_char(int(ch), 1) for ch in s]))
        needed_data[config['answer_col'][0]] = ser

    logging.info('\treformatting data')
    ret = []
    for index, row in needed_data.iterrows():
        entry = dict(question=[], choices=[], answer=[])

        # question
        for q in row[config['question_col']]:
            entry['question'].append(q)

        # choices
        for c in row[config['choice_col']]:
            # TODO: 'nan' values should be handle more carefully
            if c != 'nan':
                entry['choices'].append(c)

        # answer
        for a in row[config['answer_col']]:
            entry['answer'].append(a)

        ret.append(entry)

    logging.info('done processing data')
    # return formatted question data
    return ret


def get_document_params(doc):
    section = doc.sections[-1]
    left_margin = section.left_margin.cm
    right_margin = section.right_margin.cm
    top_margin = section.top_margin.cm
    bottom_margin = section.bottom_margin.cm
    height = section.page_height.cm
    width = section.page_width.cm
    print(height, width, left_margin, right_margin, top_margin, bottom_margin)


def write_doc_title(doc: Document, config):
    doc.add_heading(config['title'], 0)
    # doc.paragraphs[0].delete()


def write_doc_sc(data, doc: Document, config):
    # add document title
    write_doc_title(doc, config)

    logging.info('writing questions to document, {} in total'.format(len(data)))

    for entry in data:

        # question
        for item in entry['question']:
            doc.add_paragraph(item, style='Heading 2')

        # choices:
        max_ans_len = max(map(len, entry['choices']))
        if max_ans_len <= 10:
            separater = ['\t', '\t']
            tab_pos = [1 + x * 4.5 for x in range(4)]

        elif max_ans_len <= 22:
            separater = ['\t', '\n\t']
            tab_pos = [1, 10]

        else:
            separater = ['\n\t', '\n\t']
            tab_pos = [1]

        choice_count = len(entry['choices'])
        choice_string = ['\t']
        for cnt, item in enumerate(entry['choices']):
            if cnt != choice_count - 1:
                choice_string += item + separater[cnt % 2]
            else:
                choice_string += item

        p = doc.add_paragraph(''.join(choice_string))
        pf = p.paragraph_format
        for pos in tab_pos:
            pf.tab_stops.add_tab_stop(Cm(pos))

        # answer
        if not config['hide_answer']:
            answer_string = ''.join(entry['answer'])
            p = doc.add_paragraph("\t答案：" + answer_string, style='答案')
            pf = p.paragraph_format
            pf.tab_stops.add_tab_stop(Cm(18), alignment=WD_TAB_ALIGNMENT.RIGHT)

    return doc


# multiple choice
def write_doc_mc(data, doc: Document, config):
    # add document title
    doc.add_heading(config['title'], 0)

    logging.info('writing questions to document, {} in total'.format(len(data)))

    for entry in data:

        # question
        for item in entry['question']:
            doc.add_paragraph(item, style='Heading 2')

        # choices:
        max_ans_len = max(map(len, entry['choices']))
        if max_ans_len <= 10:
            separater = ['\t', '\t']
            tab_pos = [1 + x * 4.5 for x in range(4)]

        elif max_ans_len <= 22:
            separater = ['\t', '\n\t']
            tab_pos = [1, 10]

        else:
            separater = ['\n\t', '\n\t']
            tab_pos = [1]

        choice_count = len(entry['choices'])
        choice_string = ['\t']
        for cnt, item in enumerate(entry['choices']):
            if cnt != choice_count - 1:
                choice_string += item + separater[cnt % 2]
            else:
                choice_string += item

        p = doc.add_paragraph(''.join(choice_string))
        pf = p.paragraph_format
        for pos in tab_pos:
            pf.tab_stops.add_tab_stop(Cm(pos))

        # answer
        if not config['hide_answer']:
            answer_string = ''.join(entry['answer'])
            p = doc.add_paragraph("\t答案：" + answer_string, style='答案')
            pf = p.paragraph_format
            pf.tab_stops.add_tab_stop(Cm(18), alignment=WD_TAB_ALIGNMENT.RIGHT)

    return doc


def main():
    cwd = getcwd()
    source_path = path.join(cwd, 'data', r'线路架空线专业单选.xls')
    config_path = path.join(cwd, 'config.json')
    document_template_path = path.join(cwd, 'data', 'template.docx')

    configs = load_configs_from_file(config_path)
    c = configs[0]

    unify_process(c.process)
    data = get_raw_data(source_path, 'Sheet1', c.read)
    parsed_data = process_data(data, c.process)

    doc = Document(document_template_path)
    write_doc_sc(parsed_data, doc, c.write)
    doc.save(path.join(cwd, 'dx.docx'))


class Config(object):
    def __init__(self, data):
        # configs should include 4 parts: name, read method, process method and write method
        # so here don't have to check before assign
        # self.name = None
        # self.read = None
        # self.process = None
        # self.write = None

        self.load(data)

    def load(self, data):
        self.name = data['name']
        self.read = data['read']
        self.process = data['process']
        self.write = data['write']


def load_configs_from_file(filename):
    configs = []

    with open(filename, mode='r', encoding='utf-8') as f:
        raw_configs = json.loads(f.read())

        # configs should be in one list, single config can omit the outter list
        if isinstance(raw_configs, list):
            for raw_config in raw_configs:
                configs.append(Config(raw_config))
        else:
            configs.append(Config(raw_configs))

    return configs


if __name__ == '__main__':
    # cwd = getcwd()
    # document_template_path = path.join(cwd, 'data', 'template.docx')
    # doc = Document(document_template_path)
    # get_document_params(doc)
    main()
