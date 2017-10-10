#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2016 - 2017, asterodeia <d_wang_890227@outlook.com>


import logging
import re

import pandas as pd
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.shared import Cm


def concatenate(array: list, sep: str = '') -> str:
    ret = []
    for item in array:
        ret.append(item)
        ret.append(sep)

    ret.pop()
    return ''.join(ret)


def load_data(book, sheet, config) -> pd.DataFrame:
    logging.info('reading data from excel file')

    no_header = config['no_header']
    header_row_count = None if no_header else 0
    try:
        dataframe = pd.read_excel(book, sheet, header=header_row_count, dtype=str)
    except FileNotFoundError:
        logging.error('data source file: {} not found, please check the file name'.format(book))
        logging.info('exit, please check and rerun this script')
        exit(-1)

    except Exception as e:
        logging.error('unexpected error: {}'.format(e))
        exit(-1)

    else:

        if no_header or not config['use_column_name']:
            column_count = dataframe.shape[1]
            column_name = [number_2_char(i) for i in range(column_count)]
            dataframe.columns = column_name

        return dataframe


def shift_character(ch, i):
    """shift character according to ascii code"""

    if not str.isalpha(ch):
        raise ValueError('only support characters from A-Z and a-z')

    result = chr(ord(ch) + i)
    if not str.isalpha(result):
        raise ValueError('not a character after shifting')

    return result


# TODO: EXCEL-LIKE NUMBER TO CHARACTER
def number_2_char(n, base=0):
    assert n >= base
    return chr(ord('A') + n - base)


def process_data(data, config):
    logging.info('starting to process data')

    if 'choice_col' in config:
        needed_data_column = config['question_col'] + config['answer_col'] + config['choice_col']
    else:
        needed_data_column = config['question_col'] + config['answer_col']

    needed_data = data[needed_data_column]

    if config['trim']:
        logging.info('\ttrimming data')
        pattern = re.compile(r'[\r\n]')
        needed_data = needed_data.applymap(lambda s: re.sub(pattern, '', s))

    # TODO: if not choice questions, skip this
    if 'add_conventional_letter' in config and config['add_conventional_letter']:

        logging.info('\tadding conventional letter to choices')
        for i in range(len(config['choice_col'])):
            conventional_letter = number_2_char(i)

            def add_conventional_letter(s):
                return s if s == 'nan' else '{}. {}'.format(conventional_letter, s)

            # don't know if there is a better way to update just one column of dataframe
            ser = needed_data[config['choice_col'][i]]
            ser = ser.apply(add_conventional_letter)

            needed_data[config['choice_col'][i]] = ser

    if 'change_answer_letter' in config and config['change_answer_letter']:
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
        if 'choice_col' in config:
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


# builder
class DocumentWriter(object):
    def __init__(self, template):
        self._document = None
        self._template = template

    def create_document(self):
        # TODO: check if template exist
        try:
            self._document = Document(self._template)

        # TODO: NOT CORRECT RIGHT NOW, CHANGE LATER
        except ValueError as e:
            logging.error(e)

    def get_document(self):
        return self.document

    @property
    def document(self):
        return self._document

    # TODO: this is a fake one for reference only
    def get_document_params(self):
        section = self.document.sections[-1]
        left_margin = section.left_margin.cm
        right_margin = section.right_margin.cm
        top_margin = section.top_margin.cm
        bottom_margin = section.bottom_margin.cm
        height = section.page_height.cm
        width = section.page_width.cm
        return height, width, left_margin, right_margin, top_margin, bottom_margin

    def write_title(self, title, level):
        assert 0 <= level <= 9
        self.document.add_heading(title, level)

    def write_question(self, question, style):
        for item in question:
            self.document.add_paragraph(item, style=style)

    def write_choices(self, choices, style):
        # this method is not interested by every derived class,
        # so doesn't need a default implementation
        raise NotImplementedError('please specify a derived class')
        pass

    def write_answer(self, answer, style):
        for item in answer:
            self.document.add_paragraph(item, style=style)


class SingleChoiceWriter(DocumentWriter):
    def __init__(self, template=None):
        super().__init__(template)

    def _get_choice_tabstops_pos(self, choices: list):

        params = self.get_document_params()
        max_ans_len = max(map(len, choices))

        # TODO: should be changed here, using document params maybe
        if max_ans_len <= 10:
            tab_pos = [1 + x * 4.5 for x in range(4)]
        elif max_ans_len <= 22:
            tab_pos = [1, 10]
        else:
            tab_pos = [1]

        return tab_pos

    @staticmethod
    def _get_formatted_choice(choices: list, choice_count_per_line):
        formatted_choices = ['\t']
        current_item_count = choice_count_per_line

        for idx, choice in enumerate(choices, 1):
            formatted_choices.append(choice)
            current_item_count -= 1

            # decide string separator
            if idx != len(choices):
                if current_item_count != 0:
                    formatted_choices.append('\t')
                else:
                    formatted_choices.append('\n\t')
                    current_item_count = choice_count_per_line
            else:
                pass
                # don't need anything

        return ''.join(formatted_choices)

    # the default version is ok
    # def write_title(self, title, level):
    # def write_question(self, question, style):

    def write_choices(self, choices: list, style):

        tab_pos = self._get_choice_tabstops_pos(choices)
        formatted_choice = self._get_formatted_choice(choices, len(tab_pos))

        p = self.document.add_paragraph(''.join(formatted_choice))
        pf = p.paragraph_format
        for pos in tab_pos:
            pf.tab_stops.add_tab_stop(Cm(pos))

    def write_answer(self, answer, style):

        if isinstance(answer, list):
            """
            formatted_answer = []
            for ans in answer:
                answer_string = formatted_answer.append(ans)
            formatted_answer = ''.join(formatted_answer)
            """
            formatted_answer = concatenate(answer)
        else:
            formatted_answer = answer

        p = self.document.add_paragraph("\t答案：{}".format(formatted_answer), style=style)
        pf = p.paragraph_format
        pf.tab_stops.add_tab_stop(Cm(18), alignment=WD_TAB_ALIGNMENT.RIGHT)


# TODO: now this class is the same with SingleChoiceWriter, new class ObjectiveQuestionsWriter?
class MultiChoiceWriter(DocumentWriter):
    def __init__(self, template=None):
        super().__init__(template)

    def _get_choice_tabstops_pos(self, choices: list):

        params = self.get_document_params()
        max_ans_len = max(map(len, choices))

        # TODO: should be changed here, using document params maybe
        if max_ans_len <= 10:
            tab_pos = [1 + x * 4.5 for x in range(4)]
        elif max_ans_len <= 22:
            tab_pos = [1, 10]
        else:
            tab_pos = [1]

        return tab_pos

    @staticmethod
    def _get_formatted_choice(choices: list, choice_count_per_line):
        formatted_choices = ['\t']
        current_item_count = choice_count_per_line

        for idx, choice in enumerate(choices, 1):
            formatted_choices.append(choice)
            current_item_count -= 1

            # decide string separator
            if idx != len(choices):
                if current_item_count != 0:
                    formatted_choices.append('\t')
                else:
                    formatted_choices.append('\n\t')
                    current_item_count = choice_count_per_line
            else:
                pass
                # don't need anything

        return ''.join(formatted_choices)

    def write_choices(self, choices: list, style):

        tab_pos = self._get_choice_tabstops_pos(choices)
        formatted_choice = self._get_formatted_choice(choices, len(tab_pos))

        p = self.document.add_paragraph(''.join(formatted_choice))
        pf = p.paragraph_format
        for pos in tab_pos:
            pf.tab_stops.add_tab_stop(Cm(pos))

    def write_answer(self, answer, style):

        if isinstance(answer, list):
            """
            formatted_answer = []
            for ans in answer:
                answer_string = formatted_answer.append(ans)
            formatted_answer = ''.join(formatted_answer)
            """
            formatted_answer = concatenate(answer)
        else:
            formatted_answer = answer

        p = self.document.add_paragraph("\t答案：{}".format(formatted_answer), style=style)
        pf = p.paragraph_format
        pf.tab_stops.add_tab_stop(Cm(18), alignment=WD_TAB_ALIGNMENT.RIGHT)


class TrueFalseWriter(DocumentWriter):
    def __init__(self, template=None):
        super().__init__(template)

    def write_answer(self, answer, style):
        # this type of questions should have only one answer
        assert len(answer) == 1
        p = self.document.add_paragraph("\t答案：" + answer[0], style='答案')
        pf = p.paragraph_format
        pf.tab_stops.add_tab_stop(Cm(18), alignment=WD_TAB_ALIGNMENT.RIGHT)


class SubjectiveQuestionWriter(DocumentWriter):
    def __init__(self, template=None):
        super().__init__(template)

    def write_answer(self, answer, style):
        self.document.add_paragraph("\t答案：")

        # remove wrong '\n' characters in answer, the correct one is followed by spaces
        # in *MY* source file, you may have to change this behaviour
        pattern = re.compile(r'(\n)\S')
        for item in answer:
            new_answer = re.sub(pattern, '', item)
            self.document.add_paragraph(new_answer, style=style)


class QuestionRecompositor(object):
    # writer: for later use, now is None
    def __init__(self, config, writer):
        self._config = config
        self._writer = writer
        # this may change later
        self.data_processor = process_data

    def set_config(self, config):
        self._config = config

    def _load_data(self, file_name, sheet_name):
        data_frame = load_data(file_name, sheet_name, self._config.read)
        return data_frame

    def recompose(self, file_name, sheet_name, builder: DocumentWriter):
        data_frame = self._load_data(file_name, sheet_name)
        data_frame = self.data_processor(data_frame, self._config.process)
        # self._writer(data_frame, doc, self._config.write)

        builder.create_document()
        builder.write_title(self._config.write['title'], level=0)

        logging.info('writing questions to document, {} in total'.format(len(data_frame)))

        for entry in data_frame:
            builder.write_question(entry['question'], 'Heading 2')

            if entry['choices']:
                # this is an objective question
                builder.write_choices(entry['choices'], 'Normal')

            if not self._config.write['hide_answer']:
                builder.write_answer(entry['answer'], '答案')
