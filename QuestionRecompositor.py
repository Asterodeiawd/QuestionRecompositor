"""
Question Recompositor
"""
"""
usage:
./QuestionRecompositor.py -i in.xls(x) -o output(.docx) --no_title --conventional_letter
--change_answer_letter --separate_choices --separator '\s'
"""

from win32com.client.gencache import EnsureDispatch
from win32com.client import constants as c
import pandas as pd
import numpy as np
from docx.shared import Cm
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT

from argparse import ArgumentParser
from os import path
from os import getcwd


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

    def recompose(self):
        try:
            word = EnsureDispatch('Word.Application')
            word.Visible = True

        except Exception as e:
            print('create word application failed, please try again')
            print(e)

        else:
            try:
                doc = word.Documents.Add(Template=self.config.word_template)
                data = self.read()
                data = self.format_question(data)

                for index, entry in enumerate(data):
                    print('processing question No. {}'.format(index))
                    # question
                    for item in entry['question']:
                        word.Selection.Style = doc.Styles('标题 2')
                        word.Selection.TypeText(item)
                        word.Selection.TypeParagraph()

                    # CentimetersToPoints = word.CentimetersToPoints
                    CentimetersToPoints = lambda x: x * 28.35
                    # choices:
                    # max_ans_len = max(len(entry['choices']))
                    max_ans_len = max(map(len, entry['choices']))
                    if max_ans_len <= 12:
                        separater = ['\t', '\t']

                        word.Selection.Paragraphs.TabStops.Add(Position=CentimetersToPoints(1),
                                                               Alignment=c.wdAlignTabLeft)
                        word.Selection.Paragraphs.TabStops.Add(Position=CentimetersToPoints(5),
                                                               Alignment=c.wdAlignTabLeft)
                        word.Selection.Paragraphs.TabStops.Add(Position=CentimetersToPoints(9),
                                                               Alignment=c.wdAlignTabLeft)
                        word.Selection.Paragraphs.TabStops.Add(Position=CentimetersToPoints(13),
                                                               Alignment=c.wdAlignTabLeft)

                    elif max_ans_len <= 22:
                        separater = ['\t', '\n\t']
                        word.Selection.Paragraphs.TabStops.Add(Position=CentimetersToPoints(1),
                                                               Alignment=c.wdAlignTabLeft)
                        word.Selection.Paragraphs.TabStops.Add(Position=CentimetersToPoints(9),
                                                               Alignment=c.wdAlignTabLeft)

                    else:
                        separater = ['\n\t', '\n\t']
                        word.Selection.Paragraphs.TabStops.Add(Position=CentimetersToPoints(1),
                                                               Alignment=c.wdAlignTabLeft)

                    choice_count = len(entry['choices'])
                    choice_string = ['\t']
                    for cnt, item in enumerate(entry['choices']):
                        if cnt != choice_count - 1:
                            choice_string += item + separater[cnt % 2]
                        else:
                            choice_string += item
                    word.Selection.TypeText(''.join(choice_string))
                    word.Selection.TypeParagraph()

                    # answer
                    answer_string = ''.join(entry['answers'])
                    word.Selection.Style = doc.Styles('答案')
                    word.Selection.TypeText("答案：" + answer_string)
                    word.Selection.TypeParagraph()
                    word.Selection.TypeParagraph()

                doc.SaveAs2(path.join(getcwd(), 'test.docx'))
            except Exception as e:
                print(e)
            finally:
                word.Quit()

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

    def read(self):
        try:
            excel = EnsureDispatch('Excel.Application')
        except Exception as e:
            print('create excel application failed, please try again')
            print(e)
            return None
            # exit(code='1')

        else:
            try:
                wkb_source = excel.Workbooks.Open(self.config.workbook)
                wks_source = wkb_source.Worksheets(self.config.sheetname)

                # starts from 1, unlike arrays
                start_row = 1 if self.config.no_title else 2
                end_row = wks_source.UsedRange.Rows.Count

                data = []

                for row in range(start_row, end_row + 1):
                    entry = dict()
                    for key, value in self.config.needed_columns.items():
                        entry[key] = []
                        for col in value:
                            entry[key].append(wks_source.Cells(row, col).Text)

                    data.append(entry)

                wkb_source.Close(SaveChanges=False)
                return data

            except Exception as e:
                print(e)
                print('open file {} failed, please make sure the file exist or can be accessed')

            finally:
                excel.Quit()


class Config(object):
    def __init__(self):
        pass

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


if __name__ == '__main__':
    dummy_main2()
