#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2016 - 2017, asterodeia <d_wang_890227@outlook.com>


import os


def concatenate(array: list, sep: str = '') -> str:
    """ helper function to concatenate list of strings to one string, separated by sep """
    ret = []
    for item in array:
        ret.append(item)
        ret.append(sep)

    ret.pop()
    return ''.join(ret)


def check_file_exist(filename):
    return os.path.isfile(filename)


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
