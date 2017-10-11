#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2016 - 2017, asterodeia <d_wang_890227@outlook.com>

import json
import logging


class Config(object):
    def __init__(self, data):
        # configs should include 4 parts: name, read method, process method and write method
        # so here don't have to check before assign
        self.read = None
        self.process = None
        self.write = None

        self.load(data)

    def load(self, data):
        self.read = data['read']
        self.process = data['process']
        self.write = data['write']

    def unify(self):
        if not self.read['use_column_name']:
            for col in ['question_col', 'answer_col', 'choice_col']:
                if isinstance(self.process[col], str):
                    self.process[col] = list(self.process[col])


def load_configs_from_file(filename: str) -> dict:
    """
     helper function to load configs from file
    :param filename: file contains the config(s)
    :return: list of Config
    """
    configs = {}

    with open(filename, mode='r', encoding='utf-8') as f:
        raw_configs = json.loads(f.read())

        for key, value in raw_configs.items():
            configs[key] = Config(value)

    logging.info('unifying process part in config')
    for config in configs.values():
        config.unify()

    return configs
