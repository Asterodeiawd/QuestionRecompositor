# -*- coding: utf-8 -*-

import unittest
from unittest import TestCase

import QuestionRecompositor.QuestionRecompositor2 as qr


class TestQuestionRecompositor(TestCase):
    def test_shift_character(self):
        self.assertEqual(qr.shift_character('A', 1), 'B')
        self.assertEqual(qr.shift_character('B', -1), 'A')
        # TODO: should implement later
        # self.assertRaises(qr.shift_character('B', -2), Exception)


if __name__ == '__main__':
    unittest.main()
