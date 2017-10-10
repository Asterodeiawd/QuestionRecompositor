# -*- coding: utf-8 -*-

import unittest
from unittest import TestCase

import QuestionRecompositor as qr


class TestQuestionRecompositor(TestCase):
    def test_shift_character(self):
        self.assertEqual(qr.shift_character('A', 1), 'B')
        self.assertEqual(qr.shift_character('B', -1), 'A')
        # self.assertRaises(qr.shift_character('B', -2), ValueError)
        self.assertRaises(ValueError, qr.shift_character, 'B', -2)
        self.assertRaises(ValueError, qr.shift_character, '0', 1)



if __name__ == '__main__':
    unittest.main()
