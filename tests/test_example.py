# -*- coding: utf-8 -*-

import sys
import unittest
from datetime import datetime
import pdb
import re


from unipath import Path
PROJECT_DIR = Path(__file__).ancestor(2)
sys.path.append(PROJECT_DIR)

from app import app




class SomeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.client = app.test_client()


    def test_main_returns_200(self):
        response = self.client.get('/')
        self.assertEqual(response.status_code, 200)
        assert False



if __name__ == "__main__":
    unittest.main()