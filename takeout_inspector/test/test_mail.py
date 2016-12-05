"""takeout_inspector/test/mail.py

Defines unittest tests for mail functions.

Copyright (c) 2016 Christopher Charbonneau Wells

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

"""
import os
import unittest

from takeout_inspector import mail


class Mail(unittest.TestCase):

    def setUp(self):
        self.m = mail.Import(settings_file='takeout_inspector/test/data/test.cfg')
        self.m.import_messages()

    def tearDown(self):
        os.remove(self.m.config.get('mail', 'db_file'))

    def test_tables(self):
        self.assertTrue(os.path.isfile(self.m.config.get('mail', 'db_file')), 'Database file not created.')

if __name__ == '__main__':
    unittest.main()
