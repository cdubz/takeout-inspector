"""takeout_inspector/mail.py

Defines classes and methods used to parse a Gmail mbox file from Google Takeout and import data in to an sql database.

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
import mailbox
import email
import sqlite3

from datetime import datetime


class Import:
    """Parses and imports Google Takeout mbox file data in to sqlite."""
    def __init__(self, mbox_file, db_file):
        self.email = mailbox.mbox(mbox_file)
        self.conn = sqlite3.connect(db_file)
        self.create_tables()

    def create_tables(self):
        """Creates the required tables for message data storage. Indexes will be added after data import."""
        c = self.conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS messages(
              `message_key` INT,
              `from` TEXT,
              `to` TEXT,
              `subject` TEXT,
              `date` DATETIME,
              `gmail_thread_id` INT,
              `gmail_labels` TEXT,

             );
        ''')
        c.execute('''
             CREATE TABLE IF NOT EXISTS headers(
              `message_key` INT,
              `header` TEXT,
              `value` TEXT,
              FOREIGN KEY(`message_key`) REFERENCES messages(`message_key`)
             );
        ''')
        self.conn.commit()

    def import_messages(self):
        """Imports message details in to the `messages` table and all message headers in to the `headers` table."""
        c = self.conn.cursor()

        query_count = 0
        for key, message in self.email.items():
            for header, value in message.items():
                c.execute('''INSERT INTO `headers` VALUES(?, ?, ?);''', (key, header, value.decode('utf-8')))
                query_count += 1

            mail_from = message.get('From', message.get('from', message.get('FROM', ''))).decode('utf-8')
            mail_to = message.get('To', message.get('to', message.get('TO', ''))).decode('utf-8')
            mail_subject = message.get('Subject', message.get('subject', message.get('SUBJECT', ''))).decode('utf-8')
            mail_date_utc = self._get_message_date(message)
            mail_gmail_id = message.get('X-GM-THRID', '')
            mail_gmail_labels = message.get('X-Gmail-Labels', '').decode('utf-8')

            c.execute('''INSERT INTO `messages` VALUES(?, ?, ?, ?, ?, ?, ?);''',
                      (key, mail_from, mail_to, mail_subject, mail_date_utc, mail_gmail_id, mail_gmail_labels))
            query_count += 1

            if query_count > 100000000:
                self.conn.commit()
                query_count = 0

        c.execute('''CREATE INDEX `id_date` ON `messages` (`date` DESC)''')

        self.conn.commit()

    def _get_message_date(self, message):
        """Finds date and time information for `message` and converts it to ISO-8601 format and UTC timezone."""
        mail_date = message.get('Date', message.get('date', message.get('DATE', ''))).decode('utf-8')
        if not mail_date:
            # The get_from() result always (so far as I have seen) has the date string in the last 30 characters
            mail_date = message.get_from().strip()[-30:]

        mail_date_iso8601 = ''
        if mail_date:
            datetime_tuple = email.utils.parsedate_tz(mail_date)
            unix_time = email.utils.mktime_tz(datetime_tuple)
            mail_date_iso8601 = datetime.utcfromtimestamp(unix_time).isoformat(' ')

        return mail_date_iso8601