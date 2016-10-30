"""takeout_inspector/mail.py

Defines classes and methods used to parse a Google Takeout mbox file and import data in to an sqlite database.

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
import uuid

from datetime import datetime


class Import:
    """Parses and imports Google Takeout mbox file data in to sqlite.

    Keyword arguments:
        anonymize -- Replace email addresses with one-to-one randomized addresses in database tables.
    """
    def __init__(self, mbox_file, db_file, anonymize=True):
        self.email = mailbox.mbox(mbox_file)
        self.conn = sqlite3.connect(db_file)

        self.anonymize = bool(anonymize)
        if anonymize:
            self.anonymize_key = {}

        self._create_tables()
        self.query_count = 0

    def _create_tables(self):
        """Creates the required tables for message data storage. Indexes will be added after data import.
        """
        c = self.conn.cursor()

        c.execute('''
            CREATE TABLE IF NOT EXISTS messages(
              `message_key` INT PRIMARY KEY,
              `from` TEXT,
              `to` TEXT,
              `subject` TEXT,
              `date` DATETIME,
              `gmail_thread_id` INT,
              `gmail_labels` TEXT
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
        c.execute('''
             CREATE TABLE IF NOT EXISTS recipients(
              `message_key` INT,
              `name` TEXT,
              `address` TEXT,
              FOREIGN KEY(`message_key`) REFERENCES messages(`message_key`)
             );
        ''')

        if self.anonymize:
            c.execute('''
                 CREATE TABLE IF NOT EXISTS anonymize_key(
                  `real_address` TEXT,
                  `anon_address` TEXT,
                  `real_name` TEXT,
                  `anon_name` TEXT
                 );
            ''')

        self.conn.commit()

    def import_messages(self):
        """Imports message details in to the `messages` table and all message headers in to the `headers` table.
        """
        c = self.conn.cursor()

        for key, message in self.email.items():
            self._insert_messages(c, key, message)
            self._insert_headers(c, key, message)
            self._insert_recipients(c, key, message)

            if self.query_count > 100000000:
                self.conn.commit()
                self.query_count = 0

        if self.anonymize:
            for real_address, anon_info in self.anonymize_key.iteritems():
                c.execute('''INSERT INTO `anonymize_key` VALUES(?, ?, ?, ?);''',
                          (anon_info[0].decode('utf-8'), anon_info[1].decode('utf-8'),
                           anon_info[2].decode('utf-8'), anon_info[3].decode('utf-8')))

        c.execute('''CREATE INDEX `id_date` ON `messages` (`date` DESC)''')

        self.conn.commit()

    def _insert_recipients(self, c, key, message):
        """Parses contents of the To and CC headers for unique email addresses to be added to the one-row-per-address
        `recipients` table.
        """
        mail_all_to = message.get_all('To', [])
        mail_all_cc = message.get_all('CC', [])
        unique_recipients = list(set(email.utils.getaddresses(mail_all_to + mail_all_cc)))
        for name, address in unique_recipients:
            """Remove the Resourcepart from XMPP addresses.

            https://xmpp.org/rfcs/rfc6122.html#addressing-resource
            """
            address = address.split('/', 1)[0]

            if self.anonymize:
                if address not in self.anonymize_key:
                    anon_name = str(uuid.uuid4())
                    self.anonymize_key[address] = [address, anon_name + '@domain.tld', name, anon_name]
                name = self.anonymize_key[address][3]
                address = self.anonymize_key[address][1]

            c.execute('''INSERT INTO `recipients` VALUES(?, ?, ?);''',
                      (key, name.decode('utf-8'), address.decode('utf-8')))
            self.query_count += 1

    def _insert_headers(self, c, key, message):
        """Adds all headers to `headers`.
        """
        for header, value in message.items():
            c.execute('''INSERT INTO `headers` VALUES(?, ?, ?);''', (key, header, value.decode('utf-8')))
            self.query_count += 1

    def _insert_messages(self, c, key, message):
        """Creates a basic index of important message data in `messages`.
        """
        mail_from = message.get('From', '').decode('utf-8')
        mail_to = message.get('To', '').decode('utf-8')
        mail_subject = message.get('Subject', '').decode('utf-8')
        mail_date_utc = self._get_message_date(message)
        mail_gmail_id = message.get('X-GM-THRID', '')
        mail_gmail_labels = message.get('X-Gmail-Labels', '').decode('utf-8')

        c.execute('''INSERT INTO `messages` VALUES(?, ?, ?, ?, ?, ?, ?);''',
                  (key, mail_from, mail_to, mail_subject, mail_date_utc, mail_gmail_id, mail_gmail_labels))
        self.query_count += 1

    def _get_message_date(self, message):
        """Finds date and time information for `message` and converts it to ISO-8601 format and UTC timezone.
        """
        mail_date = message.get('Date', '').decode('utf-8')
        if not mail_date:
            """The get_from() result always (so far as I have seen!) has the date string in the last 30 characters"""
            mail_date = message.get_from().strip()[-30:]

        mail_date_iso8601 = ''
        if mail_date:
            datetime_tuple = email.utils.parsedate_tz(mail_date)
            unix_time = email.utils.mktime_tz(datetime_tuple)
            mail_date_iso8601 = datetime.utcfromtimestamp(unix_time).isoformat(' ')

        return mail_date_iso8601
