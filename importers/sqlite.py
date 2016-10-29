import mailbox
import email
import sqlite3


class Importer:
    def __init__(self, mbox_file, db_file):
        self.email = mailbox.mbox(mbox_file)
        self.conn = sqlite3.connect(db_file)
        self.create_tables()

    # Create the structure for data storage. Indexes will be added by individual functions to (hopefully) optimize
    # performance.
    def create_tables(self):
        c = self.conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS messages(
              `message_key` INT,
              `from` TEXT,
              `to` TEXT,
              `subject` TEXT,
              `date` TEXT,
              `gmail_thread_id` INT,
              `gmail_labels` TEXT
             );
        ''')
        c.execute('''
             CREATE TABLE IF NOT EXISTS headers(
              `message_key` INT,
              `header` TEXT,
              `value` TEXT
             );
        ''')
        self.conn.commit()

    # Import message keys and important headers to create an index of messages.
    def import_messages(self):
        c = self.conn.cursor()

        count = 0
        for key, message in self.email.items():
            mail_from = message.get('From', message.get('from', message.get('FROM', ''))).decode('utf-8')
            mail_to = message.get('To', message.get('to', message.get('TO', ''))).decode('utf-8')
            mail_subject = message.get('Subject', message.get('subject', message.get('SUBJECT', ''))).decode('utf-8')
            mail_date_utc = self._parse_datetime(message)
            mail_gmail_id = message.get('X-GM-THRID', '')
            mail_gmail_labels = message.get('X-Gmail-Labels', '').decode('utf-8')

            c.execute('''INSERT INTO `messages` VALUES(?, ?, ?, ?, ?, ?, ?);''',
                      (key, mail_from, mail_to, mail_subject, mail_date_utc, mail_gmail_id, mail_gmail_labels))

            count += 1
            if count > 100000000:
                self.conn.commit()
                count = 0

        self.conn.commit()

    # Import all headers for each message.
    def import_message_headers(self):
        c = self.conn.cursor()

        count = 0
        for key, message in self.email.items():
            for header, value in message.items():
                c.execute('''INSERT INTO `headers` VALUES(?, ?, ?);''', (key, header, value.decode('utf-8')))
            count += 1
            if count > 100000000:
                self.conn.commit()
                count = 0

        self.conn.commit()

    # Find from datetime (may be in mailbox.Message.get_from() for Chat messages) and standard form and TZ (UTC).
    def _parse_datetime(self, message):
        mail_date = message.get('Date', message.get('date', message.get('DATE', ''))).decode('utf-8')
        if not mail_date:
            # The get_from() result always (so far as I have seen) has the date string in the last 30 characters
            mail_date = message.get_from().strip()[-30:]

        mail_date_utc = ''
        if mail_date:
            datetime_tuple = email.utils.parsedate_tz(mail_date)
            datetime = email.utils.mktime_tz(datetime_tuple)
            mail_date_utc = email.utils.formatdate(datetime)

        return mail_date_utc
