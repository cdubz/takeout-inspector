import mailbox
import sqlite3


class Importer:
    def __init__(self, mbox_file, db_file):
        self.email = mailbox.mbox(mbox_file)
        self.conn = sqlite3.connect(db_file)

    def import_messages(self):
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
        self.conn.commit()

        count = 0
        for key, message in self.email.items():
            mail_from = message.get('From', message.get('from', message.get('FROM', ''))).decode('utf-8')
            mail_to = message.get('To', message.get('to', message.get('TO', ''))).decode('utf-8')
            mail_subject = message.get('Subject', message.get('subject', message.get('SUBJECT', ''))).decode('utf-8')
            mail_date = message.get('Date', message.get('date', message.get('DATE', ''))).decode('utf-8')
            mail_gmail_id = message.get('X-GM-THRID', '')
            mail_gmail_labels = message.get('X-Gmail-Labels', '').decode('utf-8')

            c.execute('''INSERT INTO `messages` VALUES(?, ?, ?, ?, ?, ?, ?);''',
                      (key, mail_from, mail_to, mail_subject, mail_date, mail_gmail_id, mail_gmail_labels))

            count += 1
            if count > 100000000:
                self.conn.commit()
                count = 0

        self.conn.commit()

    def import_message_headers(self):
        c = self.conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS headers(
              `message_key` INT,
              `header` TEXT,
              `value` TEXT
             );
         ''')
        self.conn.commit()

        count = 0
        for key, message in self.email.items():
            for header, value in message.items():
                c.execute('''INSERT INTO `headers` VALUES(?, ?, ?);''', (key, header, value.decode('utf-8')))
            count += 1
            if count > 100000000:
                self.conn.commit()
                count = 0

        self.conn.commit()
