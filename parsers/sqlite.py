import mailbox
import sqlite3


class Parser:
    def __init__(self, mbox_file, db_file):
        self.email = mailbox.mbox(mbox_file)
        self.conn = sqlite3.connect(db_file)

    def import_message_headers(self):
        c = self.conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS headers(
              message_key INT,
              header TEXT,
              value TEXT
             );
         ''')
        self.conn.commit()

        count = 0
        for key, message in self.email.items():
            for header, value in message.items():
                c.execute('''INSERT INTO headers VALUES(?, ?, ?);''', (key, header, value.decode('utf-8')))
            count += 1
            if count > 100000000:
                self.conn.commit()
                count = 0

        self.conn.commit()

    def create_headers_occurrences_table(self):
        c = self.conn.cursor()
        c.execute('''
                    CREATE TABLE IF NOT EXISTS headers_occurrences(
                      header VARCHAR(255) PRIMARY KEY,
                      occurrences INT
                     );
                 ''')

        headers = {}
        for message in self.email:
            for header in message.keys():
                if header not in headers:
                    headers[header] = 1
                else:
                    headers[header] += 1

        for header, occurrences in headers.iteritems():
            c.execute('''INSERT INTO headers_occurrences VALUES(?, ?);''', (header, occurrences))

        self.conn.commit()
