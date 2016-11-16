"""takeout_inspector/mail.py

Defines classes and methods used to parse a Google Takeout mbox file, import data in to an sqlite database and generate
graphs.

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
import ConfigParser
import email
import mailbox
import names
import plotly.offline as py
import plotly.graph_objs as go
import sqlite3

from collections import OrderedDict
from datetime import datetime


class Import:
    """Parses and imports Google Takeout mbox file data in to sqlite.
    """
    def __init__(self):
        config = ConfigParser.ConfigParser()
        config.readfp(open('settings.defaults.cfg'))
        config.read(['settings.cfg'])

        self.email = mailbox.mbox(config.get('mail', 'mbox_file'))
        self.conn = sqlite3.connect(config.get('mail', 'db_file'))

        self.anonymize = bool(config.get('mail', 'anonymize'))
        if self.anonymize:
            self.address_key = {}
            self.domain_key = {}

        self._create_tables()
        self.query_count = 0

    def _create_tables(self):
        """Creates the required tables for message data storage. Indexes will be added after data import.
        """
        c = self.conn.cursor()

        c.execute('''
            CREATE TABLE IF NOT EXISTS messages(
              message_key INT PRIMARY KEY,
              `from` TEXT,
              `to` TEXT,
              subject TEXT,
              `date` DATETIME,
              gmail_thread_id INT,
              gmail_labels TEXT
             );
        ''')
        c.execute('''
             CREATE TABLE IF NOT EXISTS headers(
              message_key INT,
              header TEXT,
              value TEXT,
              FOREIGN KEY(message_key) REFERENCES messages(message_key)
             );
        ''')
        c.execute('''
             CREATE TABLE IF NOT EXISTS recipients(
              message_key INT,
              name TEXT,
              address TEXT,
              header TEXT,
              FOREIGN KEY(message_key) REFERENCES messages(message_key)
             );
        ''')

        if self.anonymize:
            c.execute('''
                 CREATE TABLE IF NOT EXISTS address_key(
                  real_address TEXT,
                  anon_address TEXT,
                  real_name TEXT,
                  anon_name TEXT
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

            if self.query_count > 1000000:
                self.conn.commit()
                self.query_count = 0

        if self.anonymize:
            for real_address, anon_info in self.address_key.iteritems():
                c.execute('''INSERT INTO address_key VALUES(?, ?, ?, ?);''',
                          (anon_info['address'].decode('utf-8'), anon_info['anon_address'],
                           anon_info['name'].decode('utf-8'), anon_info['anon_name']))

        c.execute('''CREATE INDEX id_date ON messages (`date` DESC)''')

        self.conn.commit()

    def _insert_recipients(self, c, key, message):
        """Parses contents of the To and CC headers for unique email addresses to be added to the one-row-per-address
        `recipients` table.
        """
        mail_all_to = message.get_all('To', [])
        for name, address in self._parse_addresses(mail_all_to):
            c.execute('''INSERT INTO recipients VALUES(?, ?, ?, ?);''',
                      (key, name.decode('utf-8'), address.decode('utf-8'), 'To'))
            self.query_count += 1

        mail_all_cc = message.get_all('CC', [])
        for name, address in self._parse_addresses(mail_all_cc):
            c.execute('''INSERT INTO recipients VALUES(?, ?, ?, ?);''',
                      (key, name.decode('utf-8'), address.decode('utf-8'), 'CC'))
            self.query_count += 1

    def _insert_headers(self, c, key, message):
        """Adds all headers to `headers`.

        WARNING: Data in this table does _not_ currently respect the self.anonymize setting. This is meant to be a raw
        record of all headers.
        """
        for header, value in message.items():
            c.execute('''INSERT INTO headers VALUES(?, ?, ?);''', (key, header, value.decode('utf-8')))
            self.query_count += 1

    def _insert_messages(self, c, key, message):
        """Creates a basic index of important message data in `messages`.
        """
        mail_from = ''
        for idx, address in enumerate(self._parse_addresses(message.get_all('From', []))):
            mail_from += email.utils.formataddr(address) + ','  # Final ',' is removed at INSERT below.

        mail_to = ''
        for idx, address in enumerate(self._parse_addresses(message.get_all('To', []))):
            mail_to += email.utils.formataddr(address) + ','  # Final ',' is removed at INSERT below.

        mail_subject = message.get('Subject', '')
        mail_date_utc = self._get_message_date(message)
        mail_gmail_id = message.get('X-GM-THRID', '')
        mail_gmail_labels = message.get('X-Gmail-Labels', '')

        c.execute('''INSERT INTO messages VALUES(?, ?, ?, ?, ?, ?, ?);''',
                  (key, mail_from[:-1].decode('utf-8'), mail_to[:-1].decode('utf-8'), mail_subject.decode('utf-8'),
                   mail_date_utc, mail_gmail_id, mail_gmail_labels.decode('utf-8')))
        self.query_count += 1

    def _anonymize_address(self, address, name):
        """Turns a name and address in to an anonymized [address, anon_address, name, anon_name] dict and adds it to
        self.address_key with address as the key (if it does not already exist). Returns the full dict.

        Additionally, self.domain_key is maintained so the domain can be anonymized consistently in order to allow for
        potential querying of the database with domain-based grouping.

        As a side effect, this method will only use the first name it encounters for any particular email. Not ideal,
        but also not a big deal as long as the actual unique identifer (the email) is preserved.
        """
        domain = address.split('@', 1)[1]
        if domain not in self.domain_key:
            self.domain_key[domain] = 'domain' + str(len(self.domain_key)) + '.tld'

        if address not in self.address_key:
            anon_name = names.get_full_name()
            self.address_key[address] = {
                'address': address,
                'anon_address': anon_name.replace(' ', '-').lower() + '@' + self.domain_key[domain],
                'name': name,
                'anon_name': anon_name
            }
        return self.address_key[address]

    def _parse_addresses(self, addresses, unique=True):
        """Turns a list of address strings (e.g. from email.Message.get_all()) in to a list of formatted [name, address]
        tuples. Formatting does the following:
            1) Removes XMPP Resourceparts (https://xmpp.org/rfcs/rfc6122.html).
            2) Removes periods from the local part for @gmail.com addresses.
            3) Converts the full address to lower case.

        Keyword arguments:
            unique -- Produces a list of unique entries by email address.
        """
        addresses = email.utils.getaddresses(addresses)
        if unique:
            addresses = list(set(addresses))

        for idx, address in enumerate(addresses):
            name = address[0]
            try:
                [local_part, domain] = address[1].split('@', 1)
                domain = domain.split('/', 1)[0].lower()  # Removes Resourcepart and normalizes case.
                local_part = local_part.lower()  # Normalizes to all lower case.
                if domain in ['gmail.com']:  # Removes dots for services that disregard them.
                    local_part = local_part.replace('.', '')
                address = local_part + '@' + domain
            except ValueError:  # Throws when the address does not have an @ anywhere in the string.
                address = address[1] + '@domain-not-found.tld'

            if self.anonymize:
                anonymized_address = self._anonymize_address(address, name)
                name = anonymized_address['anon_name']
                address = anonymized_address['anon_address']

            addresses[idx] = [name, address]

        return addresses

    def _get_message_date(self, message):
        """Finds date and time information for `message` and converts it to ISO-8601 format and UTC timezone.
        """
        mail_date = message.get('Date', '').decode('utf-8')
        if not mail_date:
            """The get_from() result always (so far as I have seen!) has the date string in the last 30 characters"""
            mail_date = message.get_from().strip()[-30:]

        datetime_tuple = email.utils.parsedate_tz(mail_date)
        if datetime_tuple:
            unix_time = email.utils.mktime_tz(datetime_tuple)
            mail_date_iso8601 = datetime.utcfromtimestamp(unix_time).isoformat(' ')
        else:
            mail_date_iso8601 = ''

        return mail_date_iso8601


class Graph:
    """Creates offline plotly graphs using imported data from sqlite.
    """
    def __init__(self):
        self.config = ConfigParser.ConfigParser()
        self.config.readfp(open('settings.defaults.cfg'))
        self.config.read(['settings.cfg'])

        self.conn = sqlite3.connect(self.config.get('mail', 'db_file'))

    def all_graphs(self, top_recipients_limit=10, top_senders_limit=10):
        """Creates an HTML file containing all available graphs.

        Keyword arguments:
            top_recipients_limit -- Number of top recipients to graph.
            top_senders_limit -- Number of top senders to graph.
        """
        with open('all_graphs_mail.html', 'w') as f:
            f.write(''.join([
                '<!DOCTYPE HTML>\n',
                '<html>\n',
                '<head>\n',
                '\t<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />\n',
                '\t<title>Mail - All Graphs | Takeout Inspector</title>\n',
                '\t<script src="https://cdn.plot.ly/plotly-latest.min.js"></script>\n',
                '</head>\n',
                '<body>\n',
                self.top_recipients(top_recipients_limit) + '\n',
                self.top_senders(top_senders_limit) + '\n',
                self.chat_vs_email() + '\n',
                self.chat_vs_email(cumulative=True) + '\n',
                self.chat_times() + '\n',
                '</body>\n',
                '</html>',
            ]))

    def chat_times(self):
        """Returns a plotly graph showing chat habits by hour of the day (UTC).
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%H', `date`) AS hour, COUNT(message_key) AS chat_messages
            FROM messages
            WHERE gmail_labels LIKE '%Chat%'
            GROUP BY hour
            ORDER BY hour ASC;''')

        data = OrderedDict()
        for row in c.fetchall():
            data[row[0]] = row[1]

        total_messages = sum(data.values())
        percentages = OrderedDict()
        for hour in data.keys():
            percentages[hour] = str(round(float(data[hour])/float(total_messages) * 100, 2)) + '%'

        data_args = dict(
            x=data.keys(),
            y=data.values(),
            text=percentages.values(),
            name='Chat messages',
            marker=dict(
                color=self.config.get('color', 'primary')
            ),
            fill='tozeroy',
        )

        layout_args = self._default_layout_options()
        layout_args['title'] = 'Chat Times (UTC)'
        layout_args['xaxis']['title'] = 'Hour of day (UTC)'
        layout_args['yaxis']['title'] = 'Chat messages'

        trace = go.Scatter(**data_args)
        layout = go.Layout(**layout_args)

        return py.plot(
            go.Figure(data=[trace], layout=layout),
            output_type='div',
            include_plotlyjs=False,
        )

    def chat_vs_email(self, cumulative=False):
        """Returns a plotly graph showing chat vs. email usage over time (by year and month).

        Keyword arguments:
            cumulative -- Whether ot not to display cumulative data for each month.
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%Y-%m', `date`) as period,
          COUNT(CASE WHEN gmail_labels LIKE '%Chat%' THEN 1 ELSE NULL END) AS chat_messages,
          COUNT(CASE WHEN gmail_labels NOT LIKE '%Chat%' THEN 1 ELSE NULL END) AS email_messages
          FROM messages
          GROUP BY period
          ORDER BY period ASC;''')

        chat_data = OrderedDict()
        chat_total = 0
        email_data = OrderedDict()
        email_total = 0
        for row in c.fetchall():
            chat_total += row[1]
            email_total += row[2]
            if cumulative:
                chat_data[row[0]] = chat_total
                email_data[row[0]] = email_total
            else:
                chat_data[row[0]] = row[1]
                email_data[row[0]] = row[2]

        chat_args = dict(
            x=chat_data.keys(),
            y=chat_data.values(),
            name='Chats',
            marker=dict(
                color=self.config.get('color', 'primary'),
            ),
        )

        email_args = dict(
            x=email_data.keys(),
            y=email_data.values(),
            name='Emails',
            marker=dict(
                color=self.config.get('color', 'secondary')
            ),
        )

        layout_args = self._default_layout_options()
        layout_args['title'] = 'Chat vs. Email Usage'
        layout_args['xaxis']['title'] = 'Year and month'
        layout_args['yaxis']['title'] = 'Number of messages'

        if cumulative:
            layout_args['title'] += ' (Cumulative)'
            chat_args['fill'] = 'tonexty'
            email_args['fill'] = 'tozeroy'

        chat_trace = go.Scatter(**chat_args)
        email_trace = go.Scatter(**email_args)
        layout = go.Layout(**layout_args)

        return py.plot(
            go.Figure(data=[chat_trace, email_trace], layout=layout),
            output_type='div',
            include_plotlyjs=False,
        )

    def top_recipients(self, limit=20):
        """Returns a plotly bar graph <div> showing the top `limit` number of recipients of emails sent.

        Keyword arguments:
            limit -- Number of recipients to include.
        """
        c = self.conn.cursor()

        c.execute('''SELECT address, COUNT(r.message_key) AS message_count
            FROM recipients AS r
            LEFT JOIN messages AS m ON(m.message_key = r.message_key)
            WHERE m.gmail_labels LIKE '%Sent%'
            GROUP BY address
            ORDER BY message_count DESC
            LIMIT ?''', (limit,))

        addresses = OrderedDict()
        longest_address = 0
        for row in c.fetchall():
            addresses[row[0]] = row[1]
            longest_address = max(longest_address, len(row[0]))

        data = dict(
            x=addresses.values(),
            y=addresses.keys(),
            marker=dict(
                color=self.config.get('color', 'primary_light'),
                line=dict(
                    color=self.config.get('color', 'primary'),
                    width=1,
                ),
            ),
            orientation='h',
        )

        layout = self._default_layout_options()
        layout['margin']['l'] = longest_address * self.config.getfloat('font', 'size')/1.55
        layout['margin'] = go.Margin(**layout['margin'])
        layout['title'] = 'Top ' + str(limit) + ' Recipients'
        layout['xaxis']['title'] = 'Emails sent to'
        layout['yaxis']['title'] = 'Recipient address'

        return py.plot(
            go.Figure(data=[go.Bar(**data)], layout=go.Layout(**layout)),
            output_type='div',
            include_plotlyjs=False,
        )

    def top_senders(self, limit=20):
        """Returns a plotly bar graph <div> showing the top `limit` number of senders of emails received.

        Keyword arguments:
            limit -- Number of senders to include.
        """
        c = self.conn.cursor()

        c.execute('''SELECT `from`, COUNT(message_key) AS message_count
            FROM messages
            WHERE gmail_labels NOT LIKE '%Sent%'
                AND gmail_labels NOT LIKE '%Chat%'
            GROUP BY `from`
            ORDER BY message_count DESC
            LIMIT ?''', (limit,))

        addresses = OrderedDict()
        longest_address = 0
        for row in c.fetchall():
            addresses[row[0]] = row[1]
            longest_address = max(longest_address, len(row[0]))

        data = dict(
            x=addresses.values(),
            y=addresses.keys(),
            marker=dict(
                color=self.config.get('color', 'primary_light'),
                line=dict(
                    color=self.config.get('color', 'primary'),
                    width=1,
                ),
            ),
            orientation='h',
        )

        layout = self._default_layout_options()
        layout['margin']['l'] = longest_address * self.config.getfloat('font', 'size')/1.55
        layout['margin'] = go.Margin(**layout['margin'])
        layout['title'] = 'Top ' + str(limit) + ' Senders'
        layout['xaxis']['title'] = 'Emails received from'
        layout['yaxis']['title'] = 'Sender address'

        return py.plot(
            go.Figure(data=[go.Bar(**data)], layout=go.Layout(**layout)),
            output_type='div',
            include_plotlyjs=False,
        )

    def _default_layout_options(self):
        """Prepares default layout options for all graphs.
        """
        return dict(
            font=dict(
                color=self.config.get('color', 'text'),
                family=self.config.get('font', 'family'),
                size=self.config.get('font', 'size'),
            ),
            margin=dict(
                b=50,
                t=50,
            ),
            xaxis=dict(
                titlefont=dict(
                    color=self.config.get('color', 'text_lighter'),
                )
            ),
            yaxis=dict(
                titlefont=dict(
                    color=self.config.get('color', 'text_lighter'),
                ),
            ),
        )
