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
import calendar
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

        self.address_key = {}

        self.anonymize = config.getboolean('mail', 'anonymize')
        if self.anonymize:
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
            for address, address_info in self.address_key.iteritems():
                c.execute('''INSERT INTO address_key VALUES(?, ?, ?, ?);''',
                          (address_info['real_address'].decode('utf-8'), address_info['address'],
                           address_info['real_name'].decode('utf-8'), address_info['name']))

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
        """Turns a name and address in to an anonymized [address, anon_address, name, anon_name] dict and returns the
        result.

        Additionally, self.domain_key is maintained so the domain can be anonymized consistently in order to allow for
        potential querying of the database with domain-based grouping.
        """
        domain = address.split('@', 1)[1]
        if domain not in self.domain_key:
            self.domain_key[domain] = 'domain' + str(len(self.domain_key)) + '.tld'

        anon_name = names.get_full_name()

        return {
            'real_address': address,
            'address': anon_name.replace(' ', '-').lower() + '@' + self.domain_key[domain],
            'real_name': name,
            'name': anon_name
        }

    def _parse_addresses(self, addresses, unique=True):
        """Turns a list of address strings (e.g. from email.Message.get_all()) in to a list of formatted [name, address]
        tuples. Formatting does the following:
            1) Removes XMPP Resourceparts (https://xmpp.org/rfcs/rfc6122.html).
            2) Removes periods from the local part for @gmail.com addresses.
            3) Converts the full address to lower case.

        Also adds address information to self.address_key with address as the key (if it does not already exist). As a
        side effect, this method will only use the first name it encounters for any particular email. Not ideal, but
        also not a big deal as long as the actual unique identifer (the email) is preserved.

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

            if address not in self.address_key:
                if self.anonymize:
                    self.address_key[address] = self._anonymize_address(address, name)
                else:
                    self.address_key[address] = {
                        'address': address,
                        'name': name,
                    }

            addresses[idx] = [self.address_key[address]['name'], self.address_key[address]['address']]

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

        self.owner_email = self.config.get('mail', 'owner')
        if self.config.getboolean('mail', 'anonymize'):  # If data is anonymized, get the fake address for the owner.
            c = self.conn.cursor()
            c.execute('''SELECT anon_address FROM address_key WHERE real_address = ?;''', (self.owner_email,))
            self.owner_email = c.fetchone()[0]

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
                '<body style="max-width: 800px; margin: 0 auto;">\n',
                '<h1 style="text-align: center;">Top 10 Lists</h1>\n',
                self.top_recipients(top_recipients_limit) + '\n',
                self.top_senders(top_senders_limit) + '\n',
                '<h1 style="text-align: center;">Chat Statistics</h1>\n',
                self.chat_vs_email() + '\n',
                self.chat_vs_email(cumulative=True) + '\n',
                self.chat_clients() + '\n',
                self.chat_times() + '\n',
                self.chat_days() + '\n',
                self.chat_durations() + '\n',
                self.chat_thread_sizes() + '\n',
                self.chat_top_chatters() + '\n',
                '</body>\n',
                '</html>',
            ]))

    def chat_clients(self):
        """Returns a pie chart showing distribution of services/client used (based on known resourceparts). This likely
        not particularly accurate!
        """
        c = self.conn.cursor()

        c.execute('''SELECT value FROM headers WHERE header = 'To' AND value NOT LIKE '%,%';''')

        clients = {'android': 0, 'Adium': 0, 'BlackBerry': 0, 'Festoon': 0, 'fire': 0,
                    'Gush': 0, 'Gaim': 0, 'gmail': 0, 'Meebo': 0, 'Miranda': 0,
                    'Psi': 0, 'iChat': 0, 'iGoogle': 0, 'IM+': 0, 'Talk': 0,
                    'Trillian': 0, 'Unknown': 0
                   }
        for row in c.fetchall():
            try:
                domain = row[0].split('@', 1)[1]
                resource_part = domain.split('/', 1)[1]
            except IndexError:  # Throws when the address does not have an @ or a / in the string.
                continue

            unknown = True
            for client in clients:
                if client in resource_part:
                    clients[client] += 1
                    unknown = False

            if unknown:
                clients['Unknown'] += 1

        for client in clients.keys():
            if clients[client] is 0:
                del clients[client]

        trace = go.Pie(
            labels=clients.keys(),
            values=clients.values(),
            marker=dict(
                colors=[
                    self.config.get('color', 'primary'),
                    self.config.get('color', 'secondary'),
                ]
            )
        )

        layout_args = self._default_layout_options()
        layout_args['title'] = 'Chat Clients'
        del layout_args['xaxis']
        del layout_args['yaxis']

        layout = go.Layout(**layout_args)

        return py.plot(
            go.Figure(data=[trace], layout=layout),
            output_type='div',
            include_plotlyjs=False,
        )

    def chat_days(self):
        """Returns a stacked bar chart showing percentage of chats and emails on each day of the week.
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%w', `date`) AS dow,
            COUNT(CASE WHEN gmail_labels LIKE '%Chat%' THEN 1 ELSE NULL END) AS chat_messages,
            COUNT(CASE WHEN gmail_labels NOT LIKE '%Chat%' THEN 1 ELSE NULL END) AS email_messages
            FROM messages
            WHERE dow NOTNULL
            GROUP BY dow;''')

        chat_percentages = OrderedDict()
        chat_messages = OrderedDict()
        email_percentages = OrderedDict()
        email_messages = OrderedDict()
        for row in c.fetchall():
            dow = calendar.day_name[int(row[0]) - 1]  # sqlite strftime() uses 0 = SUNDAY.
            chat_percentages[dow] = str(round(float(row[1]) / sum([row[1], row[2]]) * 100, 2)) + '%'
            email_percentages[dow] = str(round(float(row[2]) / sum([row[1], row[2]]) * 100, 2)) + '%'
            chat_messages[dow] = row[1]
            email_messages[dow] = row[2]

        chats_trace = go.Bar(
            x=chat_messages.keys(),
            y=chat_messages.values(),
            text=chat_percentages.values(),
            name='Chat messages',
            marker=dict(
                color=self.config.get('color', 'primary'),
            ),
        )
        emails_trace = go.Bar(
            x=email_messages.keys(),
            y=email_messages.values(),
            text=email_percentages.values(),
            name='Email messages',
            marker=dict(
                color=self.config.get('color', 'secondary'),
            ),
        )

        layout = self._default_layout_options()
        layout['barmode'] = 'stack'
        layout['margin'] = go.Margin(**layout['margin'])
        layout['title'] = 'Chat (vs. Email) Days'
        layout['xaxis']['title'] = 'Day of the week'
        layout['yaxis']['title'] = 'Messages exchanged'

        return py.plot(
            go.Figure(data=[chats_trace, emails_trace], layout=go.Layout(**layout)),
            output_type='div',
            include_plotlyjs=False,
        )

    def chat_durations(self):
        """Returns a plotly pie chart showing grouped chat duration information.
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%s', MAX(`date`)) - strftime('%s', MIN(`date`)) AS duration
            FROM messages
            WHERE gmail_labels LIKE '%Chat%'
            GROUP BY gmail_thread_id
            HAVING duration > 0;''')

        data = {'<= 1 min.': 0, '1 - 10 mins.': 0,
                '10 - 30 mins.': 0, '30 mins. - 1 hr.': 0,
                '> 1 hr.': 0}
        for row in c.fetchall():
            if row[0] <= 60:
                data['<= 1 min.'] += 1
            elif row[0] <= 600:
                data['1 - 10 mins.'] += 1
            elif row[0] <= 1800:
                data['10 - 30 mins.'] += 1
            elif row[0] <= 3600:
                data['30 mins. - 1 hr.'] += 1
            else:
                data['> 1 hr.'] += 1

        trace = go.Pie(
            labels=data.keys(),
            values=data.values(),
            marker=dict(
                colors=[
                    self.config.get('color', 'primary'),
                    self.config.get('color', 'secondary'),
                ]
            )
        )

        layout_args = self._default_layout_options()
        layout_args['title'] = 'Chat Durations'
        del layout_args['xaxis']
        del layout_args['yaxis']

        layout = go.Layout(**layout_args)

        return py.plot(
            go.Figure(data=[trace], layout=layout),
            output_type='div',
            include_plotlyjs=False,
        )

    def chat_thread_sizes(self):
        """Returns a plotly scatter/bubble graph showing the sizes (by message count) of chat thread over time.
        """
        c = self.conn.cursor()

        c.execute('''SELECT gmail_thread_id,
            strftime('%Y-%m-%d', `date`) AS thread_date,
            COUNT(message_key) as thread_size,
            GROUP_CONCAT(DISTINCT `from`) AS participants
            FROM messages
            WHERE gmail_labels LIKE '%Chat%'
            GROUP BY gmail_thread_id;''')

        messages = []
        marker_sizes = []
        dates = []
        descriptions = []
        for row in c.fetchall():
            messages.append(row[2])
            marker_sizes.append(max(10, row[2]/5))
            dates.append(row[1])
            descriptions.append('Messages: ' + str(row[2]) +
                                '<br>Date: ' + str(row[1]) +
                                '<br>Participants:<br> - ' + str(row[3]).replace(',', '<br> - ')
                                )

        trace = go.Scatter(
            x=dates,
            y=messages,
            mode='markers',
            marker=dict(
                size=marker_sizes,
            ),
            text=descriptions
        )

        layout_args = self._default_layout_options()
        layout_args['title'] = 'Chat Thread Sizes'
        layout_args['hovermode'] = 'closest'
        layout_args['height'] = 800
        layout_args['margin'] = go.Margin(**layout_args['margin'])
        layout_args['xaxis']['title'] = 'Date'
        layout_args['yaxis']['title'] = 'Messages in thread'
        layout = go.Layout(**layout_args)

        return py.plot(
            go.Figure(data=[trace], layout=layout),
            output_type='div',
            include_plotlyjs=False,
        )

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

    def chat_top_chatters(self, limit=10):
        """Returns a plotly bar graph showing top chat senders with an email comparison.

        Keyword arguments:
            limit -- How many chat senders to return.
        """
        c = self.conn.cursor()

        c.execute('''SELECT `from`,
            COUNT(CASE WHEN gmail_labels LIKE '%Chat%' THEN 1 ELSE NULL END) AS chat_messages,
            COUNT(CASE WHEN gmail_labels NOT LIKE '%Chat%' THEN 1 ELSE NULL END) AS email_messages
            FROM messages
            WHERE `from` NOT LIKE ?
            GROUP BY `from`
            ORDER BY chat_messages DESC
            LIMIT ?;''', ('%' + self.owner_email + '%', limit,))

        chats = OrderedDict()
        emails = OrderedDict()
        longest_address = 0
        for row in c.fetchall():
            chats[row[0]] = row[1]
            emails[row[0]] = row[2]
            longest_address = max(longest_address, len(row[0]))

        chats_trace = go.Bar(
            x=chats.keys(),
            y=chats.values(),
            name='Chat messages',
            marker=dict(
                color=self.config.get('color', 'primary'),
            ),
        )
        emails_trace = go.Bar(
            x=emails.keys(),
            y=emails.values(),
            name='Email messages',
            marker=dict(
                color=self.config.get('color', 'secondary'),
            ),
        )

        layout = self._default_layout_options()
        layout['barmode'] = 'grouped'
        layout['height'] = longest_address * 15
        layout['margin']['b'] = longest_address * self.config.getfloat('font', 'size') / 2
        layout['margin'] = go.Margin(**layout['margin'])
        layout['title'] = 'Top ' + str(limit) + ' Chatters'
        layout['xaxis']['title'] = 'Sender address'
        layout['yaxis']['title'] = 'Messages received from'

        return py.plot(
            go.Figure(data=[chats_trace, emails_trace], layout=go.Layout(**layout)),
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
