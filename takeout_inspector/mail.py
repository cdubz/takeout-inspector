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
import plotly.graph_objs as pgo
import re
import sqlite3
import wordcloud as wc

from .utils import *
from collections import OrderedDict
from datetime import datetime

__all__ = ['Import', 'Graph']


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
                           self._decode_header(address_info['real_name']), address_info['name']))

        c.execute('''CREATE INDEX id_date ON messages (`date` DESC)''')

        self.conn.commit()

    def _insert_recipients(self, c, key, message):
        """Parses contents of the To and CC headers for unique email addresses to be added to the one-row-per-address
        `recipients` table.
        """
        mail_all_to = message.get_all('To', [])
        for name, address in self._parse_addresses(mail_all_to):
            c.execute('''INSERT INTO recipients VALUES(?, ?, ?, ?);''',
                      (key, self._decode_header(name), address.decode('utf-8'), 'To'))
            self.query_count += 1

        mail_all_cc = message.get_all('CC', [])
        for name, address in self._parse_addresses(mail_all_cc):
            c.execute('''INSERT INTO recipients VALUES(?, ?, ?, ?);''',
                      (key, self._decode_header(name), address.decode('utf-8'), 'CC'))
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
        mail_from = self._decode_header(mail_from)

        mail_to = ''
        for idx, address in enumerate(self._parse_addresses(message.get_all('To', []))):
            mail_to += email.utils.formataddr(address) + ','  # Final ',' is removed at INSERT below.
        mail_to = self._decode_header(mail_to)

        mail_subject = self._decode_header(message.get('Subject', ''))
        mail_date_utc = self._get_message_date(message)
        mail_gmail_id = message.get('X-GM-THRID', '')
        mail_gmail_labels = self._decode_header(message.get('X-Gmail-Labels', ''))

        c.execute('''INSERT INTO messages VALUES(?, ?, ?, ?, ?, ?, ?);''',
                  (key, mail_from[:-1], mail_to[:-1], mail_subject, mail_date_utc, mail_gmail_id, mail_gmail_labels))
        self.query_count += 1

    def _decode_header(self, header):
        """Attempts to clean up a header:
            1. Removes newline and tab characters.
            2. Attempts to resole bad formatting.
            3. Decodes the header (if the header starts with "=?", for example).
            4. Recombines the decoded header in to a single unicode string.

        Badly formatted characters are ignored.

        TODO: Put email.header.decode_header() in a try and do something when it raises exceptions.
        """
        if header:
            header = ' '.join(header.split())  # Gets rid of newline and tab characters.
            header = header.replace('?==?', '?= =?')  # Encoded words must be separated by a space.
            header = email.header.decode_header(header)  # Handles UTF-8 and other encoded types.
            header = ' '.join([unicode(t[0], t[1] or 'utf-8', 'ignore') for t in header])  # Recombines all pieces.
        return header

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
        self.report = 'Mail'

        self.config = ConfigParser.ConfigParser()
        self.config.readfp(open('settings.defaults.cfg'))
        self.config.read(['settings.cfg'])

        self.conn = sqlite3.connect(self.config.get('mail', 'db_file'))

        self.owner_email = self.config.get('mail', 'owner')
        if self.config.getboolean('mail', 'anonymize'):  # If data is anonymized, get the fake address for the owner.
            c = self.conn.cursor()
            c.execute('''SELECT anon_address FROM address_key WHERE real_address = ?;''', (self.owner_email,))
            self.owner_email = c.fetchone()[0]

    def day_of_week(self):
        """Returns a graph showing email activity (sent/received) by day of the week.
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%w', `date`) AS dow,
          COUNT(CASE WHEN `from` LIKE ? THEN 1 ELSE NULL END) AS emails_sent,
          COUNT(CASE WHEN `from` NOT LIKE ? THEN 1 ELSE NULL END) AS emails_received
          FROM messages
          WHERE gmail_labels LIKE '%Chat%'
          GROUP BY dow
          ORDER BY dow ASC;''', ('%' + self.owner_email + '%', '%' + self.owner_email + '%'))

        sent = OrderedDict()
        sent_text = OrderedDict()
        received = OrderedDict()
        received_text = OrderedDict()
        for row in c.fetchall():
            dow = calendar.day_name[int(row[0]) - 1]  # sqlite strftime() uses 0 = SUNDAY.
            sent[dow] = row[1]
            received[dow] = row[2]
            sent_text[dow] = str(round(float(sent[dow]) / float(sent[dow] + received[dow]) * 100, 2)) + '%'
            received_text[dow] = str(round(float(received[dow]) / float(sent[dow] + received[dow]) * 100, 2)) + '%'

        sent_args = dict(
            x=sent.keys(),
            y=sent.values(),
            text=sent_text.values(),
            name='Emails sent',
            marker=dict(
                color=self.config.get('color', 'primary'),
            ),
        )

        received_args = dict(
            x=received.keys(),
            y=received.values(),
            text=received_text.values(),
            name='Emails received',
            marker=dict(
                color=self.config.get('color', 'secondary')
            ),
        )

        layout_args = plotly_default_layout_options()
        layout_args['barmode'] = 'stack'
        layout_args['title'] = 'Activity by Day of the Week'
        layout_args['xaxis']['title'] = 'Day of the week'
        layout_args['yaxis']['title'] = 'Number of emails'

        sent_trace = pgo.Bar(**sent_args)
        received_trace = pgo.Bar(**received_args)
        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[sent_trace, received_trace], layout=layout))

#    def label_network(self):
#        """Returns a network graph of non-default labels.
#
#        Add `import igraph as ig` at top to re-enable.
#        """
#        c = self.conn.cursor()
#
#        c.execute('''SELECT message_key, gmail_labels FROM messages
#                    WHERE gmail_labels != '' AND gmail_labels NOT LIKE '%Chat%';''')
#
#        default_labels = ['Important', 'Inbox', 'Sent', 'Spam', 'Starred', 'Trash', 'Unread']
#
#        g = ig.Graph()
#        vid = 0
#        gmail_labels = {}
#        for row in c.fetchall():
#            for label in row[1].split(','):
#                if label in default_labels:
#                    continue
#                elif label not in gmail_labels:
#                    g.add_vertex(label=label)
#                    gmail_labels[label] = vid
#                    vid += 1
#
#                if message_vid is None:
#                    g.add_vertex(label='Message ' + str(row[0]))
#                    message_vid = vid
#                    vid += 1
#
#                g.add_edge(gmail_labels[label], message_vid)
#
#            message_vid = None
#
#        labels = list(g.vs['label'])
#        N = len(labels)
#        E = [e.tuple for e in g.es]
#        layt = g.layout('fr')
#
#        Xn = [layt[k][0] for k in range(N)]
#        Yn = [layt[k][1] for k in range(N)]
#        Xe = []
#        Ye = []
#        for e in E:
#            Xe += [layt[e[0]][0], layt[e[1]][0], None]
#            Ye += [layt[e[0]][1], layt[e[1]][1], None]
#
#        trace1 = pgo.Scatter(
#            x=Xe,
#            y=Ye,
#            mode='lines',
#            line=pgo.Line(color='rgb(210,210,210)', width=1),
#            hoverinfo='none'
#        )
#        trace2 = pgo.Scatter(
#            x=Xn,
#            y=Yn,
#            mode='markers',
#            marker=pgo.Marker(
#                symbol='dot',
#                size=5,
#                color='#6959CD',
#                line=pgo.Line(color=self.config.get('color', 'primary'), width=0.5)
#            ),
#            text=labels,
#            hoverinfo='text'
#        )
#        data = pgo.Data([trace1, trace2])
#
#        hide_axis = dict(showline=False, zeroline=False, showgrid=False, showticklabels=False, title='')
#        layout_args = plotly_default_layout_options()
#        layout_args['title'] = 'Labels Network Graph'
#        layout_args['hovermode'] = 'closest'
#        layout_args['showlegend'] = False
#        layout_args['xaxis'] = pgo.XAxis(hide_axis)
#        layout_args['yaxis'] = pgo.YAxis(hide_axis)
#        layout = pgo.Layout(**layout_args)
#
#        return plotly_output(pgo.Figure(data=data, layout=layout))

    def label_usage(self):
        """Returns a pie chart showing usage information for labels.
        """
        c = self.conn.cursor()

        c.execute('''SELECT gmail_labels FROM messages
            WHERE gmail_labels != '' AND gmail_labels NOT LIKE '%Chat%';''')

        counts = {}
        for row in c.fetchall():
            for label in row[0].split(','):
                if label not in counts:
                    counts[label] = 0
                counts[label] += 1

        trace = pgo.Pie(
            labels=counts.keys(),
            values=counts.values(),
            marker=dict(
                colors=[
                    self.config.get('color', 'primary'),
                    self.config.get('color', 'secondary'),
                ]
            )
        )

        layout_args = plotly_default_layout_options()
        layout_args['title'] = 'Label Usage'
        del layout_args['xaxis']
        del layout_args['yaxis']

        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[trace], layout=layout))

    def subject_word_cloud(self):
        """Returns DIV for a word cloud of words used in email subjects saved to an image file
        """
        c = self.conn.cursor()

        c.execute('''SELECT subject FROM messages
            WHERE gmail_labels NOT LIKE '%Chat%' AND subject != '';''')

        words = {}
        for row in c.fetchall():
            subject = row[0]
            for prefix in ['Re:', 'Fwd:']:
                subject = subject.replace(prefix, '')

            subject = re.sub('[^a-zA-Z. ]', '', subject).strip().lower()  # Limits to alpha characters, dots and spaces.

            for word in subject.split(' '):
                if word:
                    word = word.rstrip('.')  # Remove periods from the end of words (sentences) only.
                    if word not in words:
                        words[word] = 0
                    words[word] += 1

        common_words = []
        for word in sorted(words, key=words.get, reverse=True):
            if words[word] >= 100:
                common_words.append([word, words[word]])

        cloud = wc.WordCloud(
            height=600,
            max_words=1000,
            width=600,
        )
        cloud.generate_from_frequencies(common_words)

        # TODO: Remove assumptions about location here - this should all be handled by report.py.
        file_path = 'resources/img/mail_subject_word_cloud.png'
        cloud.to_file(self.config.get('report', 'destination') + file_path)
        return {'html': '''
            <div id="mail_subject_word_cloud" style="text-align: center;">
                <h2>Subject Word Cloud</h2>
                <img src="{file_path}" alt="Mail Subject Word Cloud" />
            </div>
        '''.format(file_path=file_path)}

    def thread_durations(self):
        """Returns a pie chart showing grouped thread duration information. A "thread" must consist of more than one
        email.
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%s', MAX(`date`)) - strftime('%s', MIN(`date`)) AS duration,
            COUNT(message_key) AS message_count
            FROM messages
            WHERE gmail_labels NOT LIKE '%Chat%'
            GROUP BY gmail_thread_id
            HAVING message_count > 1;''')

        data = {'<= 10 min.': 0, '10 mins - 1 hr.': 0, '1 - 10 hrs.': 0,
                '10 - 24 hrs.': 0, '1 - 7 days': 0, '1 - 2 weeks': 0, 'more than 2 weeks': 0}
        for row in c.fetchall():
            if row[0] <= 600:
                data['<= 10 min.'] += 1
            elif row[0] <= 3600:
                data['10 mins - 1 hr.'] += 1
            elif row[0] <= 36000:
                data['1 - 10 hrs.'] += 1
            elif row[0] <= 86400:
                data['10 - 24 hrs.'] += 1
            elif row[0] <= 604800:
                data['1 - 7 days'] += 1
            elif row[0] <= 1209600:
                data['1 - 2 weeks'] += 1
            else:
                data['more than 2 weeks'] += 1

        trace = pgo.Pie(
            labels=data.keys(),
            values=data.values(),
            marker=dict(
                colors=[
                    self.config.get('color', 'primary'),
                    self.config.get('color', 'secondary'),
                ]
            )
        )

        layout_args = plotly_default_layout_options()
        layout_args['title'] = 'Thread Durations'
        del layout_args['xaxis']
        del layout_args['yaxis']

        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[trace], layout=layout))

    def thread_sizes(self):
        """Returns a graph showing thread size information. A "thread" must consist of more than one email.
        """
        c = self.conn.cursor()

        c.execute('''SELECT COUNT(message_key) AS message_count
            FROM messages
            WHERE gmail_labels NOT LIKE '%Chat%'
            GROUP BY gmail_thread_id
            HAVING message_count > 1;''')

        counts = {}
        for row in c.fetchall():
            if row[0] not in counts:
                counts[row[0]] = 0
            counts[row[0]] += 1

        data = dict(
            x=counts.keys(),
            y=counts.values(),
            name='Emails in thread',
            mode='lines+markers',
            marker=dict(
                color=self.config.get('color', 'primary'),
            ),
        )

        layout_args = plotly_default_layout_options()
        layout_args['title'] = 'Thread Sizes'
        layout_args['xaxis']['title'] = 'Number of messages'
        layout_args['yaxis']['title'] = 'Number of threads'
        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[pgo.Scatter(**data)], layout=layout))

    def time_of_day(self):
        """Returns a graph showing email activity (sent/received) by time of day.
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%H', `date`) AS hour,
          COUNT(CASE WHEN `from` LIKE ? THEN 1 ELSE NULL END) AS emails_sent,
          COUNT(CASE WHEN `from` NOT LIKE ? THEN 1 ELSE NULL END) AS emails_received
          FROM messages
          WHERE gmail_labels LIKE '%Chat%'
          GROUP BY hour
          ORDER BY hour ASC;''', ('%' + self.owner_email + '%', '%' + self.owner_email + '%'))

        sent = OrderedDict()
        sent_total = 0
        received = OrderedDict()
        received_total = 0
        for row in c.fetchall():
            sent_total += row[1]
            received_total += row[2]
            sent[row[0]] = row[1]
            received[row[0]] = row[2]

        sent_args = dict(
            x=sent.keys(),
            y=sent.values(),
            name='Emails sent',
            marker=dict(
                color=self.config.get('color', 'primary'),
            ),
        )

        received_args = dict(
            x=received.keys(),
            y=received.values(),
            name='Emails received',
            marker=dict(
                color=self.config.get('color', 'secondary')
            ),
        )

        layout_args = plotly_default_layout_options()
        layout_args['title'] = 'Activity by Hour of the Day (UTC)'
        layout_args['xaxis']['title'] = 'Hour of the day (UTC)'
        layout_args['yaxis']['title'] = 'Number of emails'

        sent_trace = pgo.Scatter(**sent_args)
        received_trace = pgo.Scatter(**received_args)
        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[sent_trace, received_trace], layout=layout))

    def top_recipients(self, limit=10):
        """Returns a bar graph showing the top `limit` number of recipients of emails sent.

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

        layout = plotly_default_layout_options()
        layout['margin']['l'] = longest_address * self.config.getfloat('font', 'size')/1.55
        layout['margin'] = pgo.Margin(**layout['margin'])
        layout['title'] = 'Top ' + str(limit) + ' Recipients'
        layout['xaxis']['title'] = 'Emails sent to'
        layout['yaxis']['title'] = 'Recipient address'

        return plotly_output(pgo.Figure(data=[pgo.Bar(**data)], layout=pgo.Layout(**layout)))

    def top_senders(self, limit=10):
        """Returns a bar graph showing the top `limit` number of senders of emails received.

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

        layout = plotly_default_layout_options()
        layout['margin']['l'] = longest_address * self.config.getfloat('font', 'size')/1.55
        layout['margin'] = pgo.Margin(**layout['margin'])
        layout['title'] = 'Top ' + str(limit) + ' Senders'
        layout['xaxis']['title'] = 'Emails received from'
        layout['yaxis']['title'] = 'Sender address'

        return plotly_output(pgo.Figure(data=[pgo.Bar(**data)], layout=pgo.Layout(**layout)))
