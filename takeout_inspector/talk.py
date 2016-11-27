"""takeout_inspector/talk.py

Defines classes and methods used to generate graphs for Google Talk data (based on a Google Mail takeout file).

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
import plotly.graph_objs as pgo
import sqlite3

from .utils import *
from collections import OrderedDict

__all__ = ['Import', 'Graph']


class Import:
    """Print a message noting that Talk relies on Mail's import data.
    """
    def __init__(self):
        print "Talk's Import class is empty. Run Mail's Import class " \
              "first as Talk data is stored in the Mail export file."


class Graph:
    """Creates offline plotly graphs using imported data from sqlite.
    """
    def __init__(self):
        self.report = 'Talk'

        self.config = ConfigParser.ConfigParser()
        self.config.readfp(open('settings.defaults.cfg'))
        self.config.read(['settings.cfg'])

        self.conn = sqlite3.connect(self.config.get('mail', 'db_file'))

        self.owner_email = self.config.get('mail', 'owner')
        if self.config.getboolean('mail', 'anonymize'):  # If data is anonymized, get the fake address for the owner.
            c = self.conn.cursor()
            c.execute('''SELECT anon_address FROM address_key WHERE real_address = ?;''', (self.owner_email,))
            self.owner_email = c.fetchone()[0]

    def talk_clients(self):
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

        trace = pgo.Pie(
            labels=clients.keys(),
            values=clients.values(),
            marker=dict(
                colors=[
                    self.config.get('color', 'primary'),
                    self.config.get('color', 'secondary'),
                ]
            )
        )

        layout_args = plotly_default_layout_options()
        layout_args['title'] = 'Chat Clients'
        del layout_args['xaxis']
        del layout_args['yaxis']

        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[trace], layout=layout))

    def talk_days(self):
        """Returns a stacked bar chart showing percentage of chats and emails on each day of the week.
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%w', `date`) AS dow,
            COUNT(CASE WHEN gmail_labels LIKE '%Chat%' THEN 1 ELSE NULL END) AS talk_messages,
            COUNT(CASE WHEN gmail_labels NOT LIKE '%Chat%' THEN 1 ELSE NULL END) AS email_messages
            FROM messages
            WHERE dow NOTNULL
            GROUP BY dow;''')

        talk_percentages = OrderedDict()
        talk_messages = OrderedDict()
        email_percentages = OrderedDict()
        email_messages = OrderedDict()
        for row in c.fetchall():
            dow = calendar.day_name[int(row[0]) - 1]  # sqlite strftime() uses 0 = SUNDAY.
            talk_percentages[dow] = str(round(float(row[1]) / sum([row[1], row[2]]) * 100, 2)) + '%'
            email_percentages[dow] = str(round(float(row[2]) / sum([row[1], row[2]]) * 100, 2)) + '%'
            talk_messages[dow] = row[1]
            email_messages[dow] = row[2]

        chats_trace = pgo.Bar(
            x=talk_messages.keys(),
            y=talk_messages.values(),
            text=talk_percentages.values(),
            name='Chat messages',
            marker=dict(
                color=self.config.get('color', 'primary'),
            ),
        )
        emails_trace = pgo.Bar(
            x=email_messages.keys(),
            y=email_messages.values(),
            text=email_percentages.values(),
            name='Email messages',
            marker=dict(
                color=self.config.get('color', 'secondary'),
            ),
        )

        layout = plotly_default_layout_options()
        layout['barmode'] = 'stack'
        layout['margin'] = pgo.Margin(**layout['margin'])
        layout['title'] = 'Chat (vs. Email) Days'
        layout['xaxis']['title'] = 'Day of the week'
        layout['yaxis']['title'] = 'Messages exchanged'

        return plotly_output(pgo.Figure(data=[chats_trace, emails_trace], layout=pgo.Layout(**layout)))

    def talk_durations(self):
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
        layout_args['title'] = 'Chat Durations'
        del layout_args['xaxis']
        del layout_args['yaxis']

        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[trace], layout=layout))

    def talk_thread_sizes(self):
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

        trace = pgo.Scatter(
            x=dates,
            y=messages,
            mode='markers',
            marker=dict(
                size=marker_sizes,
            ),
            text=descriptions
        )

        layout_args = plotly_default_layout_options()
        layout_args['title'] = 'Chat Thread Sizes'
        layout_args['hovermode'] = 'closest'
        layout_args['height'] = 800
        layout_args['margin'] = pgo.Margin(**layout_args['margin'])
        layout_args['xaxis']['title'] = 'Date'
        layout_args['yaxis']['title'] = 'Messages in thread'
        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[trace], layout=layout))

    def talk_times(self):
        """Returns a plotly graph showing chat habits by hour of the day (UTC).
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%H', `date`) AS hour, COUNT(message_key) AS talk_messages
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

        layout_args = plotly_default_layout_options()
        layout_args['title'] = 'Chat Times (UTC)'
        layout_args['xaxis']['title'] = 'Hour of day (UTC)'
        layout_args['yaxis']['title'] = 'Chat messages'

        trace = pgo.Scatter(**data_args)
        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[trace], layout=layout))

    def talk_top_chatters(self, limit=10):
        """Returns a plotly bar graph showing top chat senders with an email comparison.

        Keyword arguments:
            limit -- How many chat senders to return.
        """
        c = self.conn.cursor()

        c.execute('''SELECT `from`,
            COUNT(CASE WHEN gmail_labels LIKE '%Chat%' THEN 1 ELSE NULL END) AS talk_messages,
            COUNT(CASE WHEN gmail_labels NOT LIKE '%Chat%' THEN 1 ELSE NULL END) AS email_messages
            FROM messages
            WHERE `from` NOT LIKE ?
            GROUP BY `from`
            ORDER BY talk_messages DESC
            LIMIT ?;''', ('%' + self.owner_email + '%', limit,))

        chats = OrderedDict()
        emails = OrderedDict()
        longest_address = 0
        for row in c.fetchall():
            chats[row[0]] = row[1]
            emails[row[0]] = row[2]
            longest_address = max(longest_address, len(row[0]))

        chats_trace = pgo.Bar(
            x=chats.keys(),
            y=chats.values(),
            name='Chat messages',
            marker=dict(
                color=self.config.get('color', 'primary'),
            ),
        )
        emails_trace = pgo.Bar(
            x=emails.keys(),
            y=emails.values(),
            name='Email messages',
            marker=dict(
                color=self.config.get('color', 'secondary'),
            ),
        )

        layout = plotly_default_layout_options()
        layout['barmode'] = 'grouped'
        layout['height'] = longest_address * 15
        layout['margin']['b'] = longest_address * self.config.getfloat('font', 'size') / 2
        layout['margin'] = pgo.Margin(**layout['margin'])
        layout['title'] = 'Top ' + str(limit) + ' Chatters'
        layout['xaxis']['title'] = 'Sender address'
        layout['yaxis']['title'] = 'Messages received from'

        return plotly_output(pgo.Figure(data=[chats_trace, emails_trace], layout=pgo.Layout(**layout)))

    def talk_vs_email(self, cumulative=False):
        """Returns a plotly graph showing chat vs. email usage over time (by year and month).

        Keyword arguments:
            cumulative -- Whether ot not to display cumulative data for each month.
        """
        c = self.conn.cursor()

        c.execute('''SELECT strftime('%Y-%m', `date`) as period,
          COUNT(CASE WHEN gmail_labels LIKE '%Chat%' THEN 1 ELSE NULL END) AS talk_messages,
          COUNT(CASE WHEN gmail_labels NOT LIKE '%Chat%' THEN 1 ELSE NULL END) AS email_messages
          FROM messages
          GROUP BY period
          ORDER BY period ASC;''')

        talk_data = OrderedDict()
        talk_total = 0
        email_data = OrderedDict()
        email_total = 0
        for row in c.fetchall():
            talk_total += row[1]
            email_total += row[2]
            if cumulative:
                talk_data[row[0]] = talk_total
                email_data[row[0]] = email_total
            else:
                talk_data[row[0]] = row[1]
                email_data[row[0]] = row[2]

        talk_args = dict(
            x=talk_data.keys(),
            y=talk_data.values(),
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

        layout_args = plotly_default_layout_options()
        layout_args['title'] = 'Chat vs. Email Usage'
        layout_args['xaxis']['title'] = 'Year and month'
        layout_args['yaxis']['title'] = 'Number of messages'

        if cumulative:
            layout_args['title'] += ' (Cumulative)'
            talk_args['fill'] = 'tonexty'
            email_args['fill'] = 'tozeroy'

        talk_trace = pgo.Scatter(**talk_args)
        email_trace = pgo.Scatter(**email_args)
        layout = pgo.Layout(**layout_args)

        return plotly_output(pgo.Figure(data=[talk_trace, email_trace], layout=layout))

    def talk_vs_email_cumulative(self):
        """Returns the results of the talk_vs_email method with the cumulative argument set to True.
        """
        return self.talk_vs_email(cumulative=True)
