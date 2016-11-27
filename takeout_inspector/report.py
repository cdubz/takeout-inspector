"""takeout_inspector/report.py

Defines classes and methods used to generate reports from imported data.

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
import inspect
import os
import shutil

from . import mail, talk

__all__ = ['Report']


class Report:
    """Creates offline plotly graphs using imported data from sqlite.
    """
    def __init__(self):
        self.config = ConfigParser.ConfigParser()
        self.config.readfp(open('settings.defaults.cfg'))
        self.config.read(['settings.cfg'])

        self.base_dir = self.config.get('report', 'destination')

        if not os.path.isdir(self.base_dir):
            os.mkdir(self.base_dir)
            shutil.copytree('resources', self.base_dir + '/resources')

    def generate(self):
        """Creates a page containing all available Talk graphs. The HTML file (talk.html) and supporting JavaScript file
        (talk.js) are both saved to the local directory. The page relies on two JavaScript libraries which are included
        in the `resources/js` directory of Takeout Inspector:
          - Plotly: https://plot.ly/javascript/
          - WayPoints: http://imakewebthings.com/waypoints/ (Note: the JS file erroneously states v4.0.0 but is v4.0.1.)
        """
        graph_classes = [mail.Graph(), talk.Graph()]
        for graph_class in graph_classes:
            report = graph_class.__dict__['report']
            methods = inspect.getmembers(graph_class, inspect.ismethod)

            html_file = self.base_dir + report.lower() + '.html'
            js_file = self.base_dir + '/resources/js/' + report.lower() + '.js'

            with open(html_file, 'w') as html, open(js_file, 'w') as js:
                html.write(''.join([
                    '<!DOCTYPE HTML>\n',
                    '<html>\n',
                    '<head>\n',
                    '\t<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />\n',
                    '\t<title>' + report + ' | Takeout Inspector</title>\n',
                    '</head>\n',
                    '<body style="max-width: 800px; margin: 0 auto;">\n',
                    '<h1 style="text-align: center;">' + report + ' Statistics</h1>\n'
                ]))

                for method in methods:
                    if method[0][0] == '_':
                        continue
                    output = method[1]()

                    try:
                        div, javascript = output.split('<script type="text/javascript">')
                        js.write(''.join([
                            'new Waypoint({\n',
                            "\telement: document.getElementById('" + div[9:45] + "'),\n",  # String location of div ID.
                            '\thandler: function() {\n',
                            '\t\t' + javascript[:-9] + ';\n',  # Removes </script> from the end of the string.
                            '\t\tthis.destroy();\n',
                            '\t},\n',
                            "\toffset: '100%'\n"
                            '});\n\n'
                        ]))
                    except ValueError:
                        div = output
                    finally:
                        html.write(div + '\n')

                html.write(''.join([
                    '<script src="resources/js/plotly-v1.20.5.min.js"></script>\n',
                    '<script src="resources/js/waypoints-v4.0.1.min.js"></script>\n',
                    '<script src="resources/js/' + report.lower() + '.js"></script>\n',
                    '</body>\n',
                    '</html>',
                ]))
