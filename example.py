"""example.py

This is a very basic example of Takeout Inspector's functionality. The code snippet below will
 1) Import mail from a Google Takeout Mail (mbox) file (specified in the settings.cfg file).
 2) Create and open (in a web browser) a Plotly-based "Top 20 Recipients" graph.

Note: Large mail files will take a bit of time to process.
"""
from takeout_inspector import mail


p = mail.Import()
p.import_messages()

g = mail.Graph()
g.top_recipients(20)
