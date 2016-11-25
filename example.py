"""example.py

This is a very basic example of Takeout Inspector's functionality. The code snippet below will
 1) Import mail from a Google Takeout Mail (mbox) file (specified in the settings.cfg file).
 2) Create Mail report, `mail.html`, containing all current Mail graphs.
 3) Create Talk report, `talk.html`, containing all current Talk graphs.

Note: Large mail files will take a bit of time to process.
"""
from takeout_inspector import mail, talk


p = mail.Import()
p.import_messages()

m = mail.Graph()
m.all_graphs()

t = talk.Graph()
t.all_graphs()
