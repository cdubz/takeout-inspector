"""example.py

This is a very basic example of Takeout Inspector's functionality. The code snippet below will
 1) Import mail from a Google Takeout Mail (mbox) file (specified in the settings.cfg file).
 2) Create a report with graphs (currently for Mail and Talk data).

Note: Large mail files will take a bit of time to process.
"""
from takeout_inspector import mail, report


mail.Import().import_messages()  # Includes Google Talk data.
report.Report().generate()
