# Takeout Inspector
Parse and inspect data from [Google Takeout](https://takeout.google.com/settings/takeout) exports.

# Prerequisites

## Python Requirements
* Python 2.7
* SQLite
* [names](https://pypi.python.org/pypi/names)
* [Plotly](https://pypi.python.org/pypi/plotly)

## Google Takeout Files
In order to do anything with Takeout Inspector, you will need some exported data from Google Takeout. To get these files:

1. Navigate to [Google Takeout](https://takeout.google.com/settings/takeout) and log in if necessary.
1. Enable the data you want to export (see **[Supported Data Types](#supported-data-types)**) under the **Select data to include** heading.
1. Click **Next**.
1. Set the **File type** and **Delivery method** as desired.
1. Click **Create archive**.
1. Wait! It will take a while for the data to arrive.

Particularly large *Mail* archives may take a very long time to process.

# Installation

1. Download and unpack the latest version.
1. Run ```pip install -r requirements.txt``` from the unpacked folder.
1. Copy ```./settings.defaults.cfg``` to ```./settings.cfg```.
1. Modify ```./settings.cfg``` to your liking (most importantly, provide data file paths).
1. Tinker with and run ```./example.py```!

# Supported Data Types

- [ ] Chrome Browser History
- [ ] Hangouts
- [ ] Location History
- [x] Mail
