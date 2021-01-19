# c2-load

This script is intended for use by teachers employed by C2 Education during the remote teaching period caused by COVID-19 to automate the workflow for opening the day's blue books. 

To install, you will need both Python 3.8 and its package manager pip. Run the following pip commands to install the necessary libraries:

pip3 install --upgrade python-dateutil google-api-python-client google-auth-httplib2 google-auth-oauthlib

Download this repository. You will also need to enable the Google doc API by visiting https://developers.google.com/docs/api/quickstart/python.

Configure your Oauth client for a Desktop app. Save the configuration into the same directory you downloaded as a clone of this repository. Then run the program by either typing either of the following

./c2_load.py

python3 c2_load.py

The first time the program runs, it will prompt you for permission to view your gmail and view metadata for your Google drive files. If you get a warning that the app is unverified, you can bypass this by clicking on "Advanced" and following the prompt. 

When run, this script will load C2 Education bluebook/red-binder Google sheets automatically from the schedule email for today's date (or another date if provided as an argument to the script in date format yyyy-mm-dd). Requires enabling both the Gmail and Google Drive APIs (running the script will prompt you to enable these automatically the first time it is run). The script will save some of the search information for quick, future lookup, and the directory this information is saved in needs to be set manually in your downloaded version of the code before running.

