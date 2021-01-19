#!/usr/bin/env python3

import webbrowser
import base64
import email
import os
import sys
from datetime import date, datetime
from dateutil import parser
import time
import re
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

PATH = "/Users/eric/Code/Python/c2-load"
# PATH = os.getcwd()
BROWSER_COMMAND = "open -a /Applications/Google\ Chrome.app %s"
SCOPES_GMAIL = ["https://www.googleapis.com/auth/gmail.readonly"]
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive.metadata.readonly"]

os.chdir(PATH)

def get_gmail_creds():
    # The file token_gmail.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds_gmail = None
    if os.path.exists('token_gmail.pickle'):
        with open('token_gmail.pickle', 'rb') as token:
            creds_gmail = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds_gmail or not creds_gmail.valid:
        if creds_gmail and creds_gmail.expired and creds_gmail.refresh_token:
            creds_gmail.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES_GMAIL)
            creds_gmail = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token_gmail.pickle', 'wb') as token:
            pickle.dump(creds_gmail, token)
    return creds_gmail

def get_drive_creds():
    creds_drive = None
    if os.path.exists('token_drive.pickle'):
        with open('token_drive.pickle', 'rb') as token:
            creds_drive = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds_drive or not creds_drive.valid:
        if creds_drive and creds_drive.expired and creds_drive.refresh_token:
            creds_drive.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES_DRIVE)
            creds_drive = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token_drive.pickle', 'wb') as token:
            pickle.dump(creds_drive, token)
    return creds_drive

def get_messages(service, user_id):
    try:
        return service.users().messages().list(userId=user_id).execute()
    except Exception as error:
        print('An error occurred: %s' % error)

def get_mime_message(service, user_id, msg_id):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id, format='raw').execute()
        # print('Message snippet: %s' % message['snippet'])
        msg_str = base64.urlsafe_b64decode(message['raw'].encode("utf-8")).decode("utf-8")
        mime_msg = email.message_from_string(msg_str)
        return mime_msg
    except Exception as error:
        print('An error occurred: %s' % error)

def get_schedule(service, date_):
    """Search inbox for the day's schedule"""
    all_msg = get_messages(service, 'me')['messages']
    all_ids = [all_msg[i]['id'] for i in range(len(all_msg))]
    for id_ in all_ids:
        msg = get_mime_message(service, 'me', id_)
        multi = msg.is_multipart()
        if not multi:
            continue
        else:
            raw_list = msg.get_payload()
            raw = raw_list[0].get_payload()
            try:
                if date_ == parser.parse(raw.split()[1]).date():
                    schedule_id = id_
                    break
            except:
                continue
    return schedule_id

def get_names(service, schedule_id):
    """Find the body of the email message"""
    msg = get_mime_message(service, "me", schedule_id)
    raw = msg.get_payload()[0].get_payload()
    # start = raw.find("5-Minute Room Prep")
    start = raw.find("Teacher/Time")
    end = raw.find("\n--")
    raw2 = raw[start:end]
    raw3 = re.sub(r"\r\n|Form|Review", " ", raw2)
    raw4 = re.sub(r">", " ", raw3)
    raw5 = re.sub(r"5-Minute Room Prep", " ", raw4)
    # clean = re.findall(r"[A-Z][a-z]*\s[A-Z][a-z]*\s+\d+|[A-Z][a-z]*\s[A-Z][a-z]*\s+College|[A-Z][a-z]*\s[A-Z][a-z]*\s+UW", raw4)
    clean = re.findall(r"[A-Z][a-z]*\s[A-Z][a-z]*\s+\d+|[A-Z][a-z]*\s+\([A-Z][a-z]*\)\s+[A-Z][a-z]*\s+\d+|[A-Z][a-z]*\s[A-Z][a-z]*\s+College|[A-Z][a-z]*\s[A-Z][a-z]*\s+UW", raw5)
    names = [re.sub(r"\s+\d+|\s+College|\s+UW", "", name) for name in clean]
    return names

def get_bluebook_id(service, name):
    s = "name contains '" + name + "' and mimeType='application/vnd.google-apps.folder'"
    results = service.files().list(q=s).execute()
    items = results.get('files', [])
    bluebook_id = ""
    if not items:
        raise Exception("No bluebook found for student " + name + ".")
        # return bluebook_id
    for item in items:
        try:
            id_ = item.get('id')
            sub_s = "'" + id_ + "' in parents and name contains 'Digital Blue Book'"
            subresults = service.files().list(q=sub_s).execute()
            subitems = subresults.get('files', [])
            if len(subitems) != 1:
                print("Unique blue book not found for " + name + ". Using the first one found.")
            bluebook_id = subitems[0].get("id")
        except:
            continue
    if not bluebook_id:
        print("No bluebook found for student " + name + ".")
    return bluebook_id

def main():
    """Loads the C2 schedule and opens all red binder entries in new tabs
    """
    gmail = build("gmail", "v1", credentials=get_gmail_creds())
    drive = build("drive", "v3", credentials=get_drive_creds())

    if os.path.isfile("students.pickle"):
        try:
            with open("students.pickle", "rb") as handle:
                students = pickle.load(handle)
        except:
            students = {}
    else:
        students = {}

    # Create the date object
    if len(sys.argv) > 1:
        date_ = parser.parse(sys.argv[1]).date()
    else:
        date_ = datetime.now().date()

    try:
        schedule_id = get_schedule(gmail, date_)
    except:
        all_msg = get_messages(gmail, 'me')['messages']
        schedule_id = all_msg[0]['id']

    names = get_names(gmail, schedule_id)
    for name in names:
        if name not in students.keys(): # Add any new students to students dictionary
            try:
                bluebook = get_bluebook_id(drive, name)
                students[name] = bluebook
            except:
                continue
        id_ = students.get(name)
        webbrowser.get(BROWSER_COMMAND).open("https://docs.google.com/spreadsheets/d/" + id_)

    webbrowser.get(BROWSER_COMMAND).open("https://c2educate.adobeconnect.com/sammamish?proto=true")

    with open("students.pickle", "wb") as handle:
            pickle.dump(students, handle, protocol=pickle.HIGHEST_PROTOCOL)

if __name__ == "__main__":
    main()
