from __future__ import print_function
import pickle
import sys
import os.path
import pprint
import re
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
"""import logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
import httplib2
httplib2.debuglevel = 4"""

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/admin.directory.group'
]

SPREADSHEET_ID = None
SPREADSHEET_RANGE_NAME = 'G2:I'
SPREADSHEET_EMAIL_INDEX = 0  # in "range space"
SPREADSHEET_STATUS_INDEX = 2  # in "range space"

GROUP_ID = None


def error_usage():
    print('Usage: {} spreadsheet_id group_id'.format(__file__))
    sys.exit(1)


def parse_args():
    global SPREADSHEET_ID, GROUP_ID
    if len(sys.argv) <= 2:
        error_usage()
    SPREADSHEET_ID = sys.argv[1]
    GROUP_ID = sys.argv[2]


def auth_google():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def get_member_emails():
    if not SPREADSHEET_ID:
        error_usage()
    print("Using Spreadsheet ID {}".format(SPREADSHEET_ID))

    creds = auth_google()

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID, range=SPREADSHEET_RANGE_NAME).execute()
    values = result.get('values', [])

    emails = []

    if not values:
        print('No data found.')
    else:
        for row in values:
            email = ""
            status = ""
            if len(row) > SPREADSHEET_EMAIL_INDEX:
                email = row[SPREADSHEET_EMAIL_INDEX]
            if len(row) > SPREADSHEET_STATUS_INDEX:
                status = row[SPREADSHEET_STATUS_INDEX]
            if status != 'Approved':
                print('SKIP: Status "{}" of {} is not "Approved", skipping...'.
                      format(status, email))
                continue
            if not re.match(r'[^@]+@[^@]+\.[^@]+', email):
                print('SKIP: E-mail "{}" failed sanity check, skipping...'.
                      format(email))
                continue
            emails.append(email.lower())

    return emails


def get_group_members():
    if not GROUP_ID:
        error_usage()
    print('Using Group ID {}'.format(GROUP_ID))

    creds = auth_google()
    service = build('admin', 'directory_v1', credentials=creds)

    members = []

    max_per_page = 2
    next_page_token = None
    gotten_all_pages = False
    while not gotten_all_pages:
        results = service.members().list(
            groupKey=GROUP_ID,
            includeDerivedMembership=True,
            maxResults=max_per_page,
            pageToken=next_page_token).execute()
        next_page_token = results.get('nextPageToken')
        if not next_page_token:
            gotten_all_pages = True
        for member in results.get('members', []):
            members.append(member['email'].lower())

    return members


def add_member_to_group(email):
    if not GROUP_ID:
        error_usage()
    creds = auth_google()
    service = build('admin', 'directory_v1', credentials=creds)

    user_body = {
        'kind': 'admin#directory#member',
        'role': 'MEMBER',
        'type': 'USER',
        'email': email
    }
    results = service.members().insert(
        groupKey=GROUP_ID, body=user_body).execute()


def remove_member_from_group(email):
    if not GROUP_ID:
        error_usage()
    creds = auth_google()
    service = build('admin', 'directory_v1', credentials=creds)

    results = service.members().delete(
        groupKey=GROUP_ID, memberKey=email).execute()


def main():
    member_emails = get_member_emails()
    google_group_members = get_group_members()

    group_members_to_add = [
        x for x in member_emails if x not in google_group_members
    ]
    group_members_to_remove = [
        x for x in google_group_members if x not in member_emails
    ]

    for to_add in group_members_to_add:
        print('+ {}'.format(to_add))
        add_member_to_group(to_add)
    for to_remove in group_members_to_remove:
        print('- {}'.format(to_remove))
        remove_member_from_group(to_remove)

    print('Done')


if __name__ == '__main__':
    parse_args()
    main()
