from __future__ import print_function
import pickle
import sys
import os.path
import pprint
import re
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SPREADSHEET_RANGE_NAME = 'G2:I'
SPREADSHEET_EMAIL_INDEX = 0  # in "range space"
SPREADSHEET_STATUS_INDEX = 2  # in "range space"


def get_member_emails():
    if not len(sys.argv) >= 2:
        print(
            'Provide Spreadsheet ID (e.g. 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms) on command line'
        )
        sys.exit(1)
    spreadsheet_id = sys.argv[1]
    print("Using Spreadsheet ID {}".format(spreadsheet_id))

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

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=spreadsheet_id, range=SPREADSHEET_RANGE_NAME).execute()
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
            emails.append(email)

    return emails


def main():
    emails = get_member_emails()
    pprint.pprint(emails)


if __name__ == '__main__':
    main()
