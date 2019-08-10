from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import sys
import re
import statistics

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE_NAME = 'Class Data!A2:E'

REGEX_FOR_SHEET_ID = r"[#&]gid=([0-9]+)"
REGEX_FOR_SPREADSHEET_ID = r"/spreadsheets/d/([a-zA-Z0-9-_]+)"


def remove_turkish_characters(word):
    word = word.replace("ç", "c")
    word = word.replace("ğ", "g")
    word = word.replace("ı", "i")
    word = word.replace("ö", "o")
    word = word.replace("ş", "s")
    word = word.replace("ü", "u")
    return word


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    spreadsheets_resource = service.spreadsheets()

    spreadsheet_url = sys.argv[1]
    spreadsheet_id = re.search(REGEX_FOR_SPREADSHEET_ID, spreadsheet_url).group(1);
    sheet_id = re.search(REGEX_FOR_SHEET_ID, spreadsheet_url).group(1);

    response = spreadsheets_resource.get(spreadsheetId=spreadsheet_id).execute()

    sheet_name = response["sheets"][0]["properties"]["title"]

    range_in_a1_notation = sheet_name

    result = spreadsheets_resource.values().get(spreadsheetId=spreadsheet_id, range=range_in_a1_notation).execute()

    attendance = {}

    rows = result.get('values', [])

    if not rows:
        print('No data found.')
    else:
        for row in rows[1:]:
            duration = row[9]
            participant_name = remove_turkish_characters(row[11].lower())
            if participant_name in attendance:
                attendance[participant_name] += float(duration)
            else:
                attendance[participant_name] = float(duration)

        durations = list(attendance.values())
        mean = statistics.mean(durations)
        stddevp = statistics.pstdev(durations, mean)
        threshold = mean - 2 * stddevp

        sorted_participant_names = sorted(list(attendance))

        with open("attendance.txt", "w") as attendance_file:
            attendance_file.write("mean:" + str(mean) + "\n" + "stddevp:" + str(stddevp) + "\n" + "threshold:" + str(threshold) + "\n\n")
            for participant_name in sorted_participant_names:
                attendance_file.write(("1" if attendance[participant_name] > threshold else "0") + " " + participant_name + ": " + str(attendance[participant_name]) + " " + "\n")

        print("attendance.txt created successfully!")

if __name__ == '__main__':
    main()