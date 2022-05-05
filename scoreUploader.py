from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import configReader as Config
import supportFunction as SFunc
import time
import os
from pathlib import Path
import re

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# READ FROM FILE
DATA = Config.configReader()
CONFIG = DATA[ "config" ]

# WAITING TIME
RELOAD_AFTER_SEC = CONFIG[ "RELOAD_AFTER_SEC" ]

# The ID and range of a sample spreadsheet.
SHEET_OUTPUT_ID = CONFIG[ 'SHEET_OUTPUT_ID' ]
SHEET_OUTPUT_NAME = CONFIG[ 'SHEET_OUTPUT_NAME' ]

# CONTEST MODE
CONTEST_MODE = CONFIG[ "CONTEST_MODE" ]

# IS DEV MODE
IS_DEV_MODE = CONFIG[ "IS_DEV_MODE" ]

# Config logs folder
FILE_OUT_AT = CONFIG[ "FILE_OUT_AT" ] #"FILE_OUT_AT": "./contestants/",
LOG_FOLDER_AT = FILE_OUT_AT + CONFIG[ "LOG_FOLDER_AT" ] #"LOG_FOLDER_AT": "Logs/",

isFirstRun = True

def main(credentialsFile, tokenFile):
    global isFirstRun
    print("Using", credentialsFile)
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(tokenFile):
        with open(tokenFile, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentialsFile, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(tokenFile, 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    # INIT SHEET
    if(CONTEST_MODE == "ACM"):
        RANGE_NAME = "!A1:B1"
        body = {'values': [["CONTESTANTS", "PENALTY"]]}
    else:
        RANGE_NAME = "!A1:A1"
        body = {'values': [["CONTESTANTS"]]}
    sheet.values().update(
        spreadsheetId=SHEET_OUTPUT_ID,
        range=RANGE_NAME,
        valueInputOption="RAW", 
        body=body
    ).execute()

    if(CONFIG["NEED_UPDATE_SHEET"] and isFirstRun):
        isFirstRun = False
        RANGE_NAME = "A1:A100"
        body = {'values': [["New Submission"]]+[["NO UPDATE"]]*99}
        sheet.values().update(
            spreadsheetId=SHEET_OUTPUT_ID,
            range=RANGE_NAME,
            valueInputOption="RAW", 
            body=body
        ).execute()


    for i in range(5):
        try:
            paths = sorted(Path("./contestants/Logs/").iterdir(), key=os.path.getmtime)
        except:
            time.sleep(5)
            continue
        for path in paths:
            path = str(path)
            if not path.endswith(".log"): continue
            print("     Loading", path)
            fp = open(path, "r", encoding="utf8")
            firstLine = fp.readline().strip()
            fp.close()
            if(IS_DEV_MODE): 
                print("In dev mode, OK rename {}".format(path))
            else: 
                os.rename(path, path+".done")
            try:
                info = re.findall("[.a-zA-Z0-9_]+", firstLine)
                studentName = info[0].strip().upper()
                problemCode = info[1].strip().upper()
                submitTime = int(re.findall("[0-9]+", path)[0])
                try:
                    score = float(info[2])
                except:
                    score = 0
                SFunc.updateScore(sheet, studentName, problemCode, score, submitTime, SFunc.countInQueue())
                print("     Done Loading", path)
            except:
                print("     Error when loading {}. Drop".format(path))
        time.sleep(5)

if __name__ == '__main__':
    while(1):
        try: main('credentials1.json', "token1.pickle")
        except: 
            print("Something goes wrong. Reloaded.")
            time.sleep(5)
        try: main('credentials.json', "token.pickle")
        except: 
            print("Something goes wrong. Reloaded.")
            time.sleep(5)
        try: main('credentials2.json', "token2.pickle")
        except: 
            print("Something goes wrong. Reloaded.")
            time.sleep(5)