from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import configReader as Config
import os
from random import randint
import time
from pathlib import Path
from datetime import datetime
import re

# writeCol = chr(writeCol-1+ord('A'))
def colNumToColString(colNum):
  # 1  -> A
  # 26 -> Z
  # 27 -> AA
  res = ''
  while colNum > 0:
    res = chr((colNum-1)%26+ord('A')) + res
    colNum = int (colNum/26)
  return res

def getRow(sheet, row):
  SHEET_INPUT_ID = Config.infomationTaker("SHEET_INPUT_ID")
  #SHEET_INPUT_NAME = Config.infomationTaker("SHEET_INPUT_NAME")
  RANGE_NAME = str(row)+":"+str(row)
  result = sheet.values().get(spreadsheetId=SHEET_INPUT_ID, range=RANGE_NAME).execute()
  values = result.get('values', [])

  return values

def writeToFile(rowValue, dateAndTime):
  # DATA
  DATA = Config.configReader()
  STUDENTS = DATA[ "students" ]
  CONFIG = DATA[ "config" ]

  print(DATA)
  print(STUDENTS)
  print(CONFIG)

  # UPDATE SUBMISSION
  SECRET_CODE_COL = CONFIG[ "SECRET_CODE_COL" ]
  PROBLEM_CODE_COL = CONFIG[ "PROBLEM_CODE_COL" ]
  CODE_COL = CONFIG[ "CODE_COL" ]
  FILE_TYPE = CONFIG[ "FILE_TYPE" ]

  # FOLDER URL
  FILE_OUT_AT = CONFIG[ "FILE_OUT_AT" ]
  if not os.path.exists(FILE_OUT_AT): os.makedirs(FILE_OUT_AT)
  
  # PRINT
  timestamp = time.mktime(time.strptime(dateAndTime, '%m/%d/%Y %H:%M:%S'))
  timestamp = int(timestamp)
  studentName = STUDENTS.get(rowValue[SECRET_CODE_COL], None)
  if( studentName==None ): return timestamp
  problemName = rowValue[PROBLEM_CODE_COL]
  fileName = "{}{}[{}][{}].{}".format(FILE_OUT_AT, timestamp, studentName, problemName, FILE_TYPE)
  code = rowValue[ CODE_COL ]
  fp = open(fileName, "w")
  fp.write( code )
  fp.close()
  return timestamp

def markDone(sheet, dateAndTime, row):
  SHEET_INPUT_ID = Config.infomationTaker("SHEET_INPUT_ID")
  #SHEET_INPUT_NAME = Config.infomationTaker("SHEET_INPUT_NAME")
  RANGE_NAME = "A"+str(row)+":D"+str(row)
  print(RANGE_NAME)
  body = {
    'values': [[dateAndTime]]
  }
  result = sheet.values().update(
    spreadsheetId=SHEET_INPUT_ID,
    range=RANGE_NAME,
    valueInputOption="RAW", 
    body=body
  ).execute()
  print('{0} cells updated.'.format(result.get('updatedCells')))

def getRangeName(sheet, student, problem):
  SHEET_OUTPUT_ID = Config.infomationTaker("SHEET_OUTPUT_ID")
  SHEET_OUTPUT_NAME = Config.infomationTaker("SHEET_OUTPUT_NAME")

  # GET ROW
  FIRST_COL_RANGE_NAME = SHEET_OUTPUT_NAME+"!A:A"
  result = sheet.values().get(spreadsheetId=SHEET_OUTPUT_ID, range=FIRST_COL_RANGE_NAME).execute()
  values = result.get('values', [])
  found = 0
  writeRow = 1
  for value in values:
    if(value[0].strip().upper() == student): 
      found = 1
      break
    writeRow+=1
  if found == 0:
    writeRow = len(values) + 1
  # REWRITE
  if found == 0:
    body = {
      'values': [[student]]
    }
    RANGE_NAME = "{}!A{}:A{}".format(SHEET_OUTPUT_NAME, writeRow, writeRow)
    result = sheet.values().update(
      spreadsheetId=SHEET_OUTPUT_ID,
      range=RANGE_NAME,
      valueInputOption="RAW", 
      body=body
    ).execute()

  # GET COL
  FIRST_ROW_RANGE_NAME = SHEET_OUTPUT_NAME+"!1:1"
  result = sheet.values().get(spreadsheetId=SHEET_OUTPUT_ID, range=FIRST_ROW_RANGE_NAME).execute()
  values = result.get('values', [])
  found = 0
  writeCol = 1
  for value in values[0]:
    if(value.strip().upper() == problem): 
      found = 1
      break
    writeCol+=1
  if found==0:
    writeCol = len( values[0] ) + 1
  writeCol = colNumToColString(writeCol)
  # REWRITE
  if found == 0:
    body = {
      'values': [[problem]]
    }
    RANGE_NAME = "{}!{}1:{}1".format(SHEET_OUTPUT_NAME, writeCol, writeCol)
    result = sheet.values().update(
      spreadsheetId=SHEET_OUTPUT_ID,
      range=RANGE_NAME,
      valueInputOption="RAW", 
      body=body
    ).execute()
  
  # Get RANGE_NAME
  RANGE_NAME = "{}!{}{}:{}{}".format(SHEET_OUTPUT_NAME, writeCol, writeRow, writeCol, writeRow)
  return RANGE_NAME

def updatePenalty(sheet, RANGE_NAME, score, submitTime):
  SHEET_OUTPUT_ID = Config.infomationTaker("SHEET_OUTPUT_ID")
  WRONG_SUBMISSION_PENALTY = Config.infomationTaker("WRONG_SUBMISSION_PENALTY")
  START_TIME = Config.infomationTaker("START_TIME")
  timestamp = time.mktime(time.strptime(START_TIME, '%m/%d/%Y %H:%M:%S'))
  START_TIMESTAMP = int(timestamp)
  try:
    result = sheet.values().get(spreadsheetId=SHEET_OUTPUT_ID, range=RANGE_NAME).execute()
    currPenalty = float(result.get('values', [])[0][0])
  except:
    currPenalty = 0
  currPenalty += (score-1)*WRONG_SUBMISSION_PENALTY + int((submitTime-START_TIMESTAMP)/60)
  body = {
    'values': [[currPenalty]]
  }
  result = sheet.values().update(
    spreadsheetId=SHEET_OUTPUT_ID,
    range=RANGE_NAME,
    valueInputOption="RAW", 
    body=body
  ).execute()
  print("          RANGE_NAME: {} - Score: {} - Penalty: {}".format(RANGE_NAME, score, currPenalty))

def updateStatus(sheet, newStatus):
  # GET VALUE
  SHEET_OUTPUT_ID = Config.infomationTaker("SHEET_OUTPUT_ID")
  SHEET_UPDATE_NAME = Config.infomationTaker("SHEET_UPDATE_NAME")
  RANGE_NAME = SHEET_UPDATE_NAME+"!A2:A99"
  result = sheet.values().get(spreadsheetId=SHEET_OUTPUT_ID, range=RANGE_NAME).execute()
  values = result.get('values', [])
  values = [[newStatus]] + values

  # RE-WRITE
  RANGE_NAME = SHEET_UPDATE_NAME+"!A2:A100"
  body = {'values': values}
  sheet.values().update(
      spreadsheetId=SHEET_OUTPUT_ID,
      range=RANGE_NAME,
      valueInputOption="RAW", 
      body=body
  ).execute()

def updateScore(sheet, student, problem, score, submitTime, inQueue = 0):
  if( Config.infomationTaker("IS_DEV_MODE") ):
    student = "{}_{}".format(student, randint(1, 10))
    problem = "{}_{}".format(problem, randint(1, 10))
    score = score+randint(1, 10)
  # Log to file
  f= open("contest_log.txt","a")
  newSubmitStatus =  "Name: {}, Problems: {}, Score: {}, Time: {}, In Queue: {}\n".format(student, problem, score, datetime.fromtimestamp(submitTime), inQueue)
  f.write(newSubmitStatus)
  f.close()
  SHEET_OUTPUT_ID = Config.infomationTaker("SHEET_OUTPUT_ID")
  CONTEST_MODE = Config.infomationTaker("CONTEST_MODE")
  RANGE_NAME = getRangeName(sheet, student, problem)
  isUpdatePenaltyNeeded = 0
  try:
    result = sheet.values().get(spreadsheetId=SHEET_OUTPUT_ID, range=RANGE_NAME).execute()
    currScore = float(result.get('values', [])[0][0])
  except:
    currScore = 0
  if CONTEST_MODE == "ACM":
    if currScore <= 0: currScore-=1 # one more submission
    if score >= 10 and currScore<0: # not yet AC and now AC
      score = -currScore
      isUpdatePenaltyNeeded = 1
    else: score = currScore
  else:
    score = max(currScore, score)
  body = {
    'values': [[score]]
  }
  result = sheet.values().update(
    spreadsheetId=SHEET_OUTPUT_ID,
    range=RANGE_NAME,
    valueInputOption="RAW", 
    body=body
  ).execute()
  if isUpdatePenaltyNeeded:
    RANGE_NAME_PENALTY = getRangeName(sheet, student, "PENALTY")
    updatePenalty(sheet, RANGE_NAME_PENALTY, score, submitTime)
  updateStatus(sheet, newSubmitStatus)
  print("          {} - {} - {}, score is {} at timestamp {}".format(student, problem, RANGE_NAME, score, submitTime))

def countInQueue(dirPath = "./contestants/Logs/"):
  resLogs = 0
  try:
    paths = sorted(Path(dirPath).iterdir(), key=os.path.getmtime)
    for path in paths:
      if (path.endswith(".log")): resLogs = resLogs+1
  except:
    resLogs = 0
  resCpps = 0
  try:
    dirPath = re.findall(r'(.+)Logs')[0]
    paths = sorted(Path(dirPath).iterdir(), key=os.path.getmtime)
    for path in paths:
      if (path.endswith(".cpp")): resCpps = resCpps+1
  except:
    resCpps = 0
  return resCpps + resLogs