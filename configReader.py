import re
import json 

def envInfo():
  try:
    configFile = open(".env").read()
  except:
    configFile = open(".env.json").read()
  return configFile

def configReader():
  result = {}
  configFile = envInfo()
  config = json.loads(configFile)
  studentsFile = open(config["STUDENT_LIST"]).read()
  students = json.loads(studentsFile)
  result[ "config" ] = config
  result[ "students" ] = students
  return result

def infomationTaker(key):
  configFile = envInfo()
  config = json.loads(configFile)
  return config[ key ]

##########################################
## DEBUG ##
##########################################

def __MAIN__():
  print( configReader() )

if __name__ == '__main__':
  __MAIN__()