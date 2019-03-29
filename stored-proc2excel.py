import mysql.connector
import json
import xlsxwriter
from time import gmtime, strftime
import os
import re

def stoi(text):
  if text.isdigit():
    return int(text)
  else:
    return text

def naturalKeys(ele):
  return [stoi(c) for c in re.split('(\d+)', ele["comment"])]

def getStoredProcedures():
  global connection
  cursor = connection.cursor()
  cursor.execute("SHOW PROCEDURE STATUS")
  procedures = []
  for (db, name, type, definer, modified, created, securityType, comment, characterSetClient, collectionConnection, databaseCollation) in cursor:
    procedures.append({ "name": name, "comment": comment })
  cursor.close()
  return procedures

def getParametersForStoredProcedure(storedProcedure):
  global connection
  cursor = connection.cursor()
  cursor.execute("SELECT * FROM information_schema.parameters WHERE SPECIFIC_NAME = '" + storedProcedure + "' ORDER BY ORDINAL_POSITION")
  parameters = []
  for (specificCatalog, specificSchema, specificName, ordinalPosition, parameterMode, parameterName, dataType, characterMaximumLengt, characterOctetLength, numericPrecision, numericscale, characterSetName, collectionName, dtdIdentifier, routineType) in cursor:
    parameters.append({ "parameterName": parameterName })
  cursor.close()
  return parameters

def getLineLength(text):
  splits = text.split("\n")
  length = 0
  for line in splits:
    if len(line) > length:
      length = len(line)
  return length

config = None
with open("config.json") as file:
  config = json.loads(file.read())

connection = mysql.connector.connect(host = config["host"], user = config["user"], password = config["pass"], database = config["db"])
print("Using Database: " + config["db"])

print("Procedures:")
procedures = getStoredProcedures()
procedures.sort(key = naturalKeys)
for index in range(0, len(procedures)):
  print(procedures[index]["comment"] + " - " + procedures[index]["name"])
choice = int(input("Take (Ctrl+C to abort): "))

parameters = getParametersForStoredProcedure(procedures[choice - 1]["name"])
filenamePrefix = None
for index in range(0, len(parameters)):
  parameters[index]["input"] = input(parameters[index]["parameterName"] + ": ")
  if filenamePrefix == None:
    filenamePrefix = parameters[index]["input"]

sql = "CALL " + procedures[choice - 1]["name"] + "("
for index in range(0, len(parameters)):
  sql = sql + parameters[index]["input"]
  if index < len(parameters) - 1:
    sql = sql + ", "
sql = sql + ");"

timeString = strftime("%Y-%m-%d %H.%M.%S", gmtime())
filename = filenamePrefix + " " + procedures[choice - 1]["name"] + "-" + timeString + ".xlsx"

cursor = connection.cursor(dictionary = True)
procParams = tuple([parameters[index]["input"] for index in range(0, len(parameters))])
cursor.callproc(procedures[choice - 1]["name"], procParams)

workbook = xlsxwriter.Workbook(filename)
defaultCellFormat = workbook.add_format()
defaultCellFormat.set_text_wrap()
defaultCellFormat.set_align("vcenter")

resultIndex = 0

nextConfig = None

for result in cursor.stored_results():
  resultIndex = resultIndex + 1
  statement = result.statement
  match = re.search('\(a result of CALL (.*)\(.*\)\)', statement)
  sheetName = "(" + str(resultIndex) + ") " + match.group(1)
  fieldNames = [field[0] for field in result.description]
  fieldLength = [0 for field in result.description]

  if len(fieldNames) == 1:
    if fieldNames[0] == "__config":
      config = json.loads(result.fetchone()[0])
      if "workbookName" in config:
        workbook.close()
        workbook = xlsxwriter.Workbook(config["workbookName"] + ".xlsx")
        defaultCellFormat = workbook.add_format()
        defaultCellFormat.set_text_wrap()
        defaultCellFormat.set_align("vcenter")
      elif "sheetName" in config:
        nextConfig = config
      continue

  worksheet = None
  if nextConfig != None:
    worksheet = workbook.add_worksheet(nextConfig["sheetName"])
  elif match:
    worksheet = workbook.add_worksheet(sheetName)
  else:
    worksheet = workbook.add_worksheet()

  if nextConfig != None:
    worksheet.set_tab_color("#00FF00") #nextConfig["tabColor"]
  else:
    worksheet.set_tab_color("#FF0000")

  for fieldIndex in range(0, len(fieldNames)):
    worksheet.write(0, fieldIndex, str(fieldNames[fieldIndex]))
    fieldLength[fieldIndex] = len(str(fieldNames[fieldIndex]))

  rowIndex = 0
  for row in result:
    rowIndex = rowIndex + 1
    for fieldIndex in range(0, len(fieldNames)):
      maxLineLength = getLineLength(str(row[fieldIndex]))
      if maxLineLength > fieldLength[fieldIndex]:
        fieldLength[fieldIndex] = maxLineLength
      if row[fieldIndex] != None:
        worksheet.write(rowIndex, fieldIndex, str(row[fieldIndex]), defaultCellFormat)

  for fieldIndex in range(0, len(fieldLength)):
    worksheet.set_column(fieldIndex, fieldIndex, fieldLength[fieldIndex] * 1.1)
  worksheet.autofilter(0, 0, rowIndex, len(fieldLength))
  nextConfig = None

workbook.close()
print("file " + filename + " written")
openChoice = input("Open it [Y/n]? ")
if openChoice.lower() == "y" or openChoice == "":
  os.system("xdg-open '" + filename + "'")
  os.system("\"" + filename + "\"")

cursor.close()
connection.close()
