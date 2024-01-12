from __future__ import division, print_function
from re import M
from googleapiclient.discovery import build
from google.oauth2 import service_account
import cv2

# Instructions for creating keys.json file: 
# Follow the instructions at https://developers.google.com/sheets/api/quickstart/python until you get to the credentials page
# Instead of creating an OAuth token just make a service account 
# after creating the service account locate the keys tab and add a key
# that should download a file that contains the keys as a json object

#Setup Sheets Api:

# This id is the spreadsheet (which is literally the characters in the link after spreadsheets/d/ )
ID = '1TKzmDRF9mtz1c9R7R1RPCXSjV9EjvK3XKKl8dheOY-E' 
service = build('sheets', 'v4', credentials=service_account.Credentials.from_service_account_file("keys.json",scopes=['https://www.googleapis.com/auth/spreadsheets']))
sheet = service.spreadsheets()

# the sheetID is the id of the individual sheet/page within the document 
# this can be found at the very end of the link #gid=[0] <- the sheet id is the number at the end
def updateCellColors(sheetID, values):
  range = {
              "sheetId": sheetID,
              "startRowIndex": 0,
              "endRowIndex": len(values),
              "startColumnIndex": 0,
              "endColumnIndex": len(values[0]['values'])
            }
  body = {
    "requests": [{
            "updateSheetProperties": {
                "properties": {
                    "gridProperties": {
                        "rowCount": len(values),
                        "columnCount": len(values[0]['values'])
                    },
                    "sheetId": sheetID
                },
                "fields": "gridProperties"
            }
        },
        {"updateDimensionProperties":{
          "range": {
            "sheetId": sheetID,
            'dimension': "ROWS",
            'startIndex': 0,
            'endIndex': len(values)
          },
          'properties': {
            'pixelSize': 3
          },
          'fields': 'pixelSize'
        }},
        {"updateDimensionProperties":{
          "range": {
            "sheetId": sheetID,
            'dimension': "COLUMNS",
            'startIndex': 0,
            'endIndex': len(values)
          },
          'properties': {
            'pixelSize': 2
          },
          'fields': 'pixelSize'
        }},
        {
          "updateCells": {
            "range": range,
            "rows": values,
            "fields": "userEnteredFormat.backgroundColor"
          }
        }
    ]
  }

  service.spreadsheets().batchUpdate(spreadsheetId=ID, body=body).execute()

# Converts image into spreadsheet color data
def makeImage(name):
  imgRows = []
  image = cv2.imread(name)
  image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB) # opencv is in BGR for some reason

  for j,y in enumerate(image): # for col val (a row) in image
    # print(y[0])
    tempvals = []
    for i,x in enumerate(y): # for row val (a pixel) in col
      tempvals.append({
                  "userEnteredFormat": {
                    "backgroundColor": {
                      "red": int(x[0]) / 255,
                      "green": int(x[1]) / 255,
                      "blue": int(x[2]) / 255
                    }
                  }
                })
    imgRows.append({
              "values": tempvals
            })
  return imgRows

# put in the path to the image file
rows = makeImage("images/fractal.jpg")
# and here set the sheet id (default is 0 for first tab)
updateCellColors(0,rows)

# Just know that if you run this currently it will wipe out the entire sheet so be careful