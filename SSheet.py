# Creste Spreadsheet to folder
def creatSpreadSheet(driveObj,FOLDER_ID,FILE_NAME):
    file_metadata = {
    'name': FILE_NAME,
    'parents': [FOLDER_ID],
    'mimeType': 'application/vnd.google-apps.spreadsheet',
    }
    res = driveObj.files().create(body=file_metadata,supportsAllDrives=True).execute()
    return res
#from pyasn1.type.univ import Null
def getSSValues(sheetObj,SPREADSHEET_ID,SHEET_NAME):
    # The ID and range of a sample spreadsheet.
    # Call the Sheets API
    sheet = sheetObj.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=SHEET_NAME).execute()
    values = result.get('values', [])
    if not values:
        return 'No data found.'
    else:
        return values

def getSSFixValues(sheetObj,SPREADSHEET_ID,SHEET_NAME,RANGE):
    # The ID and range of a sample spreadsheet.
    # Call the Sheets API
    sheet = sheetObj.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=SHEET_NAME+RANGE).execute()
    values = result.get('values',[])
    if not values:
        return 'No data found.'
    else:
        return values[0][0]

# Append SpreadSheet
def appendToSS(sheetObj,SS_ID,sheetName,body,range=""):
    sheet = sheetObj.spreadsheets()
    #result = sheet.values().update(spreadsheetId=SS_ID,
    result = sheet.values().append(spreadsheetId=SS_ID,
                    range=sheetName+range,
                    valueInputOption="USER_ENTERED",
                    body=body
                    ).execute()
    return result

# Write SpreadSheet
def writeToSS(sheetObj,SS_ID,sheetName,body,range=""):
    sheet = sheetObj.spreadsheets()
    #result = sheet.values().update(spreadsheetId=SS_ID,
    result = sheet.values().update(spreadsheetId=SS_ID,
                    range=sheetName+range,
                    valueInputOption="USER_ENTERED",
                    body=body
                    ).execute()
    return result

# Get last Row
def getLastRow(sheetObj,SS_ID,sheetName):
    rows = sheetObj.spreadsheets().values().get(spreadsheetId=SS_ID,range=sheetName+"!A1:A").execute().get('values', [])
    return len(rows)

# Set cell background color
def setCellBackGroundColor(sheetObj,SS_ID,Sheet_ID,row,col,color):
    sheet = sheetObj.spreadsheets()
    body = {
    "requests": [
        {
        "updateCells": {
            "range": {
            "sheetId": Sheet_ID,
            "startRowIndex": row,
            "endRowIndex": row+1,
            "startColumnIndex": col,
            "endColumnIndex": col+1
            },
            "rows": [
            {
                "values": [
                {
                    "userEnteredFormat": {
                    "backgroundColor": {
                        color: 1
                    }
                    }
                }
                ]
            }
            ],
            "fields": "userEnteredFormat.backgroundColor"
        }
        }
    ]
    }
    res = sheet.batchUpdate(spreadsheetId=SS_ID, body=body).execute()
    return res
